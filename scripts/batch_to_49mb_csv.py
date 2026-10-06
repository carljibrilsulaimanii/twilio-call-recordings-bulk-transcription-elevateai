# -*- coding: utf-8 -*-
# Author:   Jibril Sulaiman
# Location: save as C:\call-recordings\transcripts\batch_to_49mb_csv.py (README Step 6).
# What:     combines every .txt transcript in INPUT_DIR into CSV files with two columns,
#           file_id (the file name without .txt) and conversation (the whole transcript), one
#           row per call. A new transcripts_batch_NNNN.csv starts before a file would pass 49 MB.
#           csv_processed_log.csv records every file batched, so a re-run only adds new files.
# Why:      GoHighLevel's Knowledge Base "Table Upload" takes CSV files only, max 50 MB each.
#           Thousands of separate .txt files can't be uploaded that way; a few CSVs can.
# Run:      python batch_to_49mb_csv.py
# Needs:    Python 3.8+ (standard library only)

import os
import csv
import io
import time
from datetime import datetime

# -------------------------
# CONFIG
# -------------------------

# Change this to whatever folder you want to batch
INPUT_DIR = r"C:\call-recordings\transcripts\punctuated_txt"

OUTPUT_DIR = r"C:\call-recordings\transcripts\csv_batches_49mb"
PROCESSED_LOG = os.path.join(OUTPUT_DIR, "csv_processed_log.csv")

# Target max size per CSV batch (49MB)
MAX_BYTES_PER_BATCH = 49 * 1024 * 1024

# Only include .txt files
EXTENSIONS = (".txt",)

# Sort options: "alpha" or "mtime"
SORT_MODE = "alpha"

# Optional: set to an integer to test with first N files (e.g., 2). Set to None for all.
TEST_LIMIT = None  # e.g. 2


# -------------------------
# HELPERS
# -------------------------

def safe_mkdir(path: str):
    if not os.path.exists(path):
        os.makedirs(path)

def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def load_processed_set():
    """Returns set of filenames that have already been included in a batch."""
    processed = set()
    if not os.path.exists(PROCESSED_LOG):
        return processed

    try:
        with open(PROCESSED_LOG, "r", encoding="utf-8", newline="") as f:
            r = csv.DictReader(f)
            for row in r:
                fn = (row.get("filename") or "").strip()
                if fn:
                    processed.add(fn)
    except:
        # If log is malformed, don't block processing; start fresh behavior.
        return processed

    return processed

def ensure_processed_log_header():
    if os.path.exists(PROCESSED_LOG):
        return
    with open(PROCESSED_LOG, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["timestamp", "batch_file", "filename", "bytes"])

def log_processed(batch_file: str, filename: str, byte_count: int):
    ensure_processed_log_header()
    with open(PROCESSED_LOG, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([now_iso(), batch_file, filename, byte_count])

def list_input_files():
    files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(EXTENSIONS)]

    if SORT_MODE == "mtime":
        files.sort(key=lambda x: os.path.getmtime(os.path.join(INPUT_DIR, x)))
        return files

    # default alpha
    files.sort(key=lambda x: x.lower())
    return files

def read_text_file(path: str) -> str:
    # handle weird encoding gracefully
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()

def estimate_row_bytes(file_id: str, conversation: str) -> int:
    """
    Estimate the bytes the row will take in the CSV file by actually
    generating the CSV row in-memory (so quoting is accounted for).
    """
    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_MINIMAL)
    w.writerow([file_id, conversation])
    return len(buf.getvalue().encode("utf-8"))

def make_batch_filename(batch_index: int) -> str:
    return os.path.join(OUTPUT_DIR, f"transcripts_batch_{batch_index:04d}.csv")


# -------------------------
# MAIN
# -------------------------

def main():
    safe_mkdir(OUTPUT_DIR)

    if not os.path.exists(INPUT_DIR):
        raise Exception(f"INPUT_DIR not found: {INPUT_DIR}")

    processed = load_processed_set()
    files = list_input_files()

    # filter already processed
    files_to_process = [f for f in files if f not in processed]

    if TEST_LIMIT is not None:
        files_to_process = files_to_process[:TEST_LIMIT]

    print(f"Input dir: {INPUT_DIR}")
    print(f"Output dir: {OUTPUT_DIR}")
    print(f"Total files found: {len(files)}")
    print(f"Already processed (per log): {len(processed)}")
    print(f"Remaining to batch: {len(files_to_process)}")
    print(f"Max per batch: {MAX_BYTES_PER_BATCH} bytes (~49MB)")
    print(f"Sort mode: {SORT_MODE}")
    if TEST_LIMIT is not None:
        print(f"TEST_LIMIT enabled: {TEST_LIMIT}")

    if not files_to_process:
        print("Nothing to do.")
        return

    batch_index = 1

    # Find next available batch filename (don’t overwrite)
    while os.path.exists(make_batch_filename(batch_index)):
        batch_index += 1

    current_batch_path = make_batch_filename(batch_index)
    current_bytes = 0
    current_rows = 0

    # open first batch
    out_f = open(current_batch_path, "w", encoding="utf-8", newline="")
    writer = csv.writer(out_f, quoting=csv.QUOTE_MINIMAL)
    writer.writerow(["file_id", "conversation"])  # header
    current_bytes += len(("file_id,conversation\n").encode("utf-8"))

    print(f"\nStarting batch: {os.path.basename(current_batch_path)}")

    for i, filename in enumerate(files_to_process, start=1):
        full_path = os.path.join(INPUT_DIR, filename)
        file_id = os.path.splitext(filename)[0]

        conversation = read_text_file(full_path)

        # If you want to strip leading/trailing whitespace, uncomment:
        # conversation = conversation.strip()

        row_bytes = estimate_row_bytes(file_id, conversation)

        # If adding this row would exceed the batch limit, start a new batch
        if current_rows > 0 and (current_bytes + row_bytes) > MAX_BYTES_PER_BATCH:
            out_f.close()
            print(f"Closed batch: {os.path.basename(current_batch_path)} | rows: {current_rows} | approx bytes: {current_bytes}")

            batch_index += 1
            current_batch_path = make_batch_filename(batch_index)
            out_f = open(current_batch_path, "w", encoding="utf-8", newline="")
            writer = csv.writer(out_f, quoting=csv.QUOTE_MINIMAL)
            writer.writerow(["file_id", "conversation"])
            current_bytes = len(("file_id,conversation\n").encode("utf-8"))
            current_rows = 0

            print(f"\nStarting batch: {os.path.basename(current_batch_path)}")

        # write row
        writer.writerow([file_id, conversation])
        current_bytes += row_bytes
        current_rows += 1

        # log so we never re-process this file
        log_processed(os.path.basename(current_batch_path), filename, row_bytes)

        if i % 50 == 0:
            print(f"Processed {i}/{len(files_to_process)}...")

    out_f.close()
    print(f"\nDone.")
    print(f"Last batch: {os.path.basename(current_batch_path)} | rows: {current_rows} | approx bytes: {current_bytes}")
    print(f"Log: {PROCESSED_LOG}")


if __name__ == "__main__":
    main()
