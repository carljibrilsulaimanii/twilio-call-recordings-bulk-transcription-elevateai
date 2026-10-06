# -*- coding: utf-8 -*-
# Author:   Jibril Sulaiman
# Location: save in the folder that holds your .txt transcripts (README Step 7, optional);
#           it works on the folder you run it from.
# What:     full-run version of convert.py: converts every .txt in the current folder (ElevateAI
#           JSON saved as .txt, or plain text) to a .docx in a "transcript doc" subfolder, logs
#           to doc_conversion_log.csv, and skips files already converted.
# Why:      same as convert.py, but with no hard-coded paths, so it can be dropped into any
#           transcripts folder and run there.
# Run:      cd "C:\call-recordings\transcripts"   then   python convert_json_txt_to_docx.py
# Needs:    Python 3.8+, pip install python-docx

import os
import csv
import json
import re
from datetime import datetime
from docx import Document

# ==========================================================
# CONFIG — SCRIPT IS RUN FROM THE TRANSCRIPTS FOLDER
# ==========================================================

BASE_DIR = os.getcwd()  # e.g. C:\call-recordings\transcripts
OUTPUT_DIR = os.path.join(BASE_DIR, "transcript doc")
LOG_FILE = os.path.join(BASE_DIR, "doc_conversion_log.csv")

# ---- TEST MODE ----
TEST_ONE_FILE = False
TEST_FILENAME = None  # e.g. "RE0123456789abcdef0123456789abcdef.txt"; None = first file

# Gap threshold (ms) to break sentences when reconstructing phrases
GAP_BREAK_MS = 1500

SPEAKER_MAP = {
    "participantOne": "Speaker 1",
    "participantTwo": "Speaker 2",
    "allParticipants": "All"
}

# ==========================================================
# HELPERS
# ==========================================================

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_log():
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "status", "input_file", "output_file", "notes"])

def log(status, input_file, output_file, notes=""):
    ensure_log()
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([timestamp(), status, input_file, output_file, notes])

def is_json(text):
    return text.lstrip().startswith("{") or text.lstrip().startswith("[")

def clean_text(t):
    t = t.replace("<UNK>", "").replace("<unk>", "")
    t = re.sub(r"\s+", " ", t)
    return t.strip()

# ==========================================================
# TRANSCRIPT EXTRACTION
# ==========================================================

def extract_sentence_segments(data):
    segs = data.get("sentenceSegments")
    if not isinstance(segs, list):
        return None

    lines = []
    for s in segs:
        if isinstance(s, dict):
            t = clean_text(s.get("transcript", ""))
            if t:
                lines.append(t)

    return "\n".join(lines) if lines else None

def extract_phrases(data):
    events = []

    for speaker_key in ("participantOne", "participantTwo"):
        speaker = data.get(speaker_key)
        if not isinstance(speaker, dict):
            continue

        phrases = speaker.get("phrases", [])
        segments = speaker.get("phraseSegments", [])

        if not isinstance(phrases, list) or not isinstance(segments, list):
            continue

        for seg in segments:
            if not isinstance(seg, dict):
                continue

            idx = seg.get("phraseIndex")
            start = seg.get("startTimeOffset")

            if isinstance(idx, int) and isinstance(start, int) and idx < len(phrases):
                text = clean_text(phrases[idx])
                if text:
                    events.append((start, speaker_key, text))

    if not events:
        return None

    events.sort(key=lambda x: x[0])

    output = []
    current_speaker = None
    buffer = []
    last_time = None

    def flush():
        if current_speaker and buffer:
            label = SPEAKER_MAP.get(current_speaker, current_speaker)
            output.append(f"{label}: " + " ".join(buffer))

    for start, speaker, text in events:
        if speaker != current_speaker or (last_time and start - last_time > GAP_BREAK_MS):
            flush()
            buffer = [text]
            current_speaker = speaker
        else:
            buffer.append(text)

        last_time = start

    flush()
    return "\n".join(output) if output else None

def extract_plain_text(filepath):
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        raw = f.read().strip()

    if not raw:
        return ""

    if not is_json(raw):
        return raw  # already plain text

    try:
        data = json.loads(raw)
    except Exception:
        return raw

    if isinstance(data, dict):
        return (
            extract_sentence_segments(data)
            or extract_phrases(data)
            or clean_text(data.get("transcript", ""))
            or json.dumps(data, indent=2)
        )

    return raw

# ==========================================================
# DOCX WRITER
# ==========================================================

def write_docx(path, text):
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    doc.save(path)

# ==========================================================
# MAIN
# ==========================================================

def main():
    ensure_dir(OUTPUT_DIR)
    ensure_log()

    txt_files = [f for f in os.listdir(BASE_DIR) if f.lower().endswith(".txt")]

    if TEST_ONE_FILE:
        if TEST_FILENAME and TEST_FILENAME in txt_files:
            txt_files = [TEST_FILENAME]
        else:
            txt_files = txt_files[:1]

    if not txt_files:
        print("No .txt files found.")
        return

    for filename in txt_files:
        input_path = os.path.join(BASE_DIR, filename)
        output_path = os.path.join(OUTPUT_DIR, os.path.splitext(filename)[0] + ".docx")

        if os.path.exists(output_path):
            log("skipped", filename, output_path, "Already exists")
            continue

        try:
            text = extract_plain_text(input_path)
            if len(text.strip()) < 20:
                raise ValueError("Transcript empty or too short")

            write_docx(output_path, text)
            log("success", filename, output_path)
            print(f"✔ Converted: {filename}")

        except Exception as e:
            log("error", filename, output_path, str(e))
            print(f"✖ Error: {filename} — {e}")

if __name__ == "__main__":
    main()
