# -*- coding: utf-8 -*-
# Author:   Jibril Sulaiman
# Location: save as C:\call-recordings\transcripts\download_punctuated.py (README Step 5);
#           run from that folder.
# What:     re-downloads punctuated transcripts for interactions that were already uploaded.
#           It scans a log CSV for every interaction id (any cell that looks like a GUID, so
#           missing or odd headers don't matter), calls the punctuated transcript endpoint for
#           each, and saves punctuated\{id}.punctuated.json plus punctuated_txt\{id}.txt
#           ("Speaker N [mm:ss-mm:ss]: sentence"). Already-downloaded ids are skipped and every
#           result goes to punctuated_download_log.csv.
# Why:      early runs saved the word-by-word /transcript output (one word per line) instead of
#           punctuated sentences, and some runs died before saving. This recovers the readable
#           version without uploading (and paying for) the audio again.
# Run:      $env:ELEVATEAI_API_TOKEN = "<your token>"   then   python .\download_punctuated.py
# Needs:    Python 3.8+, pip install requests

import os
import csv
import re
import time
import json
import requests
from datetime import datetime

# -------------------
# CONFIG
# -------------------
# Your ElevateAI API token, read from an environment variable (README Step 3).
# Never paste the token into this file.
API_TOKEN = os.environ.get("ELEVATEAI_API_TOKEN", "").strip()
if not API_TOKEN:
    raise SystemExit("Set the ELEVATEAI_API_TOKEN environment variable first (README Step 3).")

BASE_PATH = r"C:\call-recordings"   # change to your folder
TRANSCRIPTS_DIR = os.path.join(BASE_PATH, "transcripts")

# Try these locations automatically
CANDIDATE_LOGS = [
    os.path.join(BASE_PATH, "processed_log1.csv"),
    os.path.join(TRANSCRIPTS_DIR, "processed_log.csv"),
]

# Output folders
OUT_JSON_DIR = os.path.join(TRANSCRIPTS_DIR, "punctuated")
OUT_TXT_DIR = os.path.join(TRANSCRIPTS_DIR, "punctuated_txt")

# Download log (prevents duplicates + shows failures)
OUT_LOG = os.path.join(TRANSCRIPTS_DIR, "punctuated_download_log.csv")

# Endpoint
PUNCT_URL_TMPL = "https://api.elevateai.com/v1/interactions/{interaction_id}/transcripts/punctuated"

HEADERS = {
    "X-API-Token": API_TOKEN,
    "Accept-Encoding": "gzip, deflate, br",
}

# Retry behavior (404 usually means “not ready yet”)
MAX_RETRIES_404 = 180
SLEEP_SECONDS = 10
REQUEST_TIMEOUT = 120

# UUID pattern (interactionIdentifier looks like this)
UUID_RE = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
)

# -------------------
# HELPERS
# -------------------
def now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def ensure_out_log_header():
    if os.path.exists(OUT_LOG):
        return
    with open(OUT_LOG, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["timestamp", "status", "interaction_id", "json_file", "txt_file", "notes"])

def log_out(status: str, interaction_id: str, json_file: str = "", txt_file: str = "", notes: str = ""):
    ensure_out_log_header()
    with open(OUT_LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([now_ts(), status, interaction_id, json_file, txt_file, notes])

def file_nonempty(path: str) -> bool:
    return os.path.exists(path) and os.path.getsize(path) > 10

def find_processed_log() -> str:
    for p in CANDIDATE_LOGS:
        if os.path.exists(p):
            return p
    raise FileNotFoundError(
        "Could not find processed_log.csv. Looked in:\n" + "\n".join(CANDIDATE_LOGS)
    )

def extract_interaction_ids_from_csv(csv_path: str):
    """
    Reads the CSV as raw rows and extracts UUIDs from ANY cell.
    Works even if headers are missing or column names differ.
    """
    found = []
    seen = set()

    with open(csv_path, "r", encoding="utf-8", errors="ignore", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row:
                continue
            line = " ".join(str(x) for x in row if x is not None)
            for m in UUID_RE.findall(line):
                mid = m.lower()
                if mid not in seen:
                    seen.add(mid)
                    found.append(m)
    return found

def download_punctuated(interaction_id: str) -> dict:
    url = PUNCT_URL_TMPL.format(interaction_id=interaction_id)

    for attempt in range(1, MAX_RETRIES_404 + 1):
        r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)

        if r.status_code == 404:
            print(f"[{interaction_id}] 404 not ready (attempt {attempt}/{MAX_RETRIES_404})")
            time.sleep(SLEEP_SECONDS)
            continue

        if r.status_code in (401, 403):
            raise Exception(f"Auth error {r.status_code}: {r.text}")

        if r.status_code != 200:
            raise Exception(f"HTTP {r.status_code}: {r.text[:500]}")

        return r.json()

    raise Exception(f"Timed out waiting for punctuated transcript (404 after {MAX_RETRIES_404} tries)")

def participant_label(p: str) -> str:
    p = (p or "").strip()
    if p.lower() == "participantone":
        return "Speaker 1"
    if p.lower() == "participanttwo":
        return "Speaker 2"
    return p or "Speaker"

def ms_to_mmss(ms: int) -> str:
    ms = ms or 0
    total_seconds = int(ms) // 1000
    m = total_seconds // 60
    s = total_seconds % 60
    return f"{m:02d}:{s:02d}"

def build_plain_text(data: dict) -> str:
    segs = data.get("sentenceSegments") or []
    lines = []
    for seg in segs:
        if not isinstance(seg, dict):
            continue
        speaker = participant_label(seg.get("participant"))
        phrase = (seg.get("phrase") or "").strip()
        if not phrase:
            continue
        start = ms_to_mmss(seg.get("startTimeOffset"))
        end = ms_to_mmss(seg.get("endTimeOffset"))
        lines.append(f"{speaker} [{start}-{end}]: {phrase}")
    return "\n".join(lines).strip()

# -------------------
# MAIN
# -------------------
def main():
    ensure_dir(OUT_JSON_DIR)
    ensure_dir(OUT_TXT_DIR)
    ensure_out_log_header()

    log_path = find_processed_log()
    interaction_ids = extract_interaction_ids_from_csv(log_path)

    print(f"Using log: {log_path}")
    print(f"Found {len(interaction_ids)} unique interaction_id(s) in processed_log.csv")
    print(f"Saving JSON to: {OUT_JSON_DIR}")
    print(f"Saving TXT  to: {OUT_TXT_DIR}")

    if not interaction_ids:
        print("\nNo interaction IDs found in the CSV.")
        print("That means your processed_log.csv likely does not contain any GUIDs (interactionIdentifier).")
        print("Open processed_log.csv and confirm the interaction IDs are present somewhere.")
        return

    for iid in interaction_ids:
        out_json = os.path.join(OUT_JSON_DIR, f"{iid}.punctuated.json")
        out_txt = os.path.join(OUT_TXT_DIR, f"{iid}.txt")

        if file_nonempty(out_json) and file_nonempty(out_txt):
            print(f"[{iid}] SKIP already downloaded")
            log_out("skip_exists", iid, out_json, out_txt, "already downloaded")
            continue

        try:
            print(f"[{iid}] Downloading punctuated transcript...")
            data = download_punctuated(iid)

            with open(out_json, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            txt = build_plain_text(data)
            with open(out_txt, "w", encoding="utf-8") as f:
                f.write(txt)

            print(f"[{iid}] SAVED -> JSON and TXT")
            log_out("success", iid, out_json, out_txt, "")

        except Exception as e:
            print(f"[{iid}] ERROR -> {e}")
            log_out("error", iid, out_json, out_txt, str(e))

    print("Done.")

if __name__ == "__main__":
    main()
