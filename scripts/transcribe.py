# -*- coding: utf-8 -*-
# Author:   Jibril Sulaiman
# Location: save as C:\call-recordings\transcribe.py (README Step 4); run from that folder.
# What:     sends every .mp3 in C:\call-recordings\recordings to the ElevateAI API one at a time:
#           declare an interaction -> upload the MP3 -> poll the punctuated transcript endpoint
#           until it is ready -> save it as JSON (punctuated\{interaction_id}.punctuated.json)
#           and as plain text (punctuated_txt\{interaction_id}.txt, one "Speaker N: sentence" per line).
#           Every step is written to processed_log.csv, and MP3s logged as "success" are skipped
#           on the next run, so you can stop with Ctrl+C and restart safely.
# Why:      ElevateAI's dashboard takes files one by one and its status endpoint never says
#           "processed", so a bulk folder of call recordings needs a script that waits on the
#           transcript itself and keeps a log of what is done.
# Run:      $env:ELEVATEAI_API_TOKEN = "<your token>"   then   python transcribe.py
# Needs:    Python 3.8+, pip install requests

import os
import csv
import json
import time
import re
import requests
from datetime import datetime

# -------------------
# CONFIG
# -------------------

# Processing order options:
#   "alpha"     = alphabetical by filename
#   "size_asc"  = smallest files first (fastest wins)
#   "size_desc" = largest files first
SORT_MODE = "size_asc"

# Your ElevateAI API token, read from an environment variable (README Step 3).
# Never paste the token into this file.
API_TOKEN = os.environ.get("ELEVATEAI_API_TOKEN", "").strip()
if not API_TOKEN:
    raise SystemExit("Set the ELEVATEAI_API_TOKEN environment variable first (README Step 3).")

BASE_PATH = r"C:\call-recordings"   # change to your folder
AUDIO_FOLDER = os.path.join(BASE_PATH, "recordings")

TRANSCRIPTS_BASE = os.path.join(BASE_PATH, "transcripts")
PUNCT_JSON_DIR = os.path.join(TRANSCRIPTS_BASE, "punctuated")
PUNCT_TXT_DIR = os.path.join(TRANSCRIPTS_BASE, "punctuated_txt")

# New log (rebuilt going forward)
LOG_PATH = os.path.join(BASE_PATH, "processed_log.csv")

# Endpoints
DECLARE_URL = "https://api.elevateai.com/v1/interactions"
UPLOAD_URL_TMPL = "https://api.elevateai.com/v1/interactions/{interaction_id}/upload"
PUNCT_URL_TMPL = "https://api.elevateai.com/v1/interactions/{interaction_id}/transcripts/punctuated"

HEADERS_JSON = {
    "Content-Type": "application/json",
    "X-API-Token": API_TOKEN,
}

HEADERS = {
    "X-API-Token": API_TOKEN,
    "Accept-Encoding": "gzip, deflate, br",
}

REQUEST_TIMEOUT = 120

# -------------------
# HELPERS
# -------------------

def safe_mkdir(path: str):
    os.makedirs(path, exist_ok=True)

def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_log_header():
    if os.path.exists(LOG_PATH):
        return
    with open(LOG_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
            "timestamp",
            "status",
            "mp3_filename",
            "interaction_id",
            "size_mb",
            "punctuated_json_path",
            "punctuated_txt_path",
            "notes",
        ])

def log_row(status: str, mp3_filename: str, interaction_id: str, size_mb: float,
            json_path: str = "", txt_path: str = "", notes: str = ""):
    ensure_log_header()
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([now_iso(), status, mp3_filename, interaction_id, size_mb, json_path, txt_path, notes])

def list_mp3_files(folder: str):
    files = [f for f in os.listdir(folder) if f.lower().endswith(".mp3")]

    if SORT_MODE == "alpha":
        return sorted(files, key=lambda x: x.lower())

    if SORT_MODE == "size_asc":
        return sorted(files, key=lambda x: os.path.getsize(os.path.join(folder, x)))

    if SORT_MODE == "size_desc":
        return sorted(files, key=lambda x: os.path.getsize(os.path.join(folder, x)), reverse=True)

    return sorted(files, key=lambda x: x.lower())

def load_processed_mp3s_from_log() -> set:
    """
    If you cleared the log, this set will be empty.
    Going forward, this prevents reprocessing the same MP3.
    """
    done = set()
    if not os.path.exists(LOG_PATH):
        return done

    try:
        with open(LOG_PATH, "r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if not row:
                    continue
                if (row.get("status") or "").strip().lower() == "success":
                    mp3 = (row.get("mp3_filename") or "").strip()
                    if mp3:
                        done.add(mp3.lower())
    except Exception:
        return done

    return done

def participant_label(p: str) -> str:
    p = (p or "").strip()
    if p.lower() == "participantone":
        return "Speaker 1"
    if p.lower() == "participanttwo":
        return "Speaker 2"
    return p or "Speaker"

def build_plain_text_from_punctuated(data: dict) -> str:
    """
    Plain text output: speaker + punctuated sentence.
    Example:
      Speaker 1: Thank you for calling...
      Speaker 2: Hi, I need help...
    """
    segs = data.get("sentenceSegments") or []
    lines = []
    for seg in segs:
        if not isinstance(seg, dict):
            continue
        speaker = participant_label(seg.get("participant"))
        phrase = (seg.get("phrase") or "").strip()
        if not phrase:
            continue
        lines.append(f"{speaker}: {phrase}")
    return "\n".join(lines).strip()

# -------------------
# API STEPS
# -------------------

def declare_interaction() -> str:
    payload = {
        "type": "audio",
        "languageTag": "en-us",
        "vertical": "default",
    }

    r = requests.post(DECLARE_URL, headers=HEADERS_JSON, json=payload, timeout=60)
    if r.status_code not in (200, 201):
        raise Exception(f"Declare failed {r.status_code}: {r.text}")

    data = r.json()
    interaction_id = data.get("interactionIdentifier")
    if not interaction_id:
        raise Exception(f"Declare response missing interactionIdentifier: {data}")

    return interaction_id

def upload_mp3(interaction_id: str, file_path: str):
    url = UPLOAD_URL_TMPL.format(interaction_id=interaction_id)

    # "filename mp3" is the field name that worked; "file" is the fallback.
    last_err = ""
    for field_name in ("filename mp3", "file"):
        with open(file_path, "rb") as f:
            files = {field_name: (os.path.basename(file_path), f, "audio/mpeg")}
            r = requests.post(url, headers={"X-API-Token": API_TOKEN}, files=files, timeout=300)

        if r.status_code in (200, 201, 202):
            return

        last_err = f"Upload failed using field '{field_name}' ({r.status_code}): {r.text}"

    raise Exception(last_err)

def wait_for_punctuated(interaction_id: str, file_size_mb: float) -> dict:
    """
    Poll punctuated endpoint until ready.
    - 404 is expected while transcript is still processing
    - use size-based timeout so larger files get more time
    """
    url = PUNCT_URL_TMPL.format(interaction_id=interaction_id)

    # Size-based waiting window (tune as needed)
    # Example: 0.5MB -> ~20 minutes, 2.5MB -> ~35 minutes, 10MB -> capped
    max_minutes = max(20, min(180, int(20 + file_size_mb * 6)))
    sleep_seconds = 10
    attempts = int((max_minutes * 60) / sleep_seconds)

    for i in range(attempts):
        r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)

        if r.status_code == 404:
            print(f"   Punctuated not ready (404). Check {i+1}/{attempts}")
            time.sleep(sleep_seconds)
            continue

        if r.status_code in (401, 403):
            raise Exception(f"Punctuated auth error {r.status_code}: {r.text}")

        if r.status_code != 200:
            print(f"   Punctuated HTTP {r.status_code}. Check {i+1}/{attempts}")
            time.sleep(sleep_seconds)
            continue

        data = r.json()
        # Must have sentenceSegments to be useful
        segs = data.get("sentenceSegments")
        if not isinstance(segs, list) or len(segs) == 0:
            print(f"   Punctuated returned empty sentenceSegments. Check {i+1}/{attempts}")
            time.sleep(sleep_seconds)
            continue

        return data

    raise Exception(f"Punctuated transcript not available before timeout ({max_minutes} min)")

# -------------------
# MAIN
# -------------------

def main():
    safe_mkdir(TRANSCRIPTS_BASE)
    safe_mkdir(PUNCT_JSON_DIR)
    safe_mkdir(PUNCT_TXT_DIR)
    ensure_log_header()

    mp3s = list_mp3_files(AUDIO_FOLDER)
    processed_mp3s = load_processed_mp3s_from_log()

    print(f"Total files found: {len(mp3s)}")
    print(f"Sorting by: {SORT_MODE}")
    print(f"Already processed (from log): {len(processed_mp3s)}")

    for filename in mp3s:
        # Skip if already processed (based on rebuilt log going forward)
        if filename.lower() in processed_mp3s:
            continue

        file_path = os.path.join(AUDIO_FOLDER, filename)
        size_mb = round(os.path.getsize(file_path) / (1024 * 1024), 2)

        print(f"\nProcessing: {filename} ({size_mb} MB)")

        interaction_id = ""
        json_out = ""
        txt_out = ""

        try:
            interaction_id = declare_interaction()
            print(f"   Declared: {interaction_id}")

            upload_mp3(interaction_id, file_path)
            print("   Upload OK.")
            log_row("uploaded", filename, interaction_id, size_mb)

            # Wait for punctuated transcript
            print("   Waiting for punctuated transcript...")
            punct_data = wait_for_punctuated(interaction_id, size_mb)

            # Save outputs by interaction_id
            json_out = os.path.join(PUNCT_JSON_DIR, f"{interaction_id}.punctuated.json")
            txt_out = os.path.join(PUNCT_TXT_DIR, f"{interaction_id}.txt")

            with open(json_out, "w", encoding="utf-8") as f:
                json.dump(punct_data, f, ensure_ascii=False, indent=2)

            txt = build_plain_text_from_punctuated(punct_data)
            with open(txt_out, "w", encoding="utf-8") as f:
                f.write(txt)

            print(f"   Saved JSON: {json_out}")
            print(f"   Saved TXT : {txt_out}")

            log_row("success", filename, interaction_id, size_mb, json_out, txt_out)

        except Exception as e:
            print(f"   ERROR: {e}")
            log_row("error", filename, interaction_id, size_mb, json_out, txt_out, notes=str(e))
            print("   Continuing...")
            time.sleep(3)

if __name__ == "__main__":
    main()
