# -*- coding: utf-8 -*-
# Author:   Jibril Sulaiman
# Location: save as C:\call-recordings\transcripts\convert.py (README Step 7, optional).
# What:     test version of the Word converter. Reads .txt transcripts from INPUT_FOLDER (either
#           ElevateAI JSON saved with a .txt extension, or plain text), pulls out readable
#           "Speaker N: ..." lines, and writes one .docx per file to "transcript doc".
#           TEST_ONE = True converts a single file so you can check the result first.
#           Results are logged to doc_conversion_log.csv; existing .docx files are skipped.
# Why:      the standard /transcript endpoint returns one word per segment, which is unreadable;
#           this rebuilds lines by speaker and time gaps and gives a document people can open.
# Run:      python convert.py
# Needs:    Python 3.8+, pip install python-docx

import os
import csv
import json
import re
from datetime import datetime
from typing import Any, Dict, List, Tuple, Optional

from docx import Document

# =========================
# CONFIG
# =========================

BASE_PATH = r"C:\call-recordings"   # change to your folder
INPUT_FOLDER = os.path.join(BASE_PATH, "transcripts")                # contains .txt files (JSON or text)
OUTPUT_FOLDER = os.path.join(INPUT_FOLDER, "transcript doc")         # output folder for .docx
LOG_PATH = os.path.join(BASE_PATH, "doc_conversion_log.csv")         # conversion log

# Test with ONE file first
TEST_ONE = True
TEST_FILENAME = None  # e.g. "RE0123456789abcdef0123456789abcdef.txt"; None = first file

# Speaker formatting
SPEAKER_LABELS = {
    "participantOne": "Speaker 1",
    "participantTwo": "Speaker 2",
    "allParticipants": "All",
}

# Aggregation tuning (ms). Helps combine single-word "phrases" into readable lines.
GAP_BREAK_MS = 1500  # if time gap between segments is big, start a new line


# =========================
# HELPERS
# =========================

def safe_mkdir(path: str) -> None:
    os.makedirs(path, exist_ok=True)

def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_log_header() -> None:
    if os.path.exists(LOG_PATH):
        return
    with open(LOG_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["timestamp", "status", "input_txt", "output_docx", "notes"])

def log_row(status: str, input_txt: str, output_docx: str, notes: str = "") -> None:
    ensure_log_header()
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([now_iso(), status, input_txt, output_docx, notes])

def is_probably_json(text: str) -> bool:
    t = text.lstrip()
    return t.startswith("{") or t.startswith("[")

def clean_token(s: str) -> str:
    # Normalize a few common artifacts
    s = s.replace("<UNK>", "").replace("<unk>", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def list_input_txt_files() -> List[str]:
    if not os.path.isdir(INPUT_FOLDER):
        raise FileNotFoundError(f"Input folder not found: {INPUT_FOLDER}")
    return sorted([f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".txt")])

def build_output_path(input_filename: str) -> str:
    base = os.path.splitext(input_filename)[0]
    return os.path.join(OUTPUT_FOLDER, base + ".docx")


# =========================
# TRANSCRIPT EXTRACTORS
# =========================

def extract_from_sentence_segments(data: Dict[str, Any]) -> Optional[str]:
    """
    Handles JSON like:
      { "sentenceSegments": [ {"transcript":"..."}, ... ] }
    """
    segs = data.get("sentenceSegments")
    if not isinstance(segs, list) or not segs:
        return None

    parts: List[str] = []
    for seg in segs:
        if isinstance(seg, dict):
            t = seg.get("transcript")
            if isinstance(t, str):
                t = clean_token(t)
                if t:
                    parts.append(t)

    if not parts:
        return None
    return "\n".join(parts)

def extract_from_phrases_and_segments(data: Dict[str, Any]) -> Optional[str]:
    """
    Handles JSON like the /transcript endpoint returns:
      participantOne.phrases + participantOne.phraseSegments (phraseIndex, startTimeOffset, etc.)
      participantTwo...
      allParticipants...
    We reconstruct time-ordered lines.
    """
    speakers = []
    for key in ("participantOne", "participantTwo"):
        if isinstance(data.get(key), dict):
            speakers.append(key)

    # If it doesn't look like that structure, bail
    if not speakers:
        return None

    # Build a list of (start_ms, speaker_key, phrase_text)
    events: List[Tuple[int, str, str]] = []

    for spk_key in speakers:
        obj = data.get(spk_key, {})
        phrases = obj.get("phrases", [])
        segments = obj.get("phraseSegments", [])

        if not isinstance(phrases, list) or not isinstance(segments, list):
            continue

        for seg in segments:
            if not isinstance(seg, dict):
                continue
            start = seg.get("startTimeOffset")
            idx = seg.get("phraseIndex")

            if not isinstance(start, int) or not isinstance(idx, int):
                continue
            if idx < 0 or idx >= len(phrases):
                continue

            phrase = phrases[idx]
            if not isinstance(phrase, str):
                continue

            phrase = clean_token(phrase)
            if phrase:
                events.append((start, spk_key, phrase))

    if not events:
        return None

    events.sort(key=lambda x: x[0])

    # Aggregate into readable lines
    lines: List[str] = []
    cur_speaker: Optional[str] = None
    cur_words: List[str] = []
    last_time: Optional[int] = None

    def flush():
        nonlocal cur_speaker, cur_words
        if cur_speaker and cur_words:
            label = SPEAKER_LABELS.get(cur_speaker, cur_speaker)
            lines.append(f"{label}: " + " ".join(cur_words).strip())
        cur_speaker = None
        cur_words = []

    for start_ms, spk_key, phrase in events:
        if cur_speaker is None:
            cur_speaker = spk_key
            cur_words = [phrase]
            last_time = start_ms
            continue

        gap = 0 if last_time is None else (start_ms - last_time)
        speaker_changed = (spk_key != cur_speaker)
        big_gap = gap > GAP_BREAK_MS

        if speaker_changed or big_gap:
            flush()
            cur_speaker = spk_key
            cur_words = [phrase]
        else:
            cur_words.append(phrase)

        last_time = start_ms

    flush()

    if not lines:
        return None

    return "\n".join(lines)

def extract_plain_text_from_file(filepath: str) -> str:
    """
    Accepts:
      - JSON text (stored in .txt)
      - already-plain transcript text
    Returns best-effort plain transcript text.
    """
    raw = ""
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        raw = f.read()

    raw = raw.strip()
    if not raw:
        return ""

    # If it isn't JSON, assume it's already a transcript
    if not is_probably_json(raw):
        return raw

    # JSON parse
    try:
        data = json.loads(raw)
    except Exception:
        # Sometimes it's almost-json; if it fails, just return raw as fallback
        return raw

    if isinstance(data, dict):
        # 1) Preferred: sentenceSegments (yields real transcript text)
        t1 = extract_from_sentence_segments(data)
        if t1:
            return t1

        # 2) The other structure: phrases + phraseSegments
        t2 = extract_from_phrases_and_segments(data)
        if t2:
            return t2

        # 3) Common fallback fields
        for key in ("transcript", "text", "result", "output"):
            v = data.get(key)
            if isinstance(v, str) and v.strip():
                return v.strip()

        # 4) If nothing else, at least return a compact JSON string (should be rare)
        return json.dumps(data, indent=2)

    # If it's a list, try to join any transcript-like strings
    if isinstance(data, list):
        parts = []
        for item in data:
            if isinstance(item, dict):
                for k in ("transcript", "text"):
                    v = item.get(k)
                    if isinstance(v, str) and v.strip():
                        parts.append(v.strip())
        if parts:
            return "\n".join(parts)
        return json.dumps(data, indent=2)

    return raw


# =========================
# DOCX WRITER
# =========================

def write_docx(output_path: str, transcript_text: str) -> None:
    doc = Document()
    # Split into paragraphs for cleaner Word formatting
    for line in transcript_text.splitlines():
        line = line.rstrip()
        if line:
            doc.add_paragraph(line)
        else:
            doc.add_paragraph("")  # blank line
    doc.save(output_path)


# =========================
# MAIN
# =========================

def main():
    safe_mkdir(OUTPUT_FOLDER)
    ensure_log_header()

    files = list_input_txt_files()

    if TEST_ONE:
        if TEST_FILENAME and TEST_FILENAME in files:
            files = [TEST_FILENAME]
        else:
            # fallback to first file found
            files = files[:1]

    if not files:
        print("No .txt files found in:", INPUT_FOLDER)
        return

    for fname in files:
        in_path = os.path.join(INPUT_FOLDER, fname)
        out_path = build_output_path(fname)

        # Skip if already converted
        if os.path.exists(out_path):
            print("SKIP (exists):", out_path)
            log_row("skip_exists", in_path, out_path, "Output already exists")
            continue

        try:
            text = extract_plain_text_from_file(in_path)
            text = (text or "").strip()

            if not text or len(text) < 10:
                msg = "Transcript extraction returned empty/too short text"
                print("ERROR:", fname, "-", msg)
                log_row("error", in_path, out_path, msg)
                continue

            write_docx(out_path, text)
            print("OK:", fname, "->", out_path)
            log_row("success", in_path, out_path, "")

        except Exception as e:
            print("ERROR:", fname, "-", str(e))
            log_row("error", in_path, out_path, str(e))


if __name__ == "__main__":
    main()
