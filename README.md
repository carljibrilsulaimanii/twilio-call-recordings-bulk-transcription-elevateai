<!--
  README.md -- Twilio call recordings: bulk transcription with ElevateAI, batched into CSV
  Author:  Jibril Sulaiman
  Date:    2026-10-06
  What:    Click-by-click guide for sending a folder of Twilio call recordings (MP3) to the
           ElevateAI API one by one (declare -> upload -> poll -> download), saving readable
           punctuated transcripts, optionally turning them into Word files, and packing them
           into CSV files under 49 MB for a GoHighLevel chatbot knowledge base.
  Why:     The ElevateAI dashboard takes files one at a time, its status endpoint never says
           "processed", the plain transcript comes back one word per line, and the GoHighLevel
           knowledge base only takes CSV tables up to 50 MB. Each script here fixes one of those.
-->

# Twilio call recordings: bulk transcription with ElevateAI, batched into CSV

Turn a folder of thousands of Twilio call recordings into readable, speaker-labelled
transcripts, then pack them into a few CSV files small enough to upload to a GoHighLevel
(GHL) AI chatbot's knowledge base. Five small Python scripts do the work on your own
computer. You don't need a server, a cloud bucket or a public link.

> ⚠️ **ElevateAI has been retired.** Its sign-in page announced that it would stop
> accepting new data on March 15, 2026, and shut down after March 31, 2026, deleting any
> data left on it.
> [`scripts/transcribe.py`](scripts/transcribe.py) and
> [`scripts/download_punctuated.py`](scripts/download_punctuated.py) call that API, so they
> won't run any more. They're kept here as a worked example of the pattern most
> speech-to-text APIs use: declare a job, upload audio, poll, download. The other three
> scripts don't call any API. They work on any folder of `.txt` transcripts.

## Why it exists

The goal was to train a GHL Conversation AI chatbot on how a business actually talks to
its customers. The raw material was a couple of thousand Twilio call recordings, downloaded
as MP3 files named by their Recording SID (`RE` followed by 32 characters). Every tool
along the way had a limit that blocked a simple "upload the folder" approach:

1. **ElevateAI's dashboard takes files by hand.** Its **Interactions** page has an
   **Upload files** button, which is fine for a handful of calls and useless for thousands.
2. **The API won't take an MP3 on the first call.** `POST /v1/interactions` only
   *declares* a job. Sending the audio there returns **HTTP 415** (Unsupported Media Type).
   The audio goes in a second call to `/v1/interactions/{id}/upload`. (The first chatbot
   suggestion was to expose the folder through ngrok so ElevateAI could fetch each file
   by URL. You don't need that. The upload endpoint takes the file directly.)
3. **The status endpoint never says "processed".** `GET /v1/interactions/{id}` returns
   the job's settings (`identifier`, `type`, `languageTag`, `vertical`, `model`, ...) but
   no `status` field. A script that waits for `status == "processed"` waits forever, even
   while the dashboard shows **Processed**. The only reliable signal is the transcript
   endpoint itself: it returns **404** until the transcript is ready, then **200**.
4. **The plain transcript is one word per line.** The `/transcript` endpoint returns
   one segment per word. Saved as text, a call reads `Speaker 2: THAT'LL`,
   `Speaker 2: GET`, `Speaker 2: YOUR`... The `/transcripts/punctuated` endpoint returns
   whole punctuated sentences with a speaker on each one. That is the one to use.
5. **A free ElevateAI account holds one active token.** The token page says *"Your Free
   account is eligible for 1 active token."* Interactions belong to the token that
   created them, so after switching tokens, older interaction ids return **404** forever.
   A downloader that treats 404 as "not ready yet" then retries for half an hour per id.
6. **GHL's knowledge base takes tables as CSV only, 50 MB max.** The **Table Upload**
   dialog says *"Upload a CSV file to train your bot with product details or other
   structured data."* and *"CSV file only (max 50 MB)"*. Thousands of separate `.txt`
   files aren't a usable source. A few CSV files with one call per row are.

## How it works

```text
 recordings\*.mp3
      |  transcribe.py  (one file at a time)
      |    1. POST /v1/interactions                          -> interactionIdentifier
      |    2. POST /v1/interactions/{id}/upload   (the MP3)  -> 200/201/202
      |    3. GET  /v1/interactions/{id}/transcripts/punctuated  every 10 s until 200
      v
 transcripts\punctuated\{id}.punctuated.json   +   transcripts\punctuated_txt\{id}.txt
      |                                                 |
      |  download_punctuated.py (recovery: re-fetch     |  batch_to_49mb_csv.py
      |  ids from the log without re-uploading)         v
      |                                     transcripts\csv_batches_49mb\transcripts_batch_0001.csv ...
      |                                                 |
      |  convert.py / convert_json_txt_to_docx.py       v
      v  (optional Word copies)                 GHL > AI Agents > Knowledge Base > Tables
 transcripts\transcript doc\*.docx
```

Every script keeps a CSV log and skips work it has already done, so you can stop any of
them with **Ctrl+C** and run it again.

## What's in this repo

| Path | What it is | Where it goes |
|---|---|---|
| [`README.md`](README.md) | This guide | — |
| [`scripts/`](scripts/) | The five Python scripts, each used in its own step below | Copied into your recordings folder (Steps 4 to 7) |
| [`.gitignore`](.gitignore) | Keeps recordings, transcripts, CSVs and tokens out of git | Repo root |

## Table of contents

- [1. Requirements](#1-requirements)
- [2. Setup, step by step](#2-setup-step-by-step)
  - [Step 1: Put the recordings in one folder](#step-1-put-the-recordings-in-one-folder)
  - [Step 2: Install Python and the two packages](#step-2-install-python-and-the-two-packages)
  - [Step 3: Get an ElevateAI API token and test it](#step-3-get-an-elevateai-api-token-and-test-it)
  - [Step 4: Transcribe the folder](#step-4-transcribe-the-folder)
  - [Step 5: Recover transcripts without re-uploading (when needed)](#step-5-recover-transcripts-without-re-uploading-when-needed)
  - [Step 6: Pack the transcripts into CSV files under 49 MB](#step-6-pack-the-transcripts-into-csv-files-under-49-mb)
  - [Step 7: Optional: Word copies of each transcript](#step-7-optional-word-copies-of-each-transcript)
  - [Step 8: Upload the CSVs to the GHL knowledge base](#step-8-upload-the-csvs-to-the-ghl-knowledge-base)
- [Output formats](#output-formats)
- [Troubleshooting](#troubleshooting)
- [Limits](#limits)
- [Security](#security)

## 1. Requirements

| Need | Why |
|---|---|
| Windows with PowerShell (macOS/Linux work too; change the paths) | All commands below are PowerShell |
| Python 3.8 or later | 3.13 was used |
| `requests` and `python-docx` | `requests` for the API scripts, `python-docx` only for the optional Word step |
| An ElevateAI account and API token | Steps 3 to 5 only. **The service has been retired** (see the note at the top). |
| A folder of call recordings as `.mp3` | Twilio recordings in the original build |
| A GoHighLevel sub-account with AI Agents | Step 8 only |

## 2. Setup, step by step

### Step 1: Put the recordings in one folder

About 10 minutes, plus however long the download takes.

**1a.** Create this layout. The scripts expect it; you can change the base path at the
top of each script.

```text
C:\call-recordings\
    recordings\        <- every .mp3 goes here
    transcripts\       <- the scripts create everything below this
```

**1b.** Download your Twilio call recordings as MP3 into `C:\call-recordings\recordings`.
How you bulk-download them from Twilio isn't covered by these scripts: the original build
started from a folder that was already downloaded, plus the account's call recording log
exported as a CSV next to it. Keep that log. The MP3 file names are Recording SIDs
(`RE...mp3`), and the log is how you map a recording back to a call.

**1c.** Sort the folder by **Size** in File Explorer (**View → Details**, then click the
**Size** column) and look at both ends. Very small files are usually voicemail or a
hang-up. Very large files take longest to transcribe (Step 4 waits up to 3 hours for one).

> ⚠️ **Delete duplicate downloads.** A repeated download leaves pairs like
> `RE...bcd.mp3` and `RE...bcd (1).mp3`. The script treats them as two different calls
> and pays to transcribe both.

✅ **Check:** `C:\call-recordings\recordings` holds only `.mp3` files and no `(1)` copies.

### Step 2: Install Python and the two packages

About 5 minutes.

**2a.** Open **PowerShell** and type `python --version`.

**2b.** If you see *"Python was not found; run without arguments to install from the
Microsoft Store, or disable this shortcut from Settings > Apps > Advanced app settings >
App execution aliases."*, install Python from https://www.python.org/downloads/ (or run
`python` with no arguments to get the Microsoft Store version). Close and reopen
PowerShell, then repeat 2a.

**2c.** Install the packages:

```powershell
python -m pip install requests python-docx
```

Success looks like *"Successfully installed ... requests-..."*. A yellow warning that a
scripts folder *"is not on PATH"* is harmless here, because you always start the scripts
with `python`.

✅ **Check:** `python -c "import requests, docx; print('ok')"` prints `ok`. If it prints
`ModuleNotFoundError: No module named 'requests'`, repeat 2c.

### Step 3: Get an ElevateAI API token and test it

About 5 minutes.

> ⚠️ ElevateAI shut down on March 31, 2026, so this step can no longer be completed. It's
> kept so the guide matches what was built. If you port the scripts to another
> speech-to-text API, this is the step you replace.

**3a.** Sign in at https://app.elevateai.com/. In the left menu (**Home, Usage, Plans,
Status, Interactions, Manage API Tokens, Manage Accounts, Explore, Docs, Support**), open
**Manage API Tokens**.

**3b.** Click **+ New Token** (or use the one listed). The table shows **Token**,
**Last Request**, **Requests**, **Status** and **Action** (column names are cut off on
screen; wording may differ). Copy the token.

> ⚠️ **A free account holds one active token.** The banner reads *"Your Free account is
> eligible for 1 active token. Upgrade your account to be eligible for up to 5 active
> tokens."* If you make a new token halfway through, every interaction created with the
> old one becomes unreachable through the API (it returns 404). Write down the time you
> switch. Step 5 needs it.

**3c.** Put the token in an environment variable for this PowerShell window. The scripts
read `ELEVATEAI_API_TOKEN`; it never goes in a file.

```powershell
$env:ELEVATEAI_API_TOKEN = "<PASTE_TOKEN_HERE>"
```

To keep it across windows, use **System Properties → Environment Variables → New**
(user variable `ELEVATEAI_API_TOKEN`) and reopen PowerShell.

**3d.** Test a single *declare* call before running anything in bulk:

```powershell
python -c "import os, requests; r=requests.post('https://api.elevateai.com/v1/interactions', headers={'Content-Type':'application/json','X-API-Token':os.environ['ELEVATEAI_API_TOKEN']}, json={'type':'audio','languageTag':'en-us','vertical':'default'}); print(r.status_code, r.text)"
```

Success looks like:

```text
201 {"interactionIdentifier":"1b2c3d4e-0000-4000-8000-123456789abc"}
```

The new interaction appears on the **Interactions** page with status **Declared** (it has
no audio yet).

| You got | Cause | Fix |
|---|---|---|
| `401` | The token isn't being accepted | Re-copy it without spaces; check it's the active one on **Manage API Tokens** |
| `415` | Audio was sent to `/v1/interactions` | That endpoint only takes the JSON above. Audio goes to `/upload` (Step 4 does this). |

> ⚠️ **Copy the header exactly: `X-API-Token`, with `Content-Type: application/json` and
> a JSON body.** The working request in the build matched ElevateAI's own sample on its
> **Integrate our API** panel. Earlier attempts that sent form data or a different
> payload were rejected.

> ⚠️ **That sample panel shows your token in full.** Don't screenshot it.

✅ **Check:** the test prints `201` and an `interactionIdentifier`.

### Step 4: Transcribe the folder

About 10 minutes to start, then it runs on its own (thousands of calls took overnight).

**4a.** Save **all** of [`scripts/transcribe.py`](scripts/transcribe.py) as
`C:\call-recordings\transcribe.py`. Download the file. Don't copy it out of a chat
window: line-wrapped pastes caused most of the errors in the original build (see
Troubleshooting).

**4b.** Open it and check the settings at the top:

| Setting | Default | What it does |
|---|---|---|
| `SORT_MODE` | `"size_asc"` | Order files are sent in: `"size_asc"` smallest first (quick early results), `"size_desc"` largest first, `"alpha"` by file name |
| `ELEVATEAI_API_TOKEN` (environment variable) | — | Your token from Step 3. The script stops with *"Set the ELEVATEAI_API_TOKEN environment variable first"* if it's missing. |
| `BASE_PATH` | `C:\call-recordings` | Your folder from Step 1 |
| `REQUEST_TIMEOUT` | `120` | Seconds before one HTTP call gives up |

The wait for each file grows with its size: 20 minutes plus 6 minutes per MB, capped at
3 hours, checking every 10 seconds. A 0.5 MB call gets 23 minutes; a 10 MB call gets 80.

**4c.** Run it:

```powershell
cd "C:\call-recordings"
python transcribe.py
```

Runs [`scripts/transcribe.py`](scripts/transcribe.py).

For each MP3 it declares an interaction, uploads the file with the multipart field name
`filename mp3` (falling back to `file`), then polls the punctuated transcript until it's
ready. Success looks like:

```text
Total files found: 1200
Sorting by: size_asc
Already processed (from log): 0

Processing: RE0123456789abcdef0123456789abcdef.mp3 (0.46 MB)
   Declared: 1b2c3d4e-0000-4000-8000-123456789abc
   Upload OK.
   Waiting for punctuated transcript...
   Punctuated not ready (404). Check 1/132
   Punctuated not ready (404). Check 2/132
   Saved JSON: C:\call-recordings\transcripts\punctuated\1b2c3d4e-....punctuated.json
   Saved TXT : C:\call-recordings\transcripts\punctuated_txt\1b2c3d4e-....txt
```

A few `404` checks per file are normal. That's how the API says "not ready yet".

**4d.** What it writes:

| Output | Contents |
|---|---|
| `transcripts\punctuated\{interaction_id}.punctuated.json` | The full API response: `sentenceSegments` (participant, start/end in ms, phrase, score) and `redactionSegments` |
| `transcripts\punctuated_txt\{interaction_id}.txt` | One line per sentence: `Speaker 1: ...` / `Speaker 2: ...` |
| `processed_log.csv` (in `C:\call-recordings`) | One row per event. Columns: `timestamp, status, mp3_filename, interaction_id, size_mb, punctuated_json_path, punctuated_txt_path, notes`. `status` is `uploaded`, `success` or `error`. |

The transcript files are named by **interaction id**, not by MP3 name. `processed_log.csv`
is the only link between the two, so don't delete it.

**4e.** To stop, press **Ctrl+C**. To resume, run 4c again: every MP3 with a `success` row
in the log is skipped (*"Already processed (from log): N"*). A file that hit an `error`
is uploaded again on the next run.

> ⚠️ **"Processed" in the dashboard doesn't mean your script got it.** In the original
> build the dashboard showed **Processed** while the script printed
> `Current Status: None` forever, because it was polling the status endpoint (which has
> no status field). This script polls the transcript endpoint instead. If you write your
> own, do the same.

> ⚠️ **Don't use the plain `/transcript` endpoint.** It saved files where every word sat
> on its own line (`Speaker 2: THAT'LL` / `Speaker 2: GET` / ...). The punctuated
> endpoint, which this script uses, gives whole sentences.

> ⚠️ **An occasional `Punctuated HTTP 500` is transient.** The script waits and checks
> again; in the build the next check succeeded.

> ⚠️ **Don't edit the log while the script runs.** In the build the log was cleared by
> accident, so the script lost track of what was done. Copy it before opening it in
> Excel.

✅ **Check:** after a few files, `punctuated_txt` contains `.txt` files with readable
`Speaker 1:` / `Speaker 2:` sentences, and `processed_log.csv` has matching `success`
rows.

### Step 5: Recover transcripts without re-uploading (when needed)

About 5 minutes. Skip this step if Step 4 saved every file.

Use it when you have interaction ids in a log but no punctuated file for them: a run was
interrupted after upload, or an earlier version of your script saved the word-by-word
transcript. Uploading again would cost credits; this only downloads.

**5a.** Save **all** of [`scripts/download_punctuated.py`](scripts/download_punctuated.py)
as `C:\call-recordings\transcripts\download_punctuated.py`.

**5b.** Make the input log. The script looks for, in this order:

| File it reads | Use it for |
|---|---|
| `C:\call-recordings\processed_log1.csv` | **Make this one.** Copy `processed_log.csv`, then delete every row created before your last token change (Step 3b), so only ids your current token owns remain. |
| `C:\call-recordings\transcripts\processed_log.csv` | Used only if `processed_log1.csv` doesn't exist |

It doesn't need headers or particular columns. It pulls every GUID-shaped value
(`8-4-4-4-12` hex characters) out of every cell.

**5c.** Check the settings at the top: `BASE_PATH` (`C:\call-recordings`),
`MAX_RETRIES_404` (`180`) and `SLEEP_SECONDS` (`10`), so it waits up to 30 minutes per id.

**5d.** Run it from the `transcripts` folder:

```powershell
cd "C:\call-recordings\transcripts"
python .\download_punctuated.py
```

Runs [`scripts/download_punctuated.py`](scripts/download_punctuated.py).

Success looks like:

```text
Using log: C:\call-recordings\processed_log1.csv
Found 300 unique interaction_id(s) in processed_log.csv
Saving JSON to: C:\call-recordings\transcripts\punctuated
Saving TXT  to: C:\call-recordings\transcripts\punctuated_txt
[1b2c3d4e-0000-4000-8000-123456789abc] Downloading punctuated transcript...
[1b2c3d4e-0000-4000-8000-123456789abc] SAVED -> JSON and TXT
[2c3d4e5f-0000-4000-8000-123456789abc] SKIP already downloaded
Done.
```

(The *"in processed_log.csv"* wording is fixed text; it's reporting whichever log it
used.)

| Output | Contents |
|---|---|
| `transcripts\punctuated\{id}.punctuated.json` | Same as Step 4 |
| `transcripts\punctuated_txt\{id}.txt` | One line per sentence **with times**: `Speaker 1 [00:01-00:03]: ...` |
| `transcripts\punctuated_download_log.csv` | `timestamp, status, interaction_id, json_file, txt_file, notes`; `status` is `success`, `skip_exists` or `error` |

> ⚠️ **Endless `404 not ready (attempt N/180)` on old ids means the wrong token, not "not
> ready".** The API answers 404 both while a transcript is being made and when the id
> belongs to another token. In the build, ids from before the token change kept
> returning 404 even when the downloader tried both tokens. Once the input log held only
> rows created with the current token (5b), every id downloaded with no errors. Press
> **Ctrl+C** rather than waiting 30 minutes per id.

> ⚠️ **In PowerShell, `download_punctuated.py` on its own fails** with *"The command
> download_punctuated.py was not found, but does exist in the current location."* Always
> start it with `python`.

✅ **Check:** `punctuated_txt` gains files and `punctuated_download_log.csv` shows
`success` rows.

### Step 6: Pack the transcripts into CSV files under 49 MB

About 5 minutes.

**6a.** Save **all** of [`scripts/batch_to_49mb_csv.py`](scripts/batch_to_49mb_csv.py) as
`C:\call-recordings\transcripts\batch_to_49mb_csv.py`.

**6b.** Check the settings at the top:

| Setting | Default | What it does |
|---|---|---|
| `INPUT_DIR` | `C:\call-recordings\transcripts\punctuated_txt` | The folder of `.txt` transcripts to combine |
| `OUTPUT_DIR` | `C:\call-recordings\transcripts\csv_batches_49mb` | Where the CSVs and the log go |
| `MAX_BYTES_PER_BATCH` | `49 * 1024 * 1024` | Size cap per CSV, safely under GHL's 50 MB |
| `SORT_MODE` | `"alpha"` | `"alpha"` by file name or `"mtime"` by date modified |
| `TEST_LIMIT` | `None` | Set to `2` for a first test run |

**6c.** Test with two files first. Set `TEST_LIMIT = 2`, save, and run:

```powershell
cd "C:\call-recordings\transcripts"
python batch_to_49mb_csv.py
```

Runs [`scripts/batch_to_49mb_csv.py`](scripts/batch_to_49mb_csv.py).

Open `csv_batches_49mb\transcripts_batch_0001.csv` in Excel or a CSV viewer. It should
have a header row `file_id, conversation` and two rows, each holding one whole call in a
single cell. Then **delete the whole `csv_batches_49mb` folder** (otherwise the log says
those two files are done and a fresh batch starts at `0002`), set `TEST_LIMIT = None`,
and run again.

**6d.** Success looks like:

```text
Total files found: 1200
Already processed (per log): 0
Remaining to batch: 1200
Max per batch: 51380224 bytes (~49MB)
Sort mode: alpha

Starting batch: transcripts_batch_0001.csv
Processed 50/1200...
...
Done.
Last batch: transcripts_batch_0001.csv | rows: 1200 | approx bytes: 6100000
Log: C:\call-recordings\transcripts\csv_batches_49mb\csv_processed_log.csv
```

| Output | Contents |
|---|---|
| `csv_batches_49mb\transcripts_batch_0001.csv`, `_0002.csv`, ... | Columns `file_id` (file name without `.txt`) and `conversation` (the whole transcript). Each file stays under 49 MB. |
| `csv_batches_49mb\csv_processed_log.csv` | `timestamp, batch_file, filename, bytes`, one row per transcript batched |

Run it again later and it only adds transcripts that aren't in the log, starting a new
batch file rather than overwriting one.

> ⚠️ **Point `INPUT_DIR` at one kind of transcript.** The first batch in the build mixed
> folders, so one row was punctuated sentences and the next was the word-per-line
> output. Batch `punctuated_txt` only.

> ⚠️ **About two thousand punctuated calls came to roughly 11 MB**, so you'll likely get
> one CSV. The 49 MB cap matters for longer calls or a much bigger archive.

✅ **Check:** every `transcripts_batch_*.csv` is under 49 MB and opens with two columns.

### Step 7: Optional: Word copies of each transcript

About 5 minutes. The CSV route (Step 6) replaced this in the original build, because the
knowledge base takes CSV tables. Use it if people need to read individual calls.

Both scripts accept either kind of `.txt`: ElevateAI JSON saved with a `.txt` extension
(from the plain `/transcript` endpoint, which they rebuild into `Speaker N:` lines,
starting a new line when the speaker changes or after a 1.5-second gap), or text that's
already readable (copied as is).

**7a. Test on one file.** Save **all** of [`scripts/convert.py`](scripts/convert.py) as
`C:\call-recordings\transcripts\convert.py`. At the top, `TEST_ONE = True` converts one
file; set `TEST_FILENAME` to a file name, or leave `None` for the first file. It reads
`INPUT_FOLDER` (default `C:\call-recordings\transcripts`; set it to
`os.path.join(BASE_PATH, "transcripts", "punctuated_txt")` to convert the punctuated
text). Run:

```powershell
cd "C:\call-recordings\transcripts"
python convert.py
```

Runs [`scripts/convert.py`](scripts/convert.py).

Success: *"OK: RE0123....txt -> C:\call-recordings\transcripts\transcript doc\RE0123....docx"*.
Open the `.docx` and check it reads well. The log is `C:\call-recordings\doc_conversion_log.csv`.

**7b. Convert the whole folder.** Save **all** of
[`scripts/convert_json_txt_to_docx.py`](scripts/convert_json_txt_to_docx.py) **in the
folder whose `.txt` files you want converted** (it works on the folder you run it from),
and run it there:

```powershell
cd "C:\call-recordings\transcripts\punctuated_txt"
python convert_json_txt_to_docx.py
```

Runs [`scripts/convert_json_txt_to_docx.py`](scripts/convert_json_txt_to_docx.py).

Each file prints *"✔ Converted: name.txt"*. Output goes to a `transcript doc` subfolder of
that folder, with `doc_conversion_log.csv` next to the script. Files already converted
are skipped silently.

| Output | Contents |
|---|---|
| `transcript doc\{name}.docx` | One paragraph per transcript line |
| `doc_conversion_log.csv` | `timestamp, status, input_txt, output_docx, notes` (7a) or `timestamp, status, input_file, output_file, notes` (7b); transcripts under 10 (7a) or 20 (7b) characters are logged as `error` |

✅ **Check:** the `.docx` opens in Word with one `Speaker N:` line per paragraph.

### Step 8: Upload the CSVs to the GHL knowledge base

About 5 minutes per CSV.

**8a.** In your GHL sub-account, open **AI Agents** and the **Knowledge Base** tab (the
tabs read **Getting Started, Agent Studio, Voice AI, Conversation AI, Knowledge Base,
Agent Templates, Content AI**). Open the knowledge base your bot uses, or create one.

**8b.** The page shows **Knowledge Sources** (*"Add and manage sources your bot will use
to learn and respond to users."*) with tabs **All, Web Crawler, FAQ, Tables, Rich Text,
File Upload**. Click **+ Add Source** and choose **Tables** (or open the **Tables** tab)
(wording may differ).

**8c.** The **Table Upload** dialog has three stages: **Upload File → Column Selection →
Summary**.
1. **Upload File:** **Click to upload or drag and drop** your
   `transcripts_batch_0001.csv` (*"CSV file only (max 50 MB)"*).
2. **Name:** a name for the table source, e.g. `Call transcripts 1`.
3. Click **Next**. On **Column Selection**, keep both `file_id` and `conversation`
   (wording may differ). Click **Next**, review the **Summary**, and finish.

**8d.** Repeat for each `transcripts_batch_*.csv`.

> ⚠️ **Raw transcripts are evidence, not answers.** They teach the bot how your team
> talks, but they also contain prices, offers and promises that change. In the original
> build the transcripts were also given to an LLM to draft FAQ and rich-text knowledge
> entries (pricing, policies, session types), which were checked against current prices
> before they went in. Expect to correct the bot where old calls quote old terms.

✅ **Check:** back on the **All** tab, the **Tables** card's count has gone up by one per
CSV.

## Output formats

**Punctuated JSON** (from ElevateAI, saved unchanged):

```json
{
  "sentenceSegments": [
    {"participant": "participantOne", "startTimeOffset": 1410, "endTimeOffset": 3210,
     "phrase": "Thanks for calling, how can I help?", "score": 0.99}
  ],
  "redactionSegments": [
    {"startTimeOffset": 24650, "endTimeOffset": 27530, "result": "CVV", "score": 0}
  ]
}
```

`participantOne` / `participantTwo` become `Speaker 1` / `Speaker 2`. ElevateAI doesn't
say which one is your team. Redacted card numbers, CVVs and Social Security numbers
appear in `phrase` as `# # # #`.

**Text from Step 4:** `Speaker 1: Thanks for calling, how can I help?`

**Text from Step 5:** `Speaker 1 [00:01-00:03]: Thanks for calling, how can I help?`

**CSV from Step 6:**

```csv
file_id,conversation
1b2c3d4e-0000-4000-8000-123456789abc,"Speaker 1: Thanks for calling, how can I help?
Speaker 2: Hi, I have a question about pricing."
```

## Troubleshooting

Most of these came straight from the original build.

| You got | Cause | Fix |
|---|---|---|
| *"Python was not found; run without arguments to install from the Microsoft Store..."* | Python isn't installed; Windows' Store shortcut answered | Step 2b |
| `can't open file 'C:\\call-recordings\\transcribe.py': [Errno 2] No such file or directory` | The script isn't saved in that folder, or saved as `transcribe.py.txt` | Save it as in Step 4a; turn on **View → Show → File name extensions** to check |
| `ModuleNotFoundError: No module named 'requests'` (or `docx`) | Packages not installed | Step 2c |
| *"Set the ELEVATEAI_API_TOKEN environment variable first"* | New PowerShell window, variable not set | Step 3c |
| `SyntaxError: unterminated string literal` / `unterminated f-string literal` / `'{' was never closed` / `IndentationError: expected an indented block after 'except'` | The script was pasted from a chat and a long line wrapped, or an emoji broke a string | Download the file from this repo instead of pasting (Step 4a) |
| `FAILED (Status 415)` | Audio sent to `/v1/interactions` | Use the two-call flow in [`scripts/transcribe.py`](scripts/transcribe.py) (Step 4) |
| `Declare failed 401` | Token not accepted, or wrong header | Step 3d: header `X-API-Token`, JSON body |
| `Current Status: None` (or `status is unknown`) forever, dashboard says **Processed** | Polling the status endpoint, which has no status field | Poll the transcript endpoint (Step 4, first ⚠️) |
| Transcript files with one word per line | Saved from `/transcript`, not `/transcripts/punctuated` | Step 4; recover old ids with Step 5 |
| `Punctuated not ready (404)` for a long time on a big file | Long calls take longer | Let it run; the wait scales with size (Step 4b) |
| `404 not ready (attempt N/180)` on every old id | Ids belong to a different (old) token | Step 5b: keep only rows after the token change |
| `Found 0 unique interaction_id(s)` | The log it found has no ids, or it read the wrong file | Check `processed_log1.csv` exists and has GUIDs (Step 5b) |
| *"The command download_punctuated.py was not found..."* | PowerShell doesn't run scripts from the current folder by name | `python .\download_punctuated.py` (Step 5d) |
| Batch starts at `transcripts_batch_0002.csv` and skips files | The two-file test is still in the log | Delete `csv_batches_49mb` and rerun (Step 6c) |
| CSV mixes readable and word-per-line rows | `INPUT_DIR` pointed at a mixed folder | Point it at `punctuated_txt` (Step 6) |
| The same call transcribed twice | `RE... (1).mp3` duplicate in `recordings` | Step 1c |

## Limits

- **ElevateAI is retired** (March 31, 2026). The API steps no longer run against it.
- **One file at a time.** Nothing runs in parallel, so thousands of calls take hours.
  That's deliberate: it kept errors readable and spend predictable.
- **English only as written.** The declare call sends `languageTag: "en-us"`.
- **Two speakers, unnamed.** You get `Speaker 1` and `Speaker 2`, not "agent" and
  "customer", and which is which can change from call to call.
- **Re-runs re-upload failures.** A file with an `error` row (and no `success`) is
  uploaded again next time.
- **Files are named by interaction id.** Keep `processed_log.csv` to map a transcript back
  to its MP3 and so to the call.

## Security

- **Call recordings and transcripts are your customers' conversations.** Keep them out of
  any repo (this repo's [`.gitignore`](.gitignore) blocks them) and out of shared drives
  that don't need them. Card numbers and SSNs are redacted in the punctuated output, but
  names, phone numbers and addresses are not.
- **The token lives in an environment variable**, never in a script. In the original
  build it was pasted into the scripts, into chats with two different AI assistants and
  into terminal screenshots; each of those is a working copy until the token is
  revoked. If a token was exposed, deactivate or delete it on **Manage API Tokens**
  (the power and bin icons in the **Action** column).
- ElevateAI's **Integrate our API** panel and **Manage API Tokens** page show the token
  in full. Don't screenshot them.

---

Built by [Jibril Sulaiman](https://github.com/carljibrilsulaimanii).
