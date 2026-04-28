---
name: translate-dataroom
description: >
  Translate a folder of Japanese data room documents to English using GCP Cloud Translation API v3
  with raw XML manipulation for Excel files (preserving formulas, pivots, and charts perfectly).
  Use this skill whenever the user mentions translating Japanese documents, translating a data room,
  converting Japanese files to English, running a translation job, or mentions translating XLSX, PDF,
  DOCX, CSV, TXT, or MP4 files from Japanese. Also trigger when the user says "translate this folder,"
  "translate these files," references Japanese-to-English document translation, or mentions
  data room translation in any context - even if they don't use the exact phrase "data room."
  This skill handles the entire end-to-end workflow: scanning, estimating cost, reorganizing folders,
  converting legacy formats, translating, transcribing audio/video, verifying Excel integrity,
  and notifying on completion.
---

# Data Room Translation: Japanese to English

Translates an entire folder of Japanese documents to English using Google Cloud Translation API v3.
Uses raw XML manipulation for Excel files to perfectly preserve formulas, pivot tables, and charts -
something neither the Document Translation API nor openpyxl can do reliably.

## Environment

This skill is **cross-platform** — works on Windows, macOS, and Linux. Scripts are bundled inside
this skill folder, which lives at `~/.claude/skills/translate-dataroom/` on every machine.

### Bundled scripts (all in `scripts/` subfolder)

| Script | Purpose |
|--------|---------|
| `translate_dataroom.py` | Main translation engine |
| `transcribe_mp4.py` | Whisper API for audio/video |
| `timer.py` | Phase-aware end-to-end progress display |
| `workflow_state.py` | Phase plan / state file helper for the timer |
| `convert_xls.py` | **Cross-platform** .xls → .xlsx converter (xlrd + openpyxl primary; Excel COM fallback on Windows) |
| `decrypt_excel.py` | **Cross-platform** decryptor for password-protected Office files (msoffcrypto-tool) |
| `fix_row_heights.py` | Post-translation Excel auto-fit safety net |
| `retranslate_xlsx_strings.py` | Single-file re-run utility |
| `check_setup.py` | Environment verification (run once on each new machine) |
| `drive_sync.py` | Cross-platform Drive API client — `down`/`up`/`ls`/`resolve` for cloud sandboxes (no Drive Streaming mount) |
| `_cloud_creds.py` | Auto-imported helper that materializes `GOOGLE_APPLICATION_CREDENTIALS_JSON` env var to a temp file. No-op on machines where `GOOGLE_APPLICATION_CREDENTIALS` already points to a real file. |

### Per-machine configuration

The skill needs these resolved before each run:

| Item | Windows default (`zach@broadbandcap.com`) | macOS default | How to override |
|------|------------------------------------------|---------------|-----------------|
| `<python>` | `C:\Users\zacha\AppData\Local\Programs\Python\Python312\python.exe` | `/opt/homebrew/bin/python3` (Homebrew) or `/usr/local/bin/python3` | `where python` (Win) / `which python3` (Mac) |
| `<HOME>` | `C:\Users\zacha` | `/Users/zach` | `$HOME` env var |
| GCP service account JSON | `C:\Users\zacha\OneDrive\Claude Apps\claude-skills.json` | `~/Library/CloudStorage/OneDrive-Personal/Desktop/Files to mount in Claude/concise-faculty-492916-r8-81bb2087af44.json` | Pass via `GOOGLE_APPLICATION_CREDENTIALS` env var |
| GCP project ID | `concise-faculty-492916-r8` | (same) | `GCP_PROJECT_ID` env var |
| OpenAI API key (MP4 only) | (asked at runtime) | (asked at runtime) | `OPENAI_API_KEY` env var |
| Drive Streaming root | `C:\Users\zacha\Google Drive Streaming\` or `G:\My Drive\` / `H:\My Drive\` | `~/Library/CloudStorage/GoogleDrive-<email>/` | Inspect filesystem or run `check_setup.py` |
| Drive FS metadata DB (for trash recovery) | `C:\Users\zacha\AppData\Local\Google\DriveFS\<account>\mirror_metadata_sqlite.db` | `~/Library/Application Support/Google/DriveFS/<account>/mirror_metadata_sqlite.db` | (same DB schema both platforms) |

### Cloud (claude.ai/code, claude.ai chat sandbox) configuration

When running in a cloud sandbox, file paths like `C:\...` and Drive Streaming
mount points don't exist. Configure via env vars instead:

| Env var | Value | Notes |
|---------|-------|-------|
| `GOOGLE_APPLICATION_CREDENTIALS_JSON` | Full contents of the GCP service account JSON (as a single string) | `_cloud_creds.py` writes this to a temp file at import time and points `GOOGLE_APPLICATION_CREDENTIALS` at it. Required. |
| `GCP_PROJECT_ID` | `concise-faculty-492916-r8` | Required (also default in code). |
| `OPENAI_API_KEY` | Your OpenAI key | Only required if transcribing MP4 files. |
| `DRIVE_OAUTH_CLIENT_SECRETS` | (optional) path to a client_secrets.json | Only if you prefer OAuth over service-account auth for Drive. SA is preferred in cloud. |

**Cloud workflow (Drive in → Drive out, no local files):**

```
python -u scripts/drive_sync.py down "<Japanese folder URL>" /tmp/japanese
python -u scripts/translate_dataroom.py /tmp/japanese /tmp/english
python -u scripts/drive_sync.py up /tmp/english "<parent Drive folder ID>" --name "English"
```

The service account email (from the SA JSON's `client_email` field) **must be
shared as a member** on the source Drive folder for `down` to work, and on the
destination parent folder for `up` to work.

**Pip install in cloud sandbox** — at the start of a cloud session, run:

```
pip install google-cloud-translate google-api-python-client google-auth google-auth-oauthlib \
            PyPDF2 xlrd openpyxl msoffcrypto-tool openai pydub imageio-ffmpeg
```

Skip `pywin32` — it is Windows-only and the pure-Python paths cover everything
needed in a Linux sandbox.

### Always run `check_setup.py` on a new machine first

```
"<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/check_setup.py"
```

It verifies Python version, all required packages, env vars, GCP auth, and detects the Drive Streaming
mount path automatically.

## Prerequisites

The following Python packages must be installed (cross-platform):

```
"<python>" -m pip install google-cloud-translate google-api-python-client google-auth-oauthlib PyPDF2 xlrd openpyxl msoffcrypto-tool openai pydub imageio-ffmpeg
```

- **`google-cloud-translate`, `PyPDF2`**: Translation API + PDF handling
- **`google-api-python-client`, `google-auth-oauthlib`**: Drive API for cloud workflow (Step 1.6)
- **`xlrd` + `openpyxl`**: Pure-Python .xls → .xlsx conversion (primary path, all platforms)
- **`msoffcrypto-tool`**: Cross-platform decryption of password-protected Office files
- **`openai`, `pydub`, `imageio-ffmpeg`**: MP4 transcription via Whisper API

### Windows-only optional package

```
"<python>" -m pip install pywin32
```

`pywin32` enables an Excel COM fallback for .xls conversion that preserves formulas/charts.
**Skip on macOS/Linux** — the pure-Python path handles all conversions reliably (without
formulas, but with all values + text).

## Workflow

Follow these steps in order. Do NOT skip the estimation step.

### Step 1: Scan, Estimate Cost, and Choose Translation Scope

When the user provides a folder path, scan it to count files, measure sizes, and compute a
**character-level cost estimate**. GCP charges $20 per million characters — file count alone is
meaningless; a single CSV can cost more than 50 small Excel files.

Run a Python script that walks the folder and for each file type measures ACTUAL Japanese characters:

```python
# For XLSX/XLSM: extract xl/sharedStrings.xml from the ZIP, count Japanese <t> tag chars
# For CSV: read with correct encoding, count chars in cells containing Japanese
# For PDF: use PyPDF2 page.extract_text(), count chars. Note: this UNDERESTIMATES true
#          Document API cost — GCP processes OCR, image text, metadata, and embedded content
#          that PyPDF2 cannot extract. Apply a 3-5x multiplier on PDF char counts.
# For DOCX/PPTX: estimate ~3,000 chars/page for DOCX, or 30% of file size in KB for PPTX
# For TXT: read file, count Japanese chars
# For XLS: count files only (will be converted to XLSX first)
```

**Cost formula**: `(total_chars / 1,000,000) × $20`

For PDFs specifically, the Document Translation API charges based on the full document content
processed (including OCR, image text, formatting, and metadata) — NOT just the visible text that
PyPDF2 can extract. Multiply PDF extractable character counts by **3–5x** to estimate true cost.

Present results in **two tables**:

**Table 1: File inventory**

| Type | Count | Size | Japanese Chars | Est. Cost |
|------|-------|------|----------------|-----------|
| XLSX | N | X MB | X chars | $X.XX |
| CSV | N | X MB | X chars | $X.XX |
| PDF | N | X MB | X chars (×4 for API) | $X.XX |
| ... | | | | |
| **Total** | | | | **$X.XX** |

**Table 2: Translation scope options**

Ask the user which scope they want. Present three options with cost and rationale:

| Option | Scope | Est. Cost | Notes |
|--------|-------|-----------|-------|
| **A** | Everything | $X.XX | Full translation including PDFs via Document API |
| **B** | Everything except PDFs | $X.XX | PDFs copied as-is; avoids expensive Document API |
| **C** | PDFs only | $X.XX | For when text files are done and PDFs still need translation |

Use AskUserQuestion to let the user select A, B, or C before proceeding.

When the user selects **Option B**, the translation script should be launched with a `--skip-pdf` flag
or equivalent filter so PDFs are copied to the English folder untranslated. If the user later wants
to add PDF translation, they can re-run with **Option C** which translates only PDFs and slots them
into the existing English folder structure.

When the user selects **Option C** (PDFs only), launch the translation script targeting only PDF files.
The script's resumability (`.translation_state.json`) ensures already-translated non-PDF files are
skipped. The translated PDFs are placed into the correct English subfolder paths matching the
Japanese folder structure.

### Step 1.5: Launch End-to-End Progress Timer (do this NOW, right after scope is chosen)

**The timer should launch IMMEDIATELY after the user picks a scope** — not after conversion or
after translation kicks off. By that point you already have:
- File counts by type (from the scan)
- Total chars / pages (from the estimate)
- Whether xls conversion is needed (xls_count)
- Whether source is on Drive Streaming (copy/deploy phases needed)
- Whether MP4 transcription is needed

That's enough for a reasonable end-to-end ETA.

**Workflow**:

1. Build an inventory dict from the scan results:
   ```python
   inventory = {
     "xls_count": <#xls files>,
     "copy_to_local_mb": <total source MB if Drive Streaming, else 0>,
     "translate_files": {"xlsx": N, "pdf": N, ...},  # by extension, after applying scope filter
     "translate_chars": <total JP chars from text files>,
     "translate_pdf_pages": <total PDF pages>,
     "mp4_minutes": <est audio minutes if transcribing, else 0>,
     "deploy_to_drive_mb": <est translated output MB if Drive Streaming, else 0>,
     "xlsx_count": <#xlsx files for verify phase>,
   }
   ```

2. Call `workflow_state.py estimate <inv.json>` to get a phase plan.

3. Call `workflow_state.py init <state_path> <english_folder> <job_name> --plan-json <plan>`
   to write `.workflow_state.json` into the English folder (the orchestrator should create
   the English folder first if it doesn't exist).

4. Launch the phase-aware timer in a separate window. Cross-platform:

   **Windows**:
   ```
   powershell.exe -Command "Start-Process -FilePath '<python>' -ArgumentList '<HOME>\.claude\skills\translate-dataroom\scripts\timer.py', '--state', '<english_folder>\.workflow_state.json'"
   ```

   **macOS** (opens new Terminal window):
   ```
   osascript -e 'tell app "Terminal" to do script "<python> <HOME>/.claude/skills/translate-dataroom/scripts/timer.py --state <english_folder>/.workflow_state.json"'
   ```

   **Linux** (depends on terminal emulator — example for gnome-terminal):
   ```
   gnome-terminal -- "<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/timer.py" --state "<english_folder>/.workflow_state.json"
   ```

   **Fallback (any platform)**: run inline in the current shell as a background process — you'll
   just see the timer output mixed with translation output:
   ```
   "<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/timer.py" --state "<english_folder>/.workflow_state.json" &
   ```

5. As each phase begins, update the state file:
   ```
   "<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/workflow_state.py" advance <state_path> <phase_name>
   ```
   Phase names: `convert_xls`, `copy_to_local`, `translate`, `transcribe`, `deploy_to_drive`, `verify`.

6. For phases without auto-detection (everything except `translate`), update progress
   periodically with a fraction:
   ```
   "<python>" workflow_state.py progress <state_path> 0.5
   ```
   The `translate` phase auto-updates from `.translation_state.json`.

7. When everything is done, mark complete:
   ```
   "<python>" workflow_state.py complete <state_path>
   ```

The timer beeps 5 times when complete and the user can close the window. The full ETA
self-corrects as actual elapsed vs estimated diverges (clamped to 0.5×–3.0× to avoid
wild swings).

### Step 2: Convert Legacy Excel Files (.xls → .xlsx)

Use the bundled cross-platform converter:

```
"<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/convert_xls.py" "<folder>"
```

**Behavior** (cross-platform):
- **Default**: pure Python via `xlrd` + `openpyxl` — fast, reliable on Windows/Mac/Linux. Loses
  formulas/charts but preserves all cell values + text. Acceptable for most dataroom files.
- **`--password <pw>`**: tries the password (try `RENGA2025` first for K-Link files)
- **`--use-excel`** (Windows-only opt-in): tries Excel COM first to preserve formulas, falls
  back to pure-Python on hang or failure. Skip this flag on macOS — Excel COM is unavailable.

**Why pure-Python is the default**: Excel COM hung on us multiple times in past runs (especially
with deeply-nested Japanese paths). Pure-Python converted 21 files in seconds during a recent
run after Excel COM stalled for 20+ minutes. Use `--use-excel` only when formula preservation
is critical AND you're on Windows.

**Encrypted .xls / .xlsx files**: run `decrypt_excel.py` first (see Step 2.5).

### Step 1.6 (cloud / no-Drive-Streaming): Sync Files via Drive API

**When to use this**: running on cloud Claude Code (claude.ai/code), running on a machine
without Drive Streaming installed, or when source folder is given as a Drive URL instead of
a local path. Skip this step if you're working with a local folder path.

```
"<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/drive_sync.py" down "<DRIVE_URL_OR_ID>" "<LOCAL_TEMP>"
```

This downloads everything in the Drive folder (recursively) to a local temp directory. You
then run the rest of the workflow against that local copy. When done, sync results back:

```
"<python>" drive_sync.py up "<LOCAL_ENGLISH_FOLDER>" "<DRIVE_PARENT_FOLDER_ID>" --name "English"
```

**Authentication priority** (drive_sync.py auto-tries):
1. **Service account** via `GOOGLE_APPLICATION_CREDENTIALS` (preferred, no UI needed). The
   service account email must be added as a member of the Shared Drive containing the folder.
2. **OAuth installed-app flow** with cached token at `~/.claude/skills/translate-dataroom/.drive_token.json`.
   Requires `DRIVE_OAUTH_CLIENT_SECRETS` env var pointing to OAuth client_secrets.json.

**One-time setup for service account auth**:
1. Enable Drive API on the GCP project: `https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=<PROJECT_ID>`
2. Add the service account email (from the JSON's `client_email` field) as a member of the
   Shared Drive (Content manager role for read+write, Viewer for read-only).

For K-Link specifically (current setup):
- Service account email: `claude-skills@concise-faculty-492916-r8.iam.gserviceaccount.com`
  (single shared identity used by all of zach's Claude Code skills that touch Google services)
- Shared Drive: `Renga` (where K-Link folder lives)
- Required role: **Content manager** (so it can write English translations back)
- JSON key location: `~/OneDrive/Claude Apps/claude-skills.json` (Windows) or
  `~/Library/CloudStorage/OneDrive-Personal/Claude Apps/claude-skills.json` (Mac)

### Step 2.5: Decrypt Password-Protected Office Files (cross-platform)

If any .xlsx, .xlsm, .docx, or .pptx files are password-protected, decrypt them in-place
before translation:

```
"<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/decrypt_excel.py" "<folder>"
```

The script automatically tries `RENGA2025` (K-Link convention). To pass a different password:

```
"<python>" decrypt_excel.py "<folder>" --password "<password>"
```

**Behavior**:
- Uses **`msoffcrypto-tool`** (pure Python, cross-platform) as the primary path
- Creates a `.bak` backup of the original before in-place replacement (use `--no-backup` to skip)
- Reports per-file: `decrypted` / `not encrypted` (skip) / `wrong password` / `failed`
- Files that aren't encrypted are skipped silently

**Detection trick**: a `.xlsx` with magic bytes `D0CF11E0` instead of `PK` (zip) is encrypted.
You can also detect by trying `zipfile.ZipFile(path)` — if it raises `BadZipFile`, the file is
likely encrypted (or genuinely corrupt).

**Failure mode**: If the password isn't `RENGA2025` and you don't know the right one, leave
the file as-is — it'll be copied untranslated, and the user can address it separately.

### Step 3: Reorganize Folder Structure

Once confirmed, move all files into a "Japanese" subfolder and create an "English" subfolder.
Preserve any nested directory structure. Skip items already named "Japanese" or "English".

### Step 4: Verify Era Date Conversion is Active

CRITICAL: The GCP Translation API mistranslates Japanese Imperial era dates (和暦). For example,
`令和07年03月01日` becomes "March 1, 2020" instead of the correct "March 1, 2025". This affects
ALL text-based files (CSV, Excel, TXT, filenames).

The translation script (`translate_dataroom.py`) has a built-in `convert_era_dates()` function that
automatically converts era dates to Gregorian BEFORE sending text to GCP. Before every translation run,
verify this function exists in the script. If for any reason the script is recreated or modified,
ensure the era conversion logic is present and hooked into both `translate_text()` and
`translate_text_batch()`.

The function must handle:
- Full kanji: 令和/平成/昭和/大正/明治 + year + 年月日
- Abbreviated: R/H/S/T/M + year + dot/slash separators
- Base years: 令和=2018+n, 平成=1988+n, 昭和=1925+n, 大正=1911+n, 明治=1867+n

### Step 5: Verify Row-Height Auto-Fit is Active

CRITICAL: Japanese is information-dense — translated English text typically expands **3–4× in
character count** (e.g., a 47-char Japanese cell becomes a 189-char English cell). When the
original spreadsheet has rows with `<row ht="15" customHeight="1">`, the row height is *locked*,
so wrapped English text overflows invisibly below the row boundary, making files hard to read.

The translation script (`translate_dataroom.py`) strips `customHeight="1"` (and `customHeight="true"`)
from every `<row>` element inside `xl/worksheets/*.xml` during XLSX/XLSM translation. Excel then
auto-fits row heights when the file is opened, so wrapped English text is fully visible.
The `ht="..."` hint is preserved; Excel grows rows as needed on open.

Before every translation run, verify this logic is present in the worksheet-rewrite block of
`translate_xlsx_raw()` (look for the `<row\b ... customHeight=` regex). If the script is recreated
or modified, ensure the worksheet loop runs unconditionally (not gated on `name_map`) and applies
the regex strip to every worksheet XML.

If a legacy translated folder is missing the fix, run the row-height fixer as a one-shot:
```
"<python>" "<HOME>/.claude/skills/translate-dataroom/scripts/fix_row_heights.py" "<FOLDER>/English"
```
The fixer is idempotent and only modifies files that still have locks. Files open in Excel will
be reported as LOCKED and skipped — close them and re-run.

### Step 6: Run Translation

Launch the translation script as a background process with 10-minute timeout.
Use the appropriate flag based on the user's scope selection from Step 1:

**Option A (everything):**
```
GOOGLE_APPLICATION_CREDENTIALS="<service-account-json>" GCP_PROJECT_ID="<project-id>" "<python>" -u "<HOME>/.claude/skills/translate-dataroom/scripts/translate_dataroom.py" "<FOLDER>/Japanese" "<FOLDER>/English"
```

**Option B (skip PDFs):** add `--skip-pdf` flag at end of Option A command.

**Option C (PDFs only):** add `--pdf-only` flag at end of Option A command.

**Defaults for `zach@broadbandcap.com` on Windows**:
- `<python>` = `C:\Users\zacha\AppData\Local\Programs\Python\Python312\python.exe`
- `<service-account-json>` = `C:\Users\zacha\OneDrive\Claude Apps\claude-skills.json`
- `<project-id>` = `concise-faculty-492916-r8`
- `<HOME>` = `C:\Users\zacha`

Always use `python -u` for unbuffered output. The script's resumability ensures Option C
can be run after Option B without re-translating already-completed files.

### Step 7: (Timer already launched in Step 1.5)

The end-to-end timer should have launched right after the user picked a scope.
At this point you should be calling `workflow_state.py advance <state_path> translate`
to mark that the translate phase has begun. The timer's `translate` phase auto-updates
from `.translation_state.json` — no further action needed during translation.

If for some reason the early-launch was skipped (e.g. legacy invocation), you can
launch the legacy translation-only timer:

```
powershell.exe -Command "Start-Process -FilePath '<python>' -ArgumentList '<HOME>\.claude\skills\translate-dataroom\scripts\timer.py', '<TOTAL_FILES>', '<FOLDER>\English'"
```

### Step 8: Transcribe and Translate MP4/Audio Files

For any MP4 (or other audio/video) files found, transcribe the spoken Japanese and translate to English
using the OpenAI Whisper API. This requires the user's OpenAI API key (ask if not provided).

Launch the transcription script as a background process:

```
OPENAI_API_KEY="<KEY>" "<python>" -u "<HOME>/.claude/skills/translate-dataroom/scripts/transcribe_mp4.py" "<MP4_PATH>" "<OUTPUT_DIR>"
```

The script:
1. Extracts audio from the video as MP3 (using ffmpeg bundled via imageio-ffmpeg)
2. Splits audio into <25 MB chunks if needed (Whisper API limit)
3. Transcribes in Japanese with timestamps via Whisper API (`whisper-1` model)
4. Translates to English via GPT-4o-mini (preserving timestamps, business terminology)
5. Saves both `_transcript_JA.txt` and `_transcript_EN.txt` in the English output folder

Cost: ~$0.006/minute for Whisper + ~$0.01 for translation. A 60-minute video costs ~$0.50.

The original MP4 file is copied as-is to the English folder (not re-encoded).

### Step 9: Verify Excel Files

After translation completes, check all XLSX/XLSM files:
- ZIP integrity (not corrupt)
- workbook.xml present
- Formulas checked for residual Japanese
- No `<row>` elements still carry `customHeight="1"` (would indicate the row-height
  auto-fit pass in `translate_xlsx_raw()` was bypassed). If any are found, run
  `fix_row_heights.py` against the English folder to remediate before handing off.

A small number of formulas with residual Japanese is normal - these are text literals
inside formulas (month names, VLOOKUP keys), not broken references. The critical check
is that sheet name references were updated and no ZIPs are corrupt.

### Step 10: Report Results

Present a final summary: files translated/copied/failed, verification results,
folder locations, approximate cost.

## File Type Handling

| Type | Method | Details |
|------|--------|---------|
| XLS | Convert + Raw XML | Converted to XLSX via Excel COM automation (pywin32), then translated as XLSX |
| XLSX/XLSM | Raw XML | Extracts ZIP, translates sharedStrings.xml and sheet names, updates formula refs, strips `customHeight="1"` from `<row>` elements (so Excel auto-fits row heights to the longer English text), re-zips |
| CSV | Batch text | Detects encoding (utf-8, shift_jis, cp932, euc-jp), translates Japanese cells only |
| PDF | Document API | Preserves layout. PDFs >20 pages are auto-split into chunks, translated, and recombined |
| DOCX | Document API | Preserves formatting |
| PPTX | Document API | Preserves formatting |
| TXT | Plain text | Chunks large files |
| MP4/Audio | Whisper API + GPT | Extracts audio, transcribes via OpenAI Whisper, translates via GPT-4o-mini. Outputs timestamped JA + EN .txt files |
| PNG/JPG | Copy as-is | Not translatable |

Filenames translated too, preserving numeric prefixes (e.g., 2.2.1.8.1.1_).

## Japanese Era Date Conversion (和暦 → 西暦)

The GCP Translation API does NOT correctly convert Japanese Imperial era dates to Gregorian years.
For example, `令和07年03月01日` gets mistranslated as "March 1, 2020" instead of the correct "March 1, 2025".

The translation script (`translate_dataroom.py`) includes a built-in `convert_era_dates()` pre-processing
function that automatically converts all era dates to Gregorian format in the Japanese source text
BEFORE sending it to the GCP API. This runs automatically on every text cell, shared string, and
filename — no manual intervention needed.

**Supported eras and base years:**

| Era | Kanji | Letter | Year 1 = | Example |
|-----|-------|--------|----------|---------|
| Reiwa | 令和 | R | 2019 | 令和7年 = 2025 |
| Heisei | 平成 | H | 1989 | 平成31年 = 2019 |
| Showa | 昭和 | S | 1926 | 昭和64年 = 1989 |
| Taisho | 大正 | T | 1912 | 大正15年 = 1926 |
| Meiji | 明治 | M | 1868 | 明治45年 = 1912 |

**Supported formats:**
- Full kanji: `令和07年03月01日`, `令和7年3月1日`, `平成31年4月`
- Abbreviated: `R07.03.01`, `R7.3.1`, `H31/04/30`

**Important**: This only applies to text-based translations (CSV, Excel shared strings, TXT, filenames).
PDF and DOCX files go through the GCP Document Translation API which handles content as a binary blob —
era dates in those files cannot be pre-processed and may still have incorrect years.

## Row-Height Auto-Fit for Translated Excel Files

Japanese is information-dense — kanji compress meaning, so each character roughly carries a
syllable or short word. English translations typically expand **3–4× in character count**
(e.g., a 47-char Japanese cell becomes a 189-char English cell).

Source spreadsheets laid out for compact Japanese frequently lock row heights with
`<row ht="15" customHeight="1">`. When wrapped English text needs 3–5 lines but the row
stays 15pt tall, the overflow is invisible below the row boundary — making translated files
hard to read.

**Fix**: `translate_dataroom.py` strips `customHeight="1"` (and `customHeight="true"`) from
every `<row>` element inside `xl/worksheets/*.xml` during XLSX/XLSM translation. Excel then
auto-fits row heights when the file is opened. The `ht="..."` hint is preserved as a starting
height; Excel grows rows as needed.

**Scope**: Per-row locks only. Sheet-level `<sheetFormatPr customHeight="1"/>` defaults,
column widths, and merged-cell ranges are not touched.

**Verify before run**: confirm the regex `<row\b ... customHeight="(?:1|true)"` is present in
the worksheet-rewrite block of `translate_xlsx_raw()`, and that the worksheet loop runs
unconditionally (not gated on `name_map`).

**Remediate after run**: if any translated XLSX/XLSM still has locked rows, run
`fix_row_heights.py` against the English folder. It is idempotent — files without locks are
no-ops, and files open in Excel are reported as LOCKED for the user to close and re-run.

## Resumability

The script writes .translation_state.json tracking completed files. If interrupted,
re-running skips completed files and picks up where it left off.

## Troubleshooting

- **429 RESOURCE_EXHAUSTED**: API quota hit. Wait and re-run (resumability handles it).
- **Timer not opening**: Use powershell Start-Process instead of start cmd /k.
- **Empty files**: Document API issue with small files. Originals copied as fallback.
- **OOM**: Files >20MB copied as-is with warning.
- **Translated XLSX is hard to read / text overflows below row boundary**: Row heights were
  locked in the source file (`<row ht="15" customHeight="1">`) and English expanded 3-4× past
  what fits. Run `fix_row_heights.py` against the English folder. If the file is open in Excel,
  close it first (the fixer will report it as LOCKED and skip).
- **Excel COM hangs on .xls conversion** (no error, no progress): kill Excel
  (`Stop-Process -Name EXCEL -Force`) and use the pure-Python fallback (xlrd + openpyxl) — see
  Step 2 Tier 2. Common when paths contain deeply-nested Japanese characters.
- **`zipfile.BadZipFile: File is not a zip file` on .xlsx**: file is encrypted/password-protected.
  Decrypt first via Excel COM with password (try `RENGA2025` for K-Link files), save as new
  password-less xlsx, then re-translate.
- **`FileNotFoundError` writing into Drive Streaming**: Drive Streaming filesystem doesn't
  support some atomic operations Python's zipfile uses. Always copy source files to local temp
  (e.g. `C:\Users\zacha\AppData\Local\Temp\<job>`), translate there, then copy results back.
- **Windows 260-char path limit on long Drive paths**: file copy fails with "cannot find path"
  even when source/dest exist. Mitigations:
  1. Use `\\?\` long-path prefix in Python (`shutil.copy2(src, '\\\\?\\' + dst)`)
  2. Truncate filename if path > 255 chars (preserve extension)
  3. As last resort, use hash-based names and write a `_hash_filename_mapping.txt` so the
     user can identify which file is which
- **Drive folder path resolution issues** (Get-ChildItem says "path not found" for paths with
  Japanese characters): use Python's `os.listdir`/`os.walk` directly, not PowerShell. Python's
  Unicode handling is more reliable on Drive Streaming paths.

## Working with Drive Streaming (G: or H: drive)

When the source folder lives in `C:\Users\<user>\Google Drive Streaming\Shared drives\...` or
similar streaming-mount path, the workflow MUST use a local temp copy because:
1. Drive Streaming does not support some atomic file operations (zipfile writes fail)
2. Path length is often near 260 chars, and the Drive Streaming layer adds overhead
3. Files may not be fully synced/cached when the script tries to read them

**Standard pattern**:
```python
# 1. Copy source to local temp (preserves structure)
local = r"C:\Users\zacha\AppData\Local\Temp\<job_name>"
os.makedirs(os.path.join(local, "Japanese"), exist_ok=True)
os.makedirs(os.path.join(local, "English"), exist_ok=True)
shutil.copytree(drive_jp_folder, local + "/Japanese", dirs_exist_ok=True)

# 2. Run translation against local
# ... translate_dataroom.py local/Japanese local/English

# 3. Deploy results back to Drive (with long-path handling for failures)
for root, dirs, files in os.walk(local + "/English"):
    for f in files:
        # Try direct copy; on path-too-long, try \\?\ prefix; on still-fail, truncate name
        ...
```

## Cross-Machine Setup

This skill folder (`~/.claude/skills/translate-dataroom/`) is portable. To use it on another
machine or under a different Claude account login:

### One-time machine setup
1. **Sync the skill folder** to `~/.claude/skills/translate-dataroom/` on the new machine. Options:
   - Copy/zip-sync the folder manually
   - Symlink to a OneDrive/iCloud/Dropbox path (e.g. point `~/.claude/skills/translate-dataroom`
     at `OneDrive/claude-skills/translate-dataroom`)
   - Put it in a private git repo: `cd ~/.claude/skills && git clone <repo>`

2. **Install Python 3.12+** if not present. Find it: `where python` (Windows) or `which python3`
   (macOS/Linux). Note the absolute path — you'll pass it as `<python>` in commands.

3. **Install Python packages**:
   ```
   "<python>" -m pip install google-cloud-translate PyPDF2 pywin32 openai pydub imageio-ffmpeg xlrd openpyxl
   ```
   (`pywin32` only needed on Windows for Excel COM automation; the xlrd fallback works
   cross-platform.)

4. **Place the GCP service account JSON** somewhere accessible. Pass its absolute path via
   `GOOGLE_APPLICATION_CREDENTIALS` env var on each run.

5. **Verify the service account** has the Cloud Translation API enabled on its GCP project.

### Per-run config the skill needs

When a user invokes this skill, confirm these are known before launching the script:
- Absolute path to Python (`where python`)
- Absolute path to GCP service account JSON
- GCP project ID (from the JSON's `project_id` field)
- Source folder path (always ask the user)
- (Optional) OpenAI API key for MP4 transcription

If any of these are unset on a new machine, ask the user before running. The defaults at the
top of this file are for `zach@broadbandcap.com` on his primary Windows machine — they will
not work elsewhere.

## Lessons Learned (running history)

These are real edge cases observed during prior runs. Refer back when something looks similar.

| Date | Folder | Lesson |
|------|--------|--------|
| 2026-04-13 | `K Link Data Room` | OneDrive sync was incomplete during first scan. Re-scan after sync settles, or assume more files than initial count shows. |
| 2026-04-13 | `K Link Data Room` | Excel COM hung on Drive Streaming paths (`Workbooks.Open` with `DisplayAlerts=False` waits silently for password prompt). Always set `AutomationSecurity=3` and try password fallback. |
| 2026-04-22 | `20260421開示` | First conversion attempt deleted the original `.xls` without producing `.xlsx` (Excel COM `SaveAs` failed silently, `Remove-Item` still ran). Recovered from Drive Trash. **Never `Remove-Item $src` until you've verified `Test-Path $dst`.** |
| 2026-04-22 | `20260421開示` | EML files aren't natively handled by the script. Custom handler: parse with `email.message_from_bytes`, translate `subject` + each `text/plain`/`text/html` part, reconstruct. |
| 2026-04-22 | `20260421開示` | PPTX falls into "Unknown type" branch in current `process_file`. Workaround: send through Document API directly via `client.translate_document(... mime_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")`. |
| 2026-04-22 | `20260421開示` | Drive Streaming path > 260 chars → `shutil.copy2` failure even with valid source. Truncated filename to fit. |
| 2026-04-23 | `Other diligence` | Excel COM hung on 21 .xls files, even with retries and per-file Excel instances. Switched to pure-Python (xlrd + openpyxl) and converted all 21 in seconds. **Don't sink time into debugging COM hangs — fall back fast.** |
| 2026-04-23 | `Other diligence` | Drive Streaming + 260-char limit caused 14 deploy failures. Hash-based fallback names + `_hash_filename_mapping.txt` works but is ugly. Better: when planning output paths, calculate length budget upfront and warn user before writing. |
| 2026-04-23 | `Other diligence` | One xlsx encrypted with a password OTHER than `RENGA2025`. Copy as-is and surface to user — don't burn time guessing. |

## K-Link specific

The K-Link data room in `Renga Partners\05_M&A & Investments\01. Information Memorandums (IMs)\90) 株式会社ケーリンク (K-link)\03. DD\` uses **`RENGA2025`** as the default password
on encrypted Excel files. Try this first when `Workbooks.Open` raises a password error or
`zipfile.BadZipFile` on a `.xlsx`.
