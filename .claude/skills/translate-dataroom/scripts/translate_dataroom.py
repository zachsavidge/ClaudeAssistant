#!/usr/bin/env python3
"""
translate_dataroom.py
====================
Translate Japanese data room documents to English using Google Cloud Translation API v3.

Uses raw XML manipulation for XLSX/XLSM to perfectly preserve formulas, pivot tables, and charts.

Usage:
    set GOOGLE_APPLICATION_CREDENTIALS=C:\\path\\to\\key.json
    set GCP_PROJECT_ID=your-project-id
    python -u translate_dataroom.py INPUT_DIR OUTPUT_DIR [--test N]
"""

import _cloud_creds  # noqa: F401  -- bootstraps GCP creds for cloud sandbox

import os
import sys
import re
import io
import csv
import json
import time
import shutil
import zipfile
import traceback
import tempfile
import argparse
from pathlib import Path
from datetime import datetime

from google.cloud import translate_v3 as translate
from PyPDF2 import PdfReader, PdfWriter

# ============================================================
# JAPANESE ERA DATE CONVERSION
# ============================================================
# GCP Translation API does not correctly convert Japanese Imperial era
# dates (和暦) to Gregorian years. For example, 令和07年 becomes "2020"
# instead of the correct "2025". This pre-processing step converts all
# era dates to Gregorian format in the Japanese source text BEFORE
# translation, so GCP handles them correctly.
#
# Supported eras:
#   令和 (Reiwa)  - 2019+  (year 1 = 2019)
#   平成 (Heisei) - 1989+  (year 1 = 1989)
#   昭和 (Showa)  - 1926+  (year 1 = 1926)
#   大正 (Taisho) - 1912+  (year 1 = 1912)
#   明治 (Meiji)  - 1868+  (year 1 = 1868)
#
# Supported formats:
#   令和07年03月01日    (full kanji with zero-padded numbers)
#   令和7年3月1日       (full kanji, no zero-padding)
#   R07.03.01 / R7.3.1  (abbreviated letter + dots)
#   R07/03/01 / R7/3/1  (abbreviated letter + slashes)

ERA_MAP = {
    '令和': 2018,  # 令和1年 = 2019 = 2018 + 1
    '平成': 1988,  # 平成1年 = 1989 = 1988 + 1
    '昭和': 1925,  # 昭和1年 = 1926 = 1925 + 1
    '大正': 1911,  # 大正1年 = 1912 = 1911 + 1
    '明治': 1867,  # 明治1年 = 1868 = 1867 + 1
}

ERA_ABBREV_MAP = {
    'R': 2018,  # Reiwa
    'H': 1988,  # Heisei
    'S': 1925,  # Showa
    'T': 1911,  # Taisho
    'M': 1867,  # Meiji
}


def _era_kanji_replace(match):
    """Replace full kanji era date: 令和07年03月01日 -> 2025年03月01日"""
    era_name = match.group(1)
    era_year = int(match.group(2))
    rest = match.group(3)  # e.g. 月01日 or 月1日 or empty
    base = ERA_MAP.get(era_name)
    if base is None:
        return match.group(0)
    gregorian_year = base + era_year
    return f"{gregorian_year}年{rest}" if rest else f"{gregorian_year}年"


def _era_abbrev_replace(match):
    """Replace abbreviated era date: R07.03.01 -> 2025.03.01"""
    letter = match.group(1).upper()
    era_year = int(match.group(2))
    sep = match.group(3)
    rest = match.group(4)  # month.day or month/day portion
    base = ERA_ABBREV_MAP.get(letter)
    if base is None:
        return match.group(0)
    gregorian_year = base + era_year
    return f"{gregorian_year}{sep}{rest}"


def convert_era_dates(text):
    """Convert Japanese Imperial era dates to Gregorian in source text.

    Must be called on Japanese text BEFORE sending to translation API.
    """
    if not text:
        return text

    # Full kanji format: 令和07年03月01日, 令和7年3月, 平成31年04月30日
    # Pattern: (era_name)(1-2 digit year)年(rest including month/day if present)
    era_names = '|'.join(ERA_MAP.keys())
    # Match era + year + 年, optionally followed by month/day
    text = re.sub(
        rf'({era_names})(\d{{1,2}})年(\d{{1,2}}月\d{{1,2}}日|(?=\d{{1,2}}月)|\d{{1,2}}月|)',
        _era_kanji_replace,
        text
    )

    # Abbreviated format: R07.03.01, R7.3.1, H31/04/30
    # Only match when followed by digit separators (not mid-word)
    text = re.sub(
        r'(?<![A-Za-z])([RrHhSsTtMm])(\d{1,2})([./])(\d{1,2}[./]\d{1,2})',
        _era_abbrev_replace,
        text
    )

    return text


# ============================================================
# CONFIGURATION
# ============================================================

PROJECT_ID = os.environ.get("GCP_PROJECT_ID", "concise-faculty-492916-r8")
LOCATION = "us-central1"
SOURCE_LANG = "ja"
TARGET_LANG = "en"
MAX_CHARS_PER_REQUEST = 29000
MAX_PDF_PAGES = 20
RATE_LIMIT_DELAY = 0.3

client = translate.TranslationServiceClient()
PARENT = f"projects/{PROJECT_ID}/locations/{LOCATION}"

stats = {"translated": 0, "copied": 0, "failed": 0, "skipped": 0, "total": 0}
translation_cache = {}


# ============================================================
# STATE MANAGEMENT (for resumability)
# ============================================================

def load_state(output_dir):
    state_file = os.path.join(output_dir, ".translation_state.json")
    if os.path.exists(state_file):
        try:
            with open(state_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {"completed": [], "started_at": datetime.now().isoformat()}


def save_state(output_dir, state):
    state_file = os.path.join(output_dir, ".translation_state.json")
    state["updated_at"] = datetime.now().isoformat()
    try:
        with open(state_file, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2)
    except:
        pass


def write_completion_marker(output_dir, stats_dict):
    marker = os.path.join(output_dir, ".translation_complete")
    with open(marker, 'w', encoding='utf-8') as f:
        total = stats_dict['translated'] + stats_dict['copied'] + stats_dict['failed']
        f.write(f"   Translation completed successfully!\n")
        f.write(f"   {stats_dict['translated']}/{total} files translated, {stats_dict['failed']} failures\n")


# ============================================================
# TRANSLATION API HELPERS
# ============================================================

def has_japanese(text):
    if not text:
        return False
    return any(
        '\u3000' <= c <= '\u9fff' or '\uff00' <= c <= '\uffef'
        for c in str(text)
    )


def translate_text(text):
    if not text or not text.strip() or not has_japanese(text):
        return text
    text = convert_era_dates(text)
    cache_key = text.strip()
    if cache_key in translation_cache:
        return translation_cache[cache_key]
    try:
        time.sleep(RATE_LIMIT_DELAY)
        response = client.translate_text(
            request={
                "parent": PARENT,
                "contents": [text],
                "mime_type": "text/plain",
                "source_language_code": SOURCE_LANG,
                "target_language_code": TARGET_LANG,
            }
        )
        result = response.translations[0].translated_text
        translation_cache[cache_key] = result
        return result
    except Exception as e:
        print(f"    WARN translate_text failed: {e}")
        return text


def translate_text_batch(texts):
    if not texts:
        return texts
    results = list(texts)
    to_translate = []
    indices = []
    for i, t in enumerate(texts):
        if t and str(t).strip():
            t_converted = convert_era_dates(str(t))
            results[i] = t_converted  # update with era-converted text
            cache_key = t_converted.strip()
            if cache_key in translation_cache:
                results[i] = translation_cache[cache_key]
            elif has_japanese(t_converted):
                to_translate.append(t_converted)
                indices.append(i)
    if not to_translate:
        return results
    batch_size = 128
    for start in range(0, len(to_translate), batch_size):
        batch = to_translate[start:start + batch_size]
        try:
            time.sleep(RATE_LIMIT_DELAY)
            response = client.translate_text(
                request={
                    "parent": PARENT,
                    "contents": batch,
                    "mime_type": "text/plain",
                    "source_language_code": SOURCE_LANG,
                    "target_language_code": TARGET_LANG,
                }
            )
            for j, t in enumerate(response.translations):
                idx = indices[start + j]
                results[idx] = t.translated_text
                translation_cache[batch[j].strip()] = t.translated_text
        except Exception as e:
            print(f"    WARN batch translate failed: {e}")
    return results


def translate_filename(name):
    if not name:
        return name
    stem, ext = os.path.splitext(name)
    if not has_japanese(stem):
        return name
    prefix_match = re.match(r'^([\d.]+_)', stem)
    prefix = ""
    rest = stem
    if prefix_match:
        prefix = prefix_match.group(1)
        rest = stem[len(prefix):]
    translated = translate_text(rest)
    for ch in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        translated = translated.replace(ch, '_')
    translated = re.sub(r'[_\s]+', ' ', translated).strip()
    return prefix + translated + ext


def translate_dirname(name):
    if not name or not has_japanese(name):
        return name
    prefix_match = re.match(r'^([\d.]+[_\s]*)', name)
    prefix = ""
    rest = name
    if prefix_match:
        prefix = prefix_match.group(1)
        rest = name[len(prefix):]
    translated = translate_text(rest)
    for ch in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        translated = translated.replace(ch, '_')
    translated = re.sub(r'[_\s]+', ' ', translated).strip()
    return prefix + translated


# ============================================================
# XML HELPERS
# ============================================================

def xml_encode(text):
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    return text


def xml_decode(text):
    text = text.replace('&quot;', '"')
    text = text.replace('&apos;', "'")
    text = text.replace('&gt;', '>')
    text = text.replace('&lt;', '<')
    text = text.replace('&amp;', '&')
    return text


def needs_formula_quoting(name):
    if not name:
        return False
    if name[0].isdigit():
        return True
    special = set(' \u3000()><!@#$%^*+=-{}|;,.')
    return bool(set(name) & special)


def make_formula_ref(name_xml_encoded):
    decoded = xml_decode(name_xml_encoded)
    if needs_formula_quoting(decoded):
        return f"'{name_xml_encoded}'!"
    return f"{name_xml_encoded}!"


def replace_sheet_refs(text, name_map):
    result = text
    for old, new in sorted(name_map.items(), key=lambda x: -len(x[0])):
        new_ref = make_formula_ref(new)
        result = result.replace(f"'{old}'!", new_ref)
        result = re.sub(r"(?<!')" + re.escape(old) + r"!", new_ref, result)
    return result


# ============================================================
# XLSX/XLSM RAW XML TRANSLATION
# ============================================================

def translate_xlsx_raw(input_path, output_path):
    temp_dir = tempfile.mkdtemp()

    try:
        with zipfile.ZipFile(input_path, 'r') as zf:
            zf.extractall(temp_dir)

        wb_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
        with open(wb_path, 'r', encoding='utf-8') as f:
            wb_xml = f.read()

        sheet_pattern = r'<sheet\s[^>]*name="([^"]*)"'
        sheet_names_encoded = re.findall(sheet_pattern, wb_xml)

        name_map = {}

        old_decoded = [xml_decode(n) for n in sheet_names_encoded]
        jp_sheets = [(i, n) for i, n in enumerate(old_decoded) if has_japanese(n)]

        if jp_sheets:
            texts = [n for _, n in jp_sheets]
            translated = translate_text_batch(texts)
            for k, (i, _) in enumerate(jp_sheets):
                new_decoded = translated[k]
                for ch in ['/', '\\', '?', '*', '[', ']', ':']:
                    new_decoded = new_decoded.replace(ch, ' ')
                new_decoded = new_decoded.strip()[:31].strip()
                if not new_decoded:
                    new_decoded = f"Sheet{i+1}"
                name_map[sheet_names_encoded[i]] = xml_encode(new_decoded)

        print(f"    Translating {len(name_map)} sheet names...")

        for old_enc, new_enc in name_map.items():
            wb_xml = wb_xml.replace(f'name="{old_enc}"', f'name="{new_enc}"')

        def _replace_dn(match):
            return match.group(1) + replace_sheet_refs(match.group(2), name_map) + match.group(3)

        wb_xml = re.sub(
            r'(<definedName[^>]*>)(.*?)(</definedName>)',
            _replace_dn, wb_xml, flags=re.DOTALL
        )
        with open(wb_path, 'w', encoding='utf-8') as f:
            f.write(wb_xml)

        app_path = os.path.join(temp_dir, 'docProps', 'app.xml')
        if os.path.exists(app_path):
            with open(app_path, 'r', encoding='utf-8') as f:
                app_xml = f.read()
            for old_enc, new_enc in name_map.items():
                app_xml = app_xml.replace(f'>{old_enc}<', f'>{new_enc}<')
            with open(app_path, 'w', encoding='utf-8') as f:
                f.write(app_xml)

        ws_dir = os.path.join(temp_dir, 'xl', 'worksheets')
        if os.path.isdir(ws_dir):
            for ws_file in sorted(os.listdir(ws_dir)):
                if not ws_file.endswith('.xml'):
                    continue
                ws_path = os.path.join(ws_dir, ws_file)
                with open(ws_path, 'r', encoding='utf-8') as f:
                    ws_xml = f.read()

                # Update sheet-name references inside formulas if any sheets were renamed
                if name_map:
                    def _replace_f(match):
                        return (
                            match.group(1)
                            + replace_sheet_refs(match.group(2), name_map)
                            + match.group(3)
                        )

                    ws_xml = re.sub(
                        r'(<f(?:\s[^>]*)?>)(.*?)(</f>)',
                        _replace_f, ws_xml, flags=re.DOTALL
                    )

                # Strip customHeight="1" / "true" from <row> elements so Excel
                # auto-fits row heights to the wrapped translated English text.
                # English typically expands 3-4x vs. Japanese, and a locked row
                # height (ht="15" customHeight="1") hides the wrapped overflow.
                # The ht="..." hint is preserved; Excel grows rows as needed on
                # open. Sheet-level <sheetFormatPr> defaults are left untouched.
                ws_xml = re.sub(
                    r'(<row\b[^>]*?)\s+customHeight="(?:1|true)"',
                    r'\1',
                    ws_xml,
                )

                with open(ws_path, 'w', encoding='utf-8') as f:
                    f.write(ws_xml)

        pivot_dir = os.path.join(temp_dir, 'xl', 'pivotCache')
        if os.path.isdir(pivot_dir):
            for pf in os.listdir(pivot_dir):
                if not pf.endswith('.xml'):
                    continue
                pp = os.path.join(pivot_dir, pf)
                with open(pp, 'r', encoding='utf-8') as f:
                    pxml = f.read()
                for old_enc, new_enc in name_map.items():
                    pxml = pxml.replace(old_enc, new_enc)
                with open(pp, 'w', encoding='utf-8') as f:
                    f.write(pxml)

        ss_path = os.path.join(temp_dir, 'xl', 'sharedStrings.xml')
        if os.path.exists(ss_path):
            with open(ss_path, 'r', encoding='utf-8') as f:
                ss_xml = f.read()

            ss_xml = re.sub(r'<rPh\s[^>]*>.*?</rPh>', '', ss_xml, flags=re.DOTALL)

            t_pattern = r'(<t(?:\s[^>]*)?>)(.*?)(</t>)'
            t_matches = list(re.finditer(t_pattern, ss_xml, flags=re.DOTALL))

            texts_to_translate = []
            match_indices = []
            for i, m in enumerate(t_matches):
                decoded = xml_decode(m.group(2))
                if has_japanese(decoded):
                    texts_to_translate.append(decoded)
                    match_indices.append(i)

            if texts_to_translate:
                print(f"    Translating {len(texts_to_translate)} shared strings...")
                translated = translate_text_batch(texts_to_translate)

                replacements = []
                for k, idx in enumerate(match_indices):
                    m = t_matches[idx]
                    new_text_enc = xml_encode(translated[k])
                    new_full = m.group(1) + new_text_enc + m.group(3)
                    replacements.append((m.start(), m.end(), new_full))

                for start, end, new_full in reversed(replacements):
                    ss_xml = ss_xml[:start] + new_full + ss_xml[end:]

            with open(ss_path, 'w', encoding='utf-8') as f:
                f.write(ss_xml)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(temp_dir):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    arcname = os.path.relpath(fpath, temp_dir)
                    zf.write(fpath, arcname)

        shutil.rmtree(temp_dir)
        return True

    except Exception as e:
        print(f"    ERROR in XLSX translation: {e}")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        shutil.copy2(input_path, output_path)
        return False


# ============================================================
# CSV TRANSLATION
# ============================================================

def translate_csv_file(input_path, output_path):
    try:
        content = None
        for enc in ['utf-8', 'shift_jis', 'cp932', 'euc-jp', 'iso-2022-jp']:
            try:
                with open(input_path, 'r', encoding=enc) as f:
                    content = f.read()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        if content is None:
            print(f"    WARN: Could not decode CSV, copying as-is")
            shutil.copy2(input_path, output_path)
            return False

        sniffer = csv.Sniffer()
        try:
            dialect = sniffer.sniff(content[:2000])
            delimiter = dialect.delimiter
        except csv.Error:
            delimiter = ','

        reader = csv.reader(io.StringIO(content), delimiter=delimiter)
        rows = list(reader)
        if not rows:
            shutil.copy2(input_path, output_path)
            return True

        all_cells = []
        cell_positions = []
        for ri, row in enumerate(rows):
            for ci, cell in enumerate(row):
                cell = cell.strip()
                if cell and has_japanese(cell):
                    all_cells.append(cell)
                    cell_positions.append((ri, ci))

        if all_cells:
            print(f"    Translating {len(all_cells)} Japanese cells...")
            translated_cells = translate_text_batch(all_cells)
            for k, (ri, ci) in enumerate(cell_positions):
                rows[ri][ci] = translated_cells[k]

        output_buf = io.StringIO()
        writer = csv.writer(output_buf, delimiter=delimiter, lineterminator='\n')
        writer.writerows(rows)
        with open(output_path, 'w', encoding='utf-8', newline='') as f:
            f.write(output_buf.getvalue())
        return True

    except Exception as e:
        print(f"    WARN CSV translation error: {e}")
        traceback.print_exc()
        shutil.copy2(input_path, output_path)
        return False


# ============================================================
# PDF / DOCX DOCUMENT TRANSLATION
# ============================================================

def translate_document(input_path, output_path, mime_type):
    with open(input_path, "rb") as f:
        content = f.read()
    if len(content) == 0:
        shutil.copy2(input_path, output_path)
        return True
    if len(content) > 20 * 1024 * 1024:
        print(f"    WARN: File >20MB, copying as-is")
        shutil.copy2(input_path, output_path)
        return False
    try:
        time.sleep(RATE_LIMIT_DELAY * 3)
        response = client.translate_document(
            request={
                "parent": PARENT,
                "source_language_code": SOURCE_LANG,
                "target_language_code": TARGET_LANG,
                "document_input_config": {
                    "content": content,
                    "mime_type": mime_type,
                },
            }
        )
        translated_bytes = response.document_translation.byte_stream_outputs[0]
        with open(output_path, "wb") as f:
            f.write(translated_bytes)
        return True
    except Exception as e:
        print(f"    WARN document translation failed: {e}")
        shutil.copy2(input_path, output_path)
        return False


def translate_pdf_chunked(input_path, output_path):
    """Split a large PDF into <=MAX_PDF_PAGES chunks, translate each, and merge."""
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    num_chunks = (total_pages + MAX_PDF_PAGES - 1) // MAX_PDF_PAGES
    print(f"    PDF has {total_pages} pages (>{MAX_PDF_PAGES}), splitting into {num_chunks} chunks...")

    chunk_paths = []
    translated_chunk_paths = []
    try:
        # Split into chunks
        for start in range(0, total_pages, MAX_PDF_PAGES):
            end = min(start + MAX_PDF_PAGES, total_pages)
            writer = PdfWriter()
            for i in range(start, end):
                writer.add_page(reader.pages[i])
            tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            writer.write(tmp)
            tmp.close()
            chunk_paths.append(tmp.name)

        # Translate each chunk
        for i, chunk_path in enumerate(chunk_paths):
            print(f"    Translating chunk {i+1}/{num_chunks}...")
            with open(chunk_path, 'rb') as f:
                content = f.read()
            try:
                time.sleep(RATE_LIMIT_DELAY * 3)
                response = client.translate_document(
                    request={
                        "parent": PARENT,
                        "source_language_code": SOURCE_LANG,
                        "target_language_code": TARGET_LANG,
                        "document_input_config": {
                            "content": content,
                            "mime_type": "application/pdf",
                        },
                    }
                )
                translated_bytes = response.document_translation.byte_stream_outputs[0]
                out_path = chunk_path.replace('.pdf', '_translated.pdf')
                with open(out_path, 'wb') as f:
                    f.write(translated_bytes)
                translated_chunk_paths.append(out_path)
                print(f"    Chunk {i+1} OK")
            except Exception as e:
                print(f"    WARN chunk {i+1} failed: {e}, using original")
                translated_chunk_paths.append(chunk_path)

        # Merge translated chunks
        print(f"    Merging {len(translated_chunk_paths)} chunks...")
        merger = PdfWriter()
        for p in translated_chunk_paths:
            chunk_reader = PdfReader(p)
            for page in chunk_reader.pages:
                merger.add_page(page)
        with open(output_path, 'wb') as f:
            merger.write(f)

        return True
    except Exception as e:
        print(f"    ERROR in chunked PDF translation: {e}")
        traceback.print_exc()
        shutil.copy2(input_path, output_path)
        return False
    finally:
        # Cleanup temp files
        for p in chunk_paths + translated_chunk_paths:
            try:
                os.unlink(p)
            except:
                pass


# ============================================================
# TEXT FILE TRANSLATION
# ============================================================

def translate_text_file(input_path, output_path):
    try:
        content = None
        for enc in ['utf-8', 'shift_jis', 'cp932', 'euc-jp']:
            try:
                with open(input_path, 'r', encoding=enc) as f:
                    content = f.read()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        if not content or not content.strip():
            shutil.copy2(input_path, output_path)
            return True

        chunks = []
        current = ""
        for line in content.split("\n"):
            if len(current) + len(line) + 1 > MAX_CHARS_PER_REQUEST:
                chunks.append(current)
                current = line
            else:
                current = current + "\n" + line if current else line
        if current:
            chunks.append(current)

        translated = [translate_text(c) for c in chunks]
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(translated))
        return True
    except Exception as e:
        print(f"    WARN text translation error: {e}")
        shutil.copy2(input_path, output_path)
        return False


# ============================================================
# FILE ROUTER
# ============================================================

def process_file(input_path, output_path, skip_pdf=False, pdf_only=False):
    stats["total"] += 1
    filename = os.path.basename(input_path)
    ext = os.path.splitext(filename)[1].lower()

    print(f"\n[{stats['total']}] {input_path}")
    print(f"  -> {output_path}")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Scope filtering
    if skip_pdf and ext == '.pdf':
        print(f"  Skipping PDF (--skip-pdf), copying as-is...")
        shutil.copy2(input_path, output_path)
        stats["copied"] += 1
        return True

    if pdf_only and ext != '.pdf':
        print(f"  Skipping non-PDF (--pdf-only), copying as-is...")
        shutil.copy2(input_path, output_path)
        stats["copied"] += 1
        return True

    success = False

    if ext in ('.xlsx', '.xlsm'):
        print(f"  Translating Excel (raw XML)...")
        success = translate_xlsx_raw(input_path, output_path)

    elif ext == '.csv':
        print(f"  Translating CSV...")
        success = translate_csv_file(input_path, output_path)

    elif ext == '.pdf':
        print(f"  Translating PDF...")
        try:
            page_count = len(PdfReader(input_path).pages)
        except:
            page_count = 0
        if page_count > MAX_PDF_PAGES:
            success = translate_pdf_chunked(input_path, output_path)
        else:
            success = translate_document(input_path, output_path, "application/pdf")

    elif ext == '.docx':
        print(f"  Translating DOCX...")
        success = translate_document(
            input_path, output_path,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    elif ext == '.txt':
        print(f"  Translating text...")
        success = translate_text_file(input_path, output_path)

    else:
        print(f"  Unknown type '{ext}', copying as-is...")
        shutil.copy2(input_path, output_path)
        stats["copied"] += 1
        return True

    if success:
        stats["translated"] += 1
        print(f"  OK")
    else:
        stats["failed"] += 1
        print(f"  FAILED (original copied)")

    return success


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Translate Japanese data room documents to English"
    )
    parser.add_argument("input_dir", help="Directory with Japanese files (walked recursively)")
    parser.add_argument("output_dir", help="Output directory for translations")
    parser.add_argument(
        "--test", type=int, default=0,
        help="Only process first N files (for testing)"
    )
    scope = parser.add_mutually_exclusive_group()
    scope.add_argument(
        "--skip-pdf", action="store_true",
        help="Translate everything EXCEPT PDFs (PDFs copied as-is)"
    )
    scope.add_argument(
        "--pdf-only", action="store_true",
        help="Translate ONLY PDF files (skip all other types)"
    )
    args = parser.parse_args()

    input_dir = os.path.normpath(os.path.expanduser(args.input_dir))
    output_dir = os.path.normpath(os.path.expanduser(args.output_dir))

    print("=" * 60)
    print("Data Room Translation - Japanese to English")
    print(f"Project:  {PROJECT_ID}")
    print(f"Input:    {input_dir}")
    print(f"Output:   {output_dir}")
    if args.test:
        print(f"TEST MODE: Processing only {args.test} files")
    if args.skip_pdf:
        print(f"SCOPE: Everything EXCEPT PDFs (PDFs copied as-is)")
    elif args.pdf_only:
        print(f"SCOPE: PDF files ONLY")
    print("=" * 60)

    os.makedirs(output_dir, exist_ok=True)

    state = load_state(output_dir)
    completed_set = set(state.get("completed", []))

    all_files = []
    for root, dirs, files in os.walk(input_dir):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for f in sorted(files):
            if f.startswith('.'):
                continue
            rel_path = os.path.relpath(os.path.join(root, f), input_dir)
            all_files.append(rel_path)

    if args.test:
        all_files = all_files[:args.test]

    remaining = [f for f in all_files if f not in completed_set]
    skipped = len(all_files) - len(remaining)
    stats["skipped"] = skipped

    print(f"\nFound {len(all_files)} total files.")
    if skipped:
        print(f"Skipping {skipped} already-translated files.")
    print(f"Processing {len(remaining)} files.\n")

    for rel_path in remaining:
        input_path = os.path.join(input_dir, rel_path)

        parts = Path(rel_path).parts
        translated_parts = []
        for i, part in enumerate(parts):
            if i == len(parts) - 1:
                translated_parts.append(translate_filename(part))
            else:
                translated_parts.append(translate_dirname(part))

        translated_rel = os.path.join(*translated_parts) if len(translated_parts) > 1 else translated_parts[0]
        output_path = os.path.join(output_dir, translated_rel)

        try:
            process_file(input_path, output_path,
                         skip_pdf=args.skip_pdf, pdf_only=args.pdf_only)
            state.setdefault("completed", []).append(rel_path)
            save_state(output_dir, state)
        except Exception as e:
            print(f"  ERROR: {e}")
            traceback.print_exc()
            stats["failed"] += 1

    write_completion_marker(output_dir, stats)

    print(f"\n{'=' * 60}")
    print("TRANSLATION COMPLETE")
    print(f"  Total files:  {len(all_files)}")
    print(f"  Translated:   {stats['translated']}")
    print(f"  Copied:       {stats['copied']}")
    print(f"  Skipped:      {stats['skipped']}")
    print(f"  Failed:       {stats['failed']}")
    print(f"{'=' * 60}")

    # Cross-platform completion beep
    if sys.platform == "win32":
        try:
            import winsound
            for _ in range(3):
                winsound.Beep(1000, 500)
                time.sleep(0.3)
        except Exception:
            print("\a")
    elif sys.platform == "darwin":
        import subprocess
        for _ in range(3):
            try:
                subprocess.run(
                    ["afplay", "/System/Library/Sounds/Glass.aiff"],
                    timeout=2, check=False,
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
            except Exception:
                print("\a", end="", flush=True)
            time.sleep(0.3)
    else:
        for _ in range(3):
            print("\a", end="", flush=True)
            time.sleep(0.3)


if __name__ == "__main__":
    main()
