"""
Retranslate remaining Japanese shared strings in an XLSX file.
Uses deep-translator (free Google Translate) as fallback when GCP quota is exhausted.
Operates on raw XML to preserve formulas, formatting, charts, and pivot tables.
"""
import sys, os, re, shutil, zipfile, time, tempfile
import xml.etree.ElementTree as ET
from deep_translator import GoogleTranslator

ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

JP_PATTERN = re.compile(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF\uFF00-\uFFEF]')

# Era date conversion (same as main script)
ERA_MAP = {
    '令和': 2018, '平成': 1988, '昭和': 1925, '大正': 1911, '明治': 1867,
    'R': 2018, 'H': 1988, 'S': 1925, 'T': 1911, 'M': 1867,
}

def convert_era_dates(text):
    if not text:
        return text
    # Full kanji: 令和07年03月01日
    def replace_full(m):
        era, yr, rest = m.group(1), int(m.group(2)), m.group(3)
        western = ERA_MAP.get(era, 0) + yr
        return str(western) + '年' + rest
    text = re.sub(r'(令和|平成|昭和|大正|明治)\s*(\d{1,2})\s*年(.*)$',
                  replace_full, text, flags=re.MULTILINE)
    # Abbreviated: R07.03.01 or R7/3/1
    def replace_abbr(m):
        era, yr, sep, rest = m.group(1), int(m.group(2)), m.group(3), m.group(4)
        western = ERA_MAP.get(era, 0) + yr
        return str(western) + sep + rest
    text = re.sub(r'([RHSTM])(\d{1,2})([./])([\d./]+)', replace_abbr, text)
    return text


def translate_batch(texts, translator, batch_size=40, delay=1.5):
    """Translate a list of texts in batches with rate limiting."""
    results = []
    total = len(texts)
    for i in range(0, total, batch_size):
        batch = texts[i:i+batch_size]
        batch_num = i // batch_size + 1
        total_batches = (total + batch_size - 1) // batch_size
        print(f"  Batch {batch_num}/{total_batches} ({len(batch)} strings)...", flush=True)

        for attempt in range(3):
            try:
                translated = translator.translate_batch(batch)
                results.extend(translated)
                break
            except Exception as e:
                if attempt < 2:
                    wait = (attempt + 1) * 5
                    print(f"    Retry in {wait}s: {e}", flush=True)
                    time.sleep(wait)
                else:
                    print(f"    FAILED, keeping originals: {e}", flush=True)
                    results.extend(batch)

        if i + batch_size < total:
            time.sleep(delay)

    return results


def retranslate_xlsx(xlsx_path, output_path):
    print(f"Input:  {xlsx_path}")
    print(f"Output: {output_path}")

    tmpdir = tempfile.mkdtemp(prefix='xlsx_retrans_')

    try:
        # Extract ZIP
        with zipfile.ZipFile(xlsx_path, 'r') as zf:
            zf.extractall(tmpdir)

        ss_path = os.path.join(tmpdir, 'xl', 'sharedStrings.xml')
        if not os.path.exists(ss_path):
            print("ERROR: No sharedStrings.xml found")
            return False

        # Parse shared strings
        tree = ET.parse(ss_path)
        root = tree.getroot()

        # Collect all string elements and their text
        si_elements = root.findall(f'{{{NS}}}si')
        print(f"Total shared strings: {len(si_elements)}")

        # Find Japanese strings that need translation
        to_translate = []  # (index, text, t_elements)
        for idx, si in enumerate(si_elements):
            t_elements = list(si.iter(f'{{{NS}}}t'))
            full_text = ''.join(t.text or '' for t in t_elements)
            if JP_PATTERN.search(full_text):
                to_translate.append((idx, full_text, t_elements))

        print(f"Strings needing translation: {len(to_translate)}")

        if not to_translate:
            print("Nothing to translate!")
            shutil.copy2(xlsx_path, output_path)
            return True

        # Pre-process era dates
        texts_to_send = []
        for idx, text, _ in to_translate:
            texts_to_send.append(convert_era_dates(text))

        # Translate
        print(f"Translating {len(texts_to_send)} strings via Google Translate...")
        translator = GoogleTranslator(source='ja', target='en')
        translated = translate_batch(texts_to_send, translator)

        # Write translations back into XML
        success_count = 0
        for i, (idx, orig_text, t_elements) in enumerate(to_translate):
            new_text = translated[i] if i < len(translated) else orig_text
            if new_text and new_text != orig_text:
                # If single <t> element, replace directly
                if len(t_elements) == 1:
                    t_elements[0].text = new_text
                else:
                    # Multiple <t> elements (rich text) - put all text in first, clear rest
                    t_elements[0].text = new_text
                    for t in t_elements[1:]:
                        t.text = ''
                success_count += 1

        print(f"Successfully translated: {success_count}/{len(to_translate)}")

        # Write updated XML
        tree.write(ss_path, xml_declaration=True, encoding='UTF-8')

        # Also translate sheet names in workbook.xml
        wb_path = os.path.join(tmpdir, 'xl', 'workbook.xml')
        if os.path.exists(wb_path):
            wb_tree = ET.parse(wb_path)
            wb_root = wb_tree.getroot()
            sheets = wb_root.findall(f'.//{{{NS}}}sheet')
            sheet_names_jp = []
            sheet_indices = []
            for i, sheet in enumerate(sheets):
                name = sheet.get('name', '')
                if JP_PATTERN.search(name):
                    sheet_names_jp.append(convert_era_dates(name))
                    sheet_indices.append(i)

            if sheet_names_jp:
                print(f"Translating {len(sheet_names_jp)} sheet names...")
                translated_names = translate_batch(sheet_names_jp, translator, batch_size=20)
                for j, si in enumerate(sheet_indices):
                    if j < len(translated_names):
                        new_name = translated_names[j][:31]  # Excel 31-char limit
                        sheets[si].set('name', new_name)
                        print(f"  Sheet: {sheet_names_jp[j]} -> {new_name}")
                wb_tree.write(wb_path, xml_declaration=True, encoding='UTF-8')

        # Re-zip, preserving all other files exactly
        print("Re-packaging XLSX...")
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for dirpath, dirnames, filenames in os.walk(tmpdir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    arcname = os.path.relpath(full, tmpdir).replace('\\', '/')
                    zout.write(full, arcname)

        # Verify
        try:
            with zipfile.ZipFile(output_path, 'r') as zf:
                zf.testzip()
            print("ZIP integrity: OK")
        except Exception as e:
            print(f"ZIP integrity: FAILED - {e}")
            return False

        return True

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: retranslate_xlsx_strings.py <input.xlsx> <output.xlsx>")
        sys.exit(1)

    ok = retranslate_xlsx(sys.argv[1], sys.argv[2])
    sys.exit(0 if ok else 1)
