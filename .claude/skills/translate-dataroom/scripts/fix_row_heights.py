#!/usr/bin/env python
"""
fix_row_heights.py
==================
Strip customHeight="1" from <row> elements in every XLSX/XLSM under a folder.

Why: Japanese-laid-out spreadsheets often lock row heights with
`<row ht="15" customHeight="1">`, which makes wrapped English translations
(typically 3-4x longer than the Japanese source) overflow invisibly below
the row boundary. Removing customHeight lets Excel auto-fit on open.

Usage:
    python fix_row_heights.py FOLDER
"""
import os
import re
import sys
import shutil
import tempfile
import zipfile

ROW_PATTERN = re.compile(r'(<row\b[^>]*?)\s+customHeight="(?:1|true)"')


def fix_xlsx(path):
    """Rewrite worksheet XMLs in-place to drop row-level customHeight locks."""
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            zf.extractall(temp_dir)

        ws_dir = os.path.join(temp_dir, 'xl', 'worksheets')
        if not os.path.isdir(ws_dir):
            return 0

        total_stripped = 0
        for name in sorted(os.listdir(ws_dir)):
            if not name.endswith('.xml'):
                continue
            ws_path = os.path.join(ws_dir, name)
            with open(ws_path, 'r', encoding='utf-8') as f:
                xml = f.read()
            new_xml, n = ROW_PATTERN.subn(r'\1', xml)
            if n:
                with open(ws_path, 'w', encoding='utf-8') as f:
                    f.write(new_xml)
                total_stripped += n

        if total_stripped == 0:
            return 0

        # Re-zip back to the original path
        tmp_out = path + '.tmp'
        with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(temp_dir):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    arcname = os.path.relpath(fpath, temp_dir).replace(os.sep, '/')
                    zf.write(fpath, arcname)
        os.replace(tmp_out, path)
        return total_stripped
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    if len(sys.argv) != 2:
        print("Usage: python fix_row_heights.py FOLDER")
        sys.exit(2)
    folder = sys.argv[1]

    files_touched = 0
    files_skipped = 0
    files_locked = 0
    rows_stripped = 0
    for root, dirs, files in os.walk(folder):
        for name in files:
            if not name.lower().endswith(('.xlsx', '.xlsm')):
                continue
            path = os.path.join(root, name)
            try:
                n = fix_xlsx(path)
                rel = os.path.relpath(path, folder)
                if n:
                    print(f"  fixed  {rel}  ({n} rows)")
                    files_touched += 1
                    rows_stripped += n
                else:
                    print(f"  noop   {rel}")
                    files_skipped += 1
            except PermissionError:
                rel = os.path.relpath(path, folder)
                print(f"  LOCKED {rel}  (close it in Excel and re-run)")
                files_locked += 1
            except Exception as e:
                rel = os.path.relpath(path, folder)
                print(f"  ERROR  {rel}: {e}")

    print()
    print(f"Files modified: {files_touched}")
    print(f"Files unchanged: {files_skipped}")
    if files_locked:
        print(f"Files locked (open in Excel): {files_locked}")
    print(f"Total rows unlocked: {rows_stripped}")


if __name__ == '__main__':
    main()
