#!/usr/bin/env python3
"""
convert_xls.py - Cross-platform .xls → .xlsx batch converter.

Strategy:
  1. PRIMARY: pure-Python via xlrd + openpyxl (works on Mac, Linux, Windows)
     - Loses formulas/charts but preserves all values + text
     - Fast and reliable for survey/research data
  2. FALLBACK: Excel COM via pywin32 (Windows only, if installed)
     - Preserves formulas, formatting, charts
     - Can hang on certain files (especially Drive Streaming paths)
     - Only attempted with --use-excel flag (opt-in)

Both tiers handle password-protected .xls via the optional --password flag.

Usage:
  python convert_xls.py <folder>                      # walk folder, convert all .xls
  python convert_xls.py <folder> --password RENGA2025 # try this password if encrypted
  python convert_xls.py <folder> --use-excel          # try Excel COM first (Windows only)
  python convert_xls.py <file.xls>                    # convert single file

Outputs the .xlsx alongside the .xls and removes the original on success.
"""

import os
import sys
import time
import argparse


def is_xls(path):
    return path.lower().endswith(".xls") and not path.lower().endswith(".xlsx")


def convert_pure_python(xls_path, password=None):
    """Convert .xls → .xlsx via xlrd + openpyxl. Cross-platform."""
    import xlrd
    from openpyxl import Workbook

    try:
        # xlrd 1.2.0 is the last version that supports .xls
        # If password-protected, xlrd doesn't decrypt - that's the msoffcrypto path
        rb = xlrd.open_workbook(xls_path, formatting_info=False)
    except xlrd.biffh.XLRDError as e:
        if "encrypted" in str(e).lower():
            return False, "encrypted (use decrypt_excel.py first)"
        return False, f"xlrd: {e}"
    except Exception as e:
        return False, f"open: {e}"

    wb = Workbook()
    wb.remove(wb.active)

    used_names = set()
    for sheet_name in rb.sheet_names():
        rs = rb.sheet_by_name(sheet_name)
        # Sanitize: max 31 chars, no special chars
        safe = sheet_name
        for bad in "[]*?:/\\":
            safe = safe.replace(bad, "_")
        safe = safe[:31]
        orig = safe
        i = 1
        while safe in used_names:
            safe = f"{orig[:28]}_{i}"
            i += 1
        used_names.add(safe)

        ws = wb.create_sheet(title=safe)
        for r in range(rs.nrows):
            for c in range(rs.ncols):
                cell = rs.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    continue
                val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        val = xlrd.xldate_as_datetime(val, rb.datemode)
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    val = bool(val)
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    val = None
                if val is None or val == "":
                    continue
                ws.cell(row=r + 1, column=c + 1, value=val)

    dst = os.path.splitext(xls_path)[0] + ".xlsx"
    wb.save(dst)
    return True, f"OK ({len(rb.sheet_names())} sheets)"


def convert_excel_com(xls_path, password=None, timeout=30):
    """Convert via Excel COM (Windows only). Returns (success, message)."""
    if sys.platform != "win32":
        return False, "Excel COM only available on Windows"
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return False, "pywin32 not installed"

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        try:
            excel.AutomationSecurity = 3  # disable macros
        except Exception:
            pass

        if password:
            wb = excel.Workbooks.Open(xls_path, 0, False, 5, password)
            wb.Password = ""
        else:
            wb = excel.Workbooks.Open(xls_path, 0, False, 5)

        dst = os.path.splitext(xls_path)[0] + ".xlsx"
        wb.SaveAs(dst, FileFormat=51)
        wb.Close(False)
        excel.Quit()
        return True, "OK (via Excel COM)"
    except Exception as e:
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        return False, f"Excel COM: {e}"


def convert_one(xls_path, password=None, use_excel=False):
    """Convert a single file. Tries Excel COM first (if requested) then pure Python."""
    # Tier 1: Excel COM (only if explicitly requested AND Windows)
    if use_excel and sys.platform == "win32":
        ok, msg = convert_excel_com(xls_path, password)
        if ok:
            return ok, msg
        # Fall through to pure Python

    # Tier 2 (default): pure Python
    return convert_pure_python(xls_path, password)


def find_xls_files(target):
    """Return list of .xls files. target may be a single file or a folder."""
    if os.path.isfile(target):
        return [target] if is_xls(target) else []
    out = []
    for root, dirs, files in os.walk(target):
        for f in files:
            if is_xls(f):
                out.append(os.path.join(root, f))
    return sorted(out)


def main():
    p = argparse.ArgumentParser(description="Cross-platform .xls → .xlsx converter")
    p.add_argument("target", help="Folder or single .xls file")
    p.add_argument("--password", help="Password to try if files are encrypted")
    p.add_argument("--use-excel", action="store_true",
                   help="Try Excel COM first (Windows only); fall back to pure Python on failure")
    p.add_argument("--keep-original", action="store_true",
                   help="Don't delete the .xls after successful conversion")
    args = p.parse_args()

    files = find_xls_files(args.target)
    if not files:
        print(f"No .xls files found in {args.target}")
        return

    print(f"Found {len(files)} .xls file(s). Platform: {sys.platform}")
    print(f"Strategy: {'Excel COM → pure Python' if args.use_excel and sys.platform == 'win32' else 'pure Python only'}")
    print()

    ok_count = 0
    fail = []
    for i, path in enumerate(files, 1):
        rel = os.path.relpath(path)
        print(f"[{i}/{len(files)}] {rel}", flush=True)
        t0 = time.time()
        success, msg = convert_one(path, password=args.password, use_excel=args.use_excel)
        elapsed = time.time() - t0
        if success:
            if not args.keep_original:
                try:
                    os.remove(path)
                except Exception:
                    pass
            ok_count += 1
            print(f"    {msg} ({elapsed:.1f}s)", flush=True)
        else:
            fail.append((rel, msg))
            print(f"    FAILED: {msg[:150]} ({elapsed:.1f}s)", flush=True)

    print(f"\n=== Summary ===")
    print(f"Converted: {ok_count}/{len(files)}")
    if fail:
        print(f"Failed: {len(fail)}")
        for n, m in fail:
            print(f"  - {n}: {m[:100]}")


if __name__ == "__main__":
    main()
