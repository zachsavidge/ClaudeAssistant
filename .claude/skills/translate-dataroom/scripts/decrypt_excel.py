#!/usr/bin/env python3
"""
decrypt_excel.py - Cross-platform decryption of password-protected Office files.

Uses msoffcrypto-tool (pure Python) to remove the password from encrypted
.xlsx, .docx, .pptx, and .xls files. Works on Mac, Linux, and Windows.

Strategy:
  1. PRIMARY: msoffcrypto-tool (cross-platform, pure Python)
  2. FALLBACK (Windows only): Excel COM via pywin32

Install: pip install msoffcrypto-tool

Usage:
  python decrypt_excel.py <file>                     # try common passwords
  python decrypt_excel.py <file> --password <pw>     # specific password
  python decrypt_excel.py <folder>                   # walk folder, decrypt all
  python decrypt_excel.py <file> --output <path>     # custom output (default: in-place)

If --output is not given, replaces the original (after writing a backup).
"""

import os
import sys
import argparse
import shutil
import tempfile


# Common passwords to try (project-specific)
DEFAULT_PASSWORDS = [
    "RENGA2025",  # K-Link dataroom convention
]


def decrypt_msoffcrypto(input_path, password, output_path=None):
    """Decrypt with msoffcrypto-tool. Returns (success, message)."""
    try:
        import msoffcrypto
    except ImportError:
        return False, "msoffcrypto-tool not installed (run: pip install msoffcrypto-tool)"

    out = output_path or (input_path + ".decrypted.tmp")
    try:
        with open(input_path, "rb") as fin:
            office_file = msoffcrypto.OfficeFile(fin)
            if not office_file.is_encrypted():
                return False, "not encrypted"
            office_file.load_key(password=password)
            with open(out, "wb") as fout:
                office_file.decrypt(fout)
        return True, f"decrypted with msoffcrypto"
    except msoffcrypto.exceptions.InvalidKeyError:
        if os.path.exists(out):
            os.remove(out)
        return False, "wrong password"
    except Exception as e:
        if os.path.exists(out):
            os.remove(out)
        return False, f"msoffcrypto error: {e}"


def decrypt_excel_com(input_path, password, output_path=None):
    """Decrypt via Excel COM (Windows only)."""
    if sys.platform != "win32":
        return False, "Excel COM only available on Windows"
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return False, "pywin32 not installed"

    out = output_path or (input_path + ".decrypted.tmp")
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_path, 0, False, [None], password)
        wb.Password = ""
        # Save as same format as input (51 for xlsx, 56 for xls)
        ext = os.path.splitext(input_path)[1].lower()
        fmt = 51 if ext == ".xlsx" else 56
        wb.SaveAs(out, FileFormat=fmt)
        wb.Close(False)
        excel.Quit()
        return True, "decrypted with Excel COM"
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


def decrypt_one(path, passwords, output_path=None, use_excel=False):
    """Try each password until one works. Returns (success, password_used, msg)."""
    for pw in passwords:
        # Try msoffcrypto first
        ok, msg = decrypt_msoffcrypto(path, pw, output_path)
        if ok:
            return True, pw, msg
        if "not encrypted" in msg:
            return False, None, "not encrypted"
        if "wrong password" not in msg and use_excel and sys.platform == "win32":
            # msoffcrypto failed for some other reason; try Excel COM
            ok, msg2 = decrypt_excel_com(path, pw, output_path)
            if ok:
                return True, pw, msg2
    return False, None, f"all {len(passwords)} password(s) failed"


def is_office_file(path):
    return path.lower().endswith((".xlsx", ".xlsm", ".xls", ".docx", ".pptx"))


def find_office_files(target):
    if os.path.isfile(target):
        return [target] if is_office_file(target) else []
    out = []
    for root, dirs, files in os.walk(target):
        for f in files:
            if is_office_file(f):
                out.append(os.path.join(root, f))
    return sorted(out)


def main():
    p = argparse.ArgumentParser(description="Cross-platform Office file decryptor")
    p.add_argument("target", help="File or folder")
    p.add_argument("--password", help="Specific password (overrides default list)")
    p.add_argument("--output", help="Output path (default: replace input with backup)")
    p.add_argument("--use-excel", action="store_true",
                   help="Also try Excel COM as fallback (Windows only)")
    p.add_argument("--no-backup", action="store_true",
                   help="Don't create a .bak backup of the original")
    args = p.parse_args()

    passwords = [args.password] if args.password else DEFAULT_PASSWORDS

    files = find_office_files(args.target)
    if not files:
        print(f"No Office files found in {args.target}")
        return

    print(f"Found {len(files)} Office file(s). Platform: {sys.platform}")
    print(f"Passwords to try: {passwords}")
    print()

    ok_count = 0
    not_encrypted = 0
    failed = []
    for i, path in enumerate(files, 1):
        rel = os.path.relpath(path)
        print(f"[{i}/{len(files)}] {rel}", flush=True)

        # Decrypt to a temp file then replace
        out = args.output if (args.output and len(files) == 1) else (path + ".decrypted.tmp")
        success, pw, msg = decrypt_one(path, passwords, out, args.use_excel)
        if success:
            if args.output and len(files) == 1:
                print(f"    OK -> {args.output} (password: {pw}, {msg})")
            else:
                # Replace in place
                if not args.no_backup:
                    bak = path + ".bak"
                    shutil.copy2(path, bak)
                shutil.move(out, path)
                print(f"    OK (in-place, password: {pw}, {msg})")
            ok_count += 1
        elif msg == "not encrypted":
            not_encrypted += 1
            print(f"    skip (not encrypted)")
        else:
            failed.append((rel, msg))
            print(f"    FAILED: {msg}")

    print(f"\n=== Summary ===")
    print(f"Decrypted:     {ok_count}")
    print(f"Not encrypted: {not_encrypted}")
    print(f"Failed:        {len(failed)}")
    for n, m in failed:
        print(f"  - {n}: {m[:100]}")


if __name__ == "__main__":
    main()
