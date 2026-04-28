"""Verify the local environment is ready to run translate_dataroom.

Cross-platform — works on Windows, macOS, and Linux. Run before each new
translation job (especially on a new machine):

    python check_setup.py
"""
import _cloud_creds  # noqa: F401  -- bootstraps GCP creds for cloud sandbox

import os
import sys
import importlib
import platform

print("=" * 60)
print("translate-dataroom skill: setup check")
print("=" * 60)

ok = True

# Platform info
print(f"\nPlatform: {sys.platform} ({platform.platform()})")
print(f"Python:   {sys.version.split()[0]}  [{sys.executable}]")
if sys.version_info < (3, 10):
    print("  ⚠️  Recommended: Python 3.10+")

# Cross-platform required packages
required = {
    "google.cloud.translate_v3": "google-cloud-translate",
    "PyPDF2": "PyPDF2",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",
    "msoffcrypto": "msoffcrypto-tool",  # cross-platform decryption
}

# Optional packages (Windows-specific or extra features)
optional = {
    "openai": ("openai", "MP4 transcription"),
    "pydub": ("pydub", "MP4 audio splitting"),
    "imageio_ffmpeg": ("imageio-ffmpeg", "MP4 audio extraction"),
}
windows_only = {
    "win32com.client": ("pywin32", "Excel COM automation (Windows-only fallback)"),
}

print("\nRequired packages:")
for mod, pkg in required.items():
    try:
        importlib.import_module(mod)
        print(f"  ✓ {pkg}")
    except ImportError:
        print(f"  ✗ {pkg} — install with: pip install {pkg}")
        ok = False

print("\nOptional (MP4 transcription):")
for mod, (pkg, desc) in optional.items():
    try:
        importlib.import_module(mod)
        print(f"  ✓ {pkg}  ({desc})")
    except ImportError:
        print(f"  ○ {pkg} not installed  ({desc})")

if sys.platform == "win32":
    print("\nWindows-only (optional):")
    for mod, (pkg, desc) in windows_only.items():
        try:
            importlib.import_module(mod)
            print(f"  ✓ {pkg}  ({desc})")
        except ImportError:
            print(f"  ○ {pkg} not installed  ({desc})")
else:
    print(f"\nWindows-only packages skipped on {sys.platform}.")
    print("  (Excel COM unavailable — pure-Python xlrd path will be used for .xls)")

# Env vars
print("\nEnvironment variables:")
gac = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
if gac:
    if os.path.exists(gac):
        print(f"  ✓ GOOGLE_APPLICATION_CREDENTIALS = {gac}")
    else:
        print(f"  ⚠️  GOOGLE_APPLICATION_CREDENTIALS set but file not found: {gac}")
        ok = False
else:
    print("  ○ GOOGLE_APPLICATION_CREDENTIALS not set (must pass at runtime)")

proj = os.environ.get("GCP_PROJECT_ID")
print(f"  {'✓' if proj else '○'} GCP_PROJECT_ID = {proj or '(not set)'}")

oai = os.environ.get("OPENAI_API_KEY")
if oai:
    print(f"  ✓ OPENAI_API_KEY set (length {len(oai)})")
else:
    print("  ○ OPENAI_API_KEY not set (only needed for MP4)")

# Skill folder integrity
print("\nSkill folder:")
script_dir = os.path.dirname(os.path.abspath(__file__))
expected = [
    "translate_dataroom.py",
    "transcribe_mp4.py",
    "timer.py",
    "fix_row_heights.py",
    "retranslate_xlsx_strings.py",
    "check_setup.py",
    "convert_xls.py",
    "decrypt_excel.py",
    "workflow_state.py",
]
for f in expected:
    p = os.path.join(script_dir, f)
    if os.path.exists(p):
        print(f"  ✓ scripts/{f}")
    else:
        print(f"  ✗ scripts/{f} MISSING")
        ok = False

# Test GCP auth (only if env vars set)
if gac and os.path.exists(gac) and proj:
    print("\nGCP Translation API test:")
    try:
        from google.cloud import translate_v3
        client = translate_v3.TranslationServiceClient()
        parent = f"projects/{proj}/locations/us-central1"
        resp = client.translate_text(
            request={
                "parent": parent,
                "contents": ["こんにちは"],
                "mime_type": "text/plain",
                "source_language_code": "ja",
                "target_language_code": "en",
            }
        )
        print(f"  ✓ Translation API works: 'こんにちは' → '{resp.translations[0].translated_text}'")
    except Exception as e:
        print(f"  ✗ Translation API failed: {e}")
        ok = False

    print("\nGCP Drive API test (optional, for cloud workflow):")
    try:
        from googleapiclient.discovery import build
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_file(
            gac, scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        svc = build("drive", "v3", credentials=creds, cache_discovery=False)
        # Just call about() to confirm access works
        svc.about().get(fields="user(emailAddress)").execute()
        print(f"  ✓ Drive API enabled and service account authenticated")
    except ImportError:
        print(f"  ○ google-api-python-client not installed (only needed for cloud Drive workflow)")
    except Exception as e:
        msg = str(e)
        if "Drive API has not been used" in msg or "is disabled" in msg:
            print(f"  ○ Drive API not enabled on project (only needed if using Drive folder URLs as input)")
            print(f"     Enable at: https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project={proj}")
        else:
            print(f"  ○ Drive API check failed: {msg[:200]}")

# Drive Streaming detection (informational)
print("\nGoogle Drive Streaming detection:")
candidate_paths = []
if sys.platform == "win32":
    home = os.path.expanduser("~")
    candidate_paths = [
        os.path.join(home, "Google Drive Streaming"),
        # Also check for letter drives (G:, H:)
        "G:\\My Drive",
        "H:\\My Drive",
    ]
elif sys.platform == "darwin":
    home = os.path.expanduser("~")
    candidate_paths = [
        os.path.join(home, "Library", "CloudStorage"),
        "/Volumes/GoogleDrive",
    ]
found_drive = False
for p in candidate_paths:
    if os.path.exists(p):
        print(f"  ✓ Found: {p}")
        found_drive = True
if not found_drive:
    print("  ○ No Drive Streaming mount detected (OK if not using Drive)")

print("\n" + "=" * 60)
print("READY" if ok else "NOT READY — see issues above")
print("=" * 60)
