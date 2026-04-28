#!/usr/bin/env python3
"""
drive_sync.py - Download from / upload to Google Drive folders via Drive API v3.

Cross-platform alternative to reading from Drive Streaming filesystem. Use this
when running in cloud Claude Code (claude.ai/code) or when Drive Streaming isn't
available.

Authentication strategies (auto-tried in order):
  1. Service account via GOOGLE_APPLICATION_CREDENTIALS env var (preferred — no UI)
  2. OAuth installed-app flow with cached token at ~/.claude/skills/translate-dataroom/.drive_token.json

For service account auth to work, the service account email must be added as a
member of the Shared Drive containing the folder. Find the email in the SA JSON
file's `client_email` field.

Usage:
  # Download a Drive folder to local
  python drive_sync.py down <folder_id_or_url> <local_dest>

  # Upload local folder to a Drive folder (creates if needed)
  python drive_sync.py up <local_src> <parent_folder_id> --name "English"

  # List a folder's contents (recursive)
  python drive_sync.py ls <folder_id_or_url>

  # Resolve a Drive URL to a folder ID
  python drive_sync.py resolve <url>
"""

import _cloud_creds  # noqa: F401  -- bootstraps GCP creds for cloud sandbox

import os
import sys
import re
import io
import argparse
import mimetypes
from pathlib import Path

SCOPES_RW = ["https://www.googleapis.com/auth/drive"]
SCOPES_RO = ["https://www.googleapis.com/auth/drive.readonly"]

# Google Workspace MIME types and what to export them as
GOOGLE_EXPORT = {
    "application/vnd.google-apps.document":
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
    "application/vnd.google-apps.spreadsheet":
        ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "application/vnd.google-apps.presentation":
        ("application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"),
}


def extract_folder_id(url_or_id):
    """Extract a Google Drive folder ID from a URL or pass through if already an ID."""
    if not url_or_id:
        return None
    # If it looks like a bare ID (no slashes, alphanumeric+_-), return as-is
    if re.match(r'^[a-zA-Z0-9_-]+$', url_or_id) and len(url_or_id) > 20:
        return url_or_id
    # URL patterns:
    # https://drive.google.com/drive/folders/<ID>
    # https://drive.google.com/drive/folders/<ID>?usp=sharing
    # https://drive.google.com/drive/u/0/folders/<ID>
    m = re.search(r'/folders/([a-zA-Z0-9_-]+)', url_or_id)
    if m:
        return m.group(1)
    # File patterns
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url_or_id)
    if m:
        return m.group(1)
    raise ValueError(f"Could not extract folder ID from: {url_or_id}")


def get_service(scopes=SCOPES_RW):
    """Return an authenticated Drive service. Tries SA first, then OAuth."""
    from googleapiclient.discovery import build

    # Method 1: Service account via GOOGLE_APPLICATION_CREDENTIALS
    sa_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if sa_path and os.path.exists(sa_path):
        try:
            from google.oauth2 import service_account
            creds = service_account.Credentials.from_service_account_file(sa_path, scopes=scopes)
            print(f"  Auth: service account ({os.path.basename(sa_path)})")
            return build("drive", "v3", credentials=creds, cache_discovery=False)
        except Exception as e:
            print(f"  Service account auth failed: {e}")

    # Method 2: OAuth installed app
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from google_auth_oauthlib.flow import InstalledAppFlow

        token_path = os.path.expanduser(
            "~/.claude/skills/translate-dataroom/.drive_token.json"
        )
        client_secrets = os.environ.get(
            "DRIVE_OAUTH_CLIENT_SECRETS",
            os.path.expanduser("~/.claude/skills/translate-dataroom/.oauth_client.json"),
        )

        creds = None
        if os.path.exists(token_path):
            creds = Credentials.from_authorized_user_file(token_path, scopes)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(client_secrets):
                    raise RuntimeError(
                        f"No service account auth and no OAuth client at {client_secrets}. "
                        "Set GOOGLE_APPLICATION_CREDENTIALS to a service account JSON, or "
                        "set DRIVE_OAUTH_CLIENT_SECRETS to an OAuth client_secrets.json."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(client_secrets, scopes)
                creds = flow.run_local_server(port=0)
            os.makedirs(os.path.dirname(token_path), exist_ok=True)
            with open(token_path, "w") as f:
                f.write(creds.to_json())
        print(f"  Auth: user OAuth (token at {token_path})")
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"All auth methods failed. Last error: {e}")


def list_folder(service, folder_id, recursive=False, _path="", out=None):
    """Yield (rel_path, file_metadata) for every file in folder. If recursive, descends subfolders."""
    if out is None:
        out = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType, size, modifiedTime, parents)",
            pageSize=1000,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            corpora="allDrives",
            pageToken=page_token,
        ).execute()
        for f in resp.get("files", []):
            rel = f["name"] if not _path else f"{_path}/{f['name']}"
            if f["mimeType"] == "application/vnd.google-apps.folder":
                if recursive:
                    list_folder(service, f["id"], True, rel, out)
                # Skip folders themselves in output (we want only files)
            else:
                f["_rel_path"] = rel
                out.append(f)
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return out


def download_file(service, file_meta, dst_path):
    """Download a single Drive file to dst_path. Handles Google Workspace exports."""
    from googleapiclient.http import MediaIoBaseDownload

    os.makedirs(os.path.dirname(dst_path) or ".", exist_ok=True)
    mime = file_meta["mimeType"]

    if mime in GOOGLE_EXPORT:
        export_mime, ext = GOOGLE_EXPORT[mime]
        if not dst_path.endswith(ext):
            dst_path = dst_path + ext
        request = service.files().export_media(fileId=file_meta["id"], mimeType=export_mime)
    else:
        request = service.files().get_media(fileId=file_meta["id"], supportsAllDrives=True)

    fh = io.FileIO(dst_path, "wb")
    downloader = MediaIoBaseDownload(fh, request, chunksize=1024 * 1024)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()
    return dst_path


def cmd_down(folder_id, local_dest, recursive=True):
    """Download all files from a Drive folder to local_dest."""
    service = get_service(SCOPES_RO)
    folder_id = extract_folder_id(folder_id)

    # Get folder metadata for the name
    try:
        folder_meta = service.files().get(
            fileId=folder_id, fields="id,name,mimeType,parents",
            supportsAllDrives=True,
        ).execute()
        print(f"  Source folder: {folder_meta['name']} (id={folder_id})")
    except Exception as e:
        print(f"  ERROR getting folder metadata: {e}")
        raise

    print(f"  Listing files (recursive={recursive})...")
    files = list_folder(service, folder_id, recursive=recursive)
    print(f"  Found {len(files)} files")
    total_size = sum(int(f.get("size", 0)) for f in files)
    print(f"  Total size: {total_size / 1024 / 1024:.1f} MB")

    os.makedirs(local_dest, exist_ok=True)
    for i, f in enumerate(files, 1):
        rel = f["_rel_path"]
        size = int(f.get("size", 0))
        dst = os.path.join(local_dest, rel)
        print(f"[{i}/{len(files)}] {rel} ({size} bytes)", flush=True)
        try:
            actual = download_file(service, f, dst)
            if actual != dst:
                print(f"    -> {os.path.basename(actual)}")
        except Exception as e:
            print(f"    FAILED: {e}")

    print(f"\nDone. Files in: {local_dest}")


def cmd_up(local_src, parent_folder_id, name=None, drive_id=None):
    """Upload a local folder to Drive under parent_folder_id. Optionally name the new subfolder."""
    from googleapiclient.http import MediaFileUpload

    service = get_service(SCOPES_RW)
    parent_folder_id = extract_folder_id(parent_folder_id)
    new_folder_name = name or os.path.basename(os.path.normpath(local_src))

    # Create top-level folder
    folder_meta = service.files().create(
        body={
            "name": new_folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_folder_id],
        },
        fields="id,name",
        supportsAllDrives=True,
    ).execute()
    print(f"  Created Drive folder: {folder_meta['name']} (id={folder_meta['id']})")

    # Walk local files, mirroring directory structure
    folder_cache = {".": folder_meta["id"]}

    def get_or_create_subfolder(rel_dir):
        """Get cached or create a subfolder, return its Drive ID."""
        if rel_dir in folder_cache:
            return folder_cache[rel_dir]
        parent_rel = os.path.dirname(rel_dir) or "."
        parent_id = get_or_create_subfolder(parent_rel)
        sub = service.files().create(
            body={
                "name": os.path.basename(rel_dir),
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_id],
            },
            fields="id",
            supportsAllDrives=True,
        ).execute()
        folder_cache[rel_dir] = sub["id"]
        return sub["id"]

    files = []
    for root, dirs, fnames in os.walk(local_src):
        for f in fnames:
            full = os.path.join(root, f)
            rel = os.path.relpath(full, local_src)
            files.append((full, rel.replace(os.sep, "/")))

    print(f"  Uploading {len(files)} files...")
    for i, (full, rel) in enumerate(files, 1):
        rel_dir = os.path.dirname(rel) or "."
        parent_id = get_or_create_subfolder(rel_dir)
        mime, _ = mimetypes.guess_type(full)
        media = MediaFileUpload(full, mimetype=mime or "application/octet-stream")
        try:
            up = service.files().create(
                body={"name": os.path.basename(rel), "parents": [parent_id]},
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            ).execute()
            print(f"  [{i}/{len(files)}] {rel} -> {up['id']}", flush=True)
        except Exception as e:
            print(f"  [{i}/{len(files)}] {rel} FAILED: {e}")

    print(f"\nDone. Drive folder: https://drive.google.com/drive/folders/{folder_meta['id']}")
    return folder_meta['id']


def cmd_ls(folder_id, recursive=False):
    service = get_service(SCOPES_RO)
    folder_id = extract_folder_id(folder_id)
    files = list_folder(service, folder_id, recursive=recursive)
    print(f"\n{len(files)} files:")
    total = 0
    for f in files:
        size = int(f.get("size", 0))
        total += size
        print(f"  {size:>10}  {f['_rel_path']}  [{f['mimeType']}]")
    print(f"\nTotal: {total / 1024 / 1024:.1f} MB")


def cmd_resolve(url):
    fid = extract_folder_id(url)
    print(fid)


def main():
    p = argparse.ArgumentParser(description="Google Drive sync via API")
    sub = p.add_subparsers(dest="cmd", required=True)

    p_down = sub.add_parser("down", help="Download a Drive folder to local")
    p_down.add_argument("folder", help="Drive folder URL or ID")
    p_down.add_argument("dest", help="Local destination directory")
    p_down.add_argument("--non-recursive", action="store_true",
                        help="Only download files in this folder, not subfolders")

    p_up = sub.add_parser("up", help="Upload a local folder to Drive")
    p_up.add_argument("src", help="Local source directory")
    p_up.add_argument("parent", help="Drive parent folder ID/URL")
    p_up.add_argument("--name", help="Name for the new Drive folder (default: local folder name)")

    p_ls = sub.add_parser("ls", help="List a Drive folder")
    p_ls.add_argument("folder", help="Drive folder URL or ID")
    p_ls.add_argument("--recursive", action="store_true", help="Recurse into subfolders")

    p_res = sub.add_parser("resolve", help="Extract folder ID from URL")
    p_res.add_argument("url")

    args = p.parse_args()

    if args.cmd == "down":
        cmd_down(args.folder, args.dest, recursive=not args.non_recursive)
    elif args.cmd == "up":
        cmd_up(args.src, args.parent, name=args.name)
    elif args.cmd == "ls":
        cmd_ls(args.folder, recursive=args.recursive)
    elif args.cmd == "resolve":
        cmd_resolve(args.url)


if __name__ == "__main__":
    main()
