"""
_cloud_creds.py — bootstrap GCP credentials from a JSON-string env var.

Cloud sandboxes (claude.ai/code, claude.ai chat) can only set string env vars,
not upload files. The Google client libraries, however, expect
GOOGLE_APPLICATION_CREDENTIALS to be a *file path*. This module bridges that gap:

    if GOOGLE_APPLICATION_CREDENTIALS_JSON is set and GOOGLE_APPLICATION_CREDENTIALS
    is not, write the JSON to a secure temp file and point
    GOOGLE_APPLICATION_CREDENTIALS at that file.

Import this module at the top of any script that needs GCP auth, BEFORE any
`from google...` imports. It runs the bootstrap on import and is a no-op on
machines where GOOGLE_APPLICATION_CREDENTIALS already points at a real file
(i.e., your local Windows / macOS setup is unaffected).

Usage:
    import _cloud_creds  # noqa: F401  -- bootstraps GCP creds for cloud sandbox
"""

import json
import os
import stat
import tempfile


def _bootstrap() -> None:
    existing = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if existing and os.path.exists(existing):
        # Local setup — file path is real, do nothing.
        return

    raw = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if not raw:
        # Nothing to do. Caller will get the usual "auth not configured" error.
        return

    # Validate it parses as JSON so we fail fast with a clear message rather
    # than producing a malformed credentials file.
    try:
        json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            "GOOGLE_APPLICATION_CREDENTIALS_JSON is set but is not valid JSON: "
            f"{e}. Paste the full contents of the service account JSON file as "
            "the value of this env var."
        ) from e

    # Write to a tempfile that survives for the lifetime of the process.
    fd, path = tempfile.mkstemp(prefix="gcp-sa-", suffix=".json")
    try:
        with os.fdopen(fd, "w") as f:
            f.write(raw)
        # Tighten perms (no-op on Windows, defensive on Linux/macOS).
        try:
            os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)
        except OSError:
            pass
    except Exception:
        # Clean up if we failed mid-write.
        try:
            os.unlink(path)
        except OSError:
            pass
        raise

    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = path


_bootstrap()
