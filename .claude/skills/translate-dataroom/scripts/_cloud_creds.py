"""
_cloud_creds.py — bootstrap GCP credentials and TLS settings for cloud sandboxes.

Two responsibilities, both no-ops on local Windows/macOS where things just work:

1. **Credential materialization.** Cloud sandboxes (claude.ai/code) can only set
   string env vars, but the Google client libraries expect
   `GOOGLE_APPLICATION_CREDENTIALS` to be a *file path*. If
   `GOOGLE_APPLICATION_CREDENTIALS_JSON` is set and
   `GOOGLE_APPLICATION_CREDENTIALS` isn't, write the JSON to a secure temp file
   and point the standard env var at it.

2. **TLS trust on intercepting Linux sandboxes.** Claude Code on the Web routes
   outbound HTTPS through a TLS-inspecting proxy whose CA
   (`O=Anthropic, CN=sandbox-egress-production TLS Inspection CA`) is installed
   into the system bundle (`/etc/ssl/certs/ca-certificates.crt`) but NOT into
   certifi's bundled `cacert.pem`. So:
   - **gRPC** ignores `SSL_CERT_FILE` / `REQUESTS_CA_BUNDLE` and needs its own
     env var: `GRPC_DEFAULT_SSL_ROOTS_FILE_PATH`. Required for the GCP
     Translation API client.
   - **httplib2** (the default transport for `googleapiclient`, used for the
     Drive API) reads its `CA_CERTS` class attribute, which defaults to
     `certifi.where()`. We patch it on import.

   Both patches only fire on Linux when the system bundle exists, so local dev
   on Windows / macOS is untouched.

Import this at the top of any script that uses GCP libraries, BEFORE any
`from google...` or `import googleapiclient` imports:

    import _cloud_creds  # noqa: F401  -- bootstraps GCP creds + TLS for cloud
"""

import json
import os
import stat
import sys
import tempfile

SYSTEM_CA_BUNDLE = "/etc/ssl/certs/ca-certificates.crt"


def _bootstrap_gcp_creds() -> None:
    """Materialize GOOGLE_APPLICATION_CREDENTIALS_JSON to a file path."""
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
        try:
            os.unlink(path)
        except OSError:
            pass
        raise

    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = path


def _bootstrap_tls() -> None:
    """Make gRPC and httplib2 trust the system CA bundle on Linux.

    Linux-only because that's where TLS-intercepting cloud sandboxes run; on
    Windows/macOS local dev, the system trust stores already cover everything.
    """
    if sys.platform != "linux":
        return
    if not os.path.exists(SYSTEM_CA_BUNDLE):
        return

    # gRPC's C-Core ignores SSL_CERT_FILE / REQUESTS_CA_BUNDLE; it reads only
    # this var. setdefault so an explicit user override still wins.
    os.environ.setdefault("GRPC_DEFAULT_SSL_ROOTS_FILE_PATH", SYSTEM_CA_BUNDLE)

    # googleapiclient defaults to httplib2, which reads CA_CERTS at request
    # time. Patch the class attribute if the lib is importable; no-op
    # otherwise. Done at module load so any later `from googleapiclient...`
    # gets the patched version.
    try:
        import httplib2  # type: ignore[import-untyped]

        httplib2.CA_CERTS = SYSTEM_CA_BUNDLE
    except ImportError:
        pass


_bootstrap_gcp_creds()
_bootstrap_tls()
