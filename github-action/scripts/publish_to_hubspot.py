#!/usr/bin/env python3
"""
Publish the generated PP SFDC Data Dictionary template to HubSpot's CMS source-code API.

This is the second half of the weekly sync. generate_sfdc_dictionary_html.py produces
deploy_template_peapod_sfdc_data_dictionary.html locally; this script PUTs that file
to HubSpot's /cms/v3/source-code endpoint for both draft and published environments,
keeping them in sync.

Auth: reads a HubSpot personal-access-key token from the HUBSPOT_TOKEN env var.
The token needs the cms.source_code.read and cms.source_code.write scopes — the
existing key in hubspot.config.yml already has these.

Usage:
    HUBSPOT_TOKEN=pat-xxx python3 scripts/publish_to_hubspot.py
    HUBSPOT_TOKEN=pat-xxx python3 scripts/publish_to_hubspot.py --dry-run
    HUBSPOT_TOKEN=pat-xxx python3 scripts/publish_to_hubspot.py --template path/to/template.html

Exits non-zero on any HTTP error so a CI workflow notices.

stdlib-only, matches the style of generate_sfdc_dictionary_html.py (no requests, no PyYAML).
"""
from __future__ import annotations

import argparse
import http.client
import io
import mimetypes
import os
import sys
import urllib.parse
import uuid
from pathlib import Path

# The two pieces of identity we need are well-defined by the existing project (see
# UPLOAD_TO_HUBSPOT.md). They are constants, not configuration.
HUBSPOT_HOST = "api.hubapi.com"
TEMPLATE_PATH_IN_HUBSPOT = "custom/pages/PP SFDC Data Dictionary.html"
ENVIRONMENTS = ("draft", "published")

# Default local template file produced by generate_sfdc_dictionary_html.py
DEFAULT_LOCAL_TEMPLATE = Path(__file__).resolve().parent.parent / "deploy_template_peapod_sfdc_data_dictionary.html"


def _build_multipart(field_name: str, filename: str, file_bytes: bytes) -> tuple[bytes, str]:
    """Build a minimal multipart/form-data body. Returns (body, content_type)."""
    boundary = "----HSFormBoundary" + uuid.uuid4().hex
    content_type = mimetypes.guess_type(filename)[0] or "text/html"
    buf = io.BytesIO()
    buf.write(f"--{boundary}\r\n".encode())
    buf.write(
        f'Content-Disposition: form-data; name="{field_name}"; filename="{filename}"\r\n'.encode()
    )
    buf.write(f"Content-Type: {content_type}\r\n\r\n".encode())
    buf.write(file_bytes)
    buf.write(f"\r\n--{boundary}--\r\n".encode())
    return buf.getvalue(), f"multipart/form-data; boundary={boundary}"


def _put_template(token: str, env: str, template_bytes: bytes, *, dry_run: bool) -> int:
    # Per UPLOAD_TO_HUBSPOT.md the endpoint is:
    #   PUT /cms/v3/source-code/{env}/content/{path}
    # The path segment must be URL-encoded (spaces -> %20, slashes preserved as separators
    # by HubSpot — encoding the whole thing including slashes is the documented contract).
    encoded_path = urllib.parse.quote(TEMPLATE_PATH_IN_HUBSPOT, safe="")
    url_path = f"/cms/v3/source-code/{env}/content/{encoded_path}"

    body, content_type = _build_multipart(
        field_name="file",
        filename=Path(TEMPLATE_PATH_IN_HUBSPOT).name,
        file_bytes=template_bytes,
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type,
        "Content-Length": str(len(body)),
        "Accept": "application/json",
    }

    print(f"[publish] PUT https://{HUBSPOT_HOST}{url_path}  ({len(template_bytes):,} bytes, env={env})")
    if dry_run:
        print(f"[publish] DRY RUN — would send {len(body):,}-byte multipart body, content-type={content_type}")
        return 0

    conn = http.client.HTTPSConnection(HUBSPOT_HOST, timeout=60)
    try:
        conn.request("PUT", url_path, body=body, headers=headers)
        resp = conn.getresponse()
        resp_body = resp.read().decode("utf-8", errors="replace")
        print(f"[publish] env={env} -> HTTP {resp.status} {resp.reason}")
        if resp.status >= 400:
            print(f"[publish] response body: {resp_body[:2000]}")
            return resp.status
        # On success HubSpot returns a JSON record of the source-code file
        print(f"[publish] env={env} ok ({len(resp_body):,} bytes response)")
        return 0
    finally:
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_LOCAL_TEMPLATE,
        help=f"Path to the local template HTML file (default: {DEFAULT_LOCAL_TEMPLATE.name})",
    )
    parser.add_argument(
        "--env",
        choices=("draft", "published", "both"),
        default="both",
        help="Which HubSpot environment(s) to push to. Default: both (matches existing manual workflow).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Build the multipart body and print what would be sent, but do not call HubSpot.",
    )
    args = parser.parse_args()

    token = os.environ.get("HUBSPOT_TOKEN", "").strip()
    if not token and not args.dry_run:
        print("ERROR: HUBSPOT_TOKEN env var is empty. Refusing to send a request without a token.", file=sys.stderr)
        print("Set HUBSPOT_TOKEN to a HubSpot personal access key with cms.source_code.write scope.", file=sys.stderr)
        return 2

    template_path: Path = args.template
    if not template_path.is_file():
        print(f"ERROR: template file not found: {template_path}", file=sys.stderr)
        return 2

    template_bytes = template_path.read_bytes()
    if not template_bytes.strip():
        print(f"ERROR: template file is empty: {template_path}", file=sys.stderr)
        return 2
    print(f"[publish] loaded {template_path.name} ({len(template_bytes):,} bytes)")

    envs = (args.env,) if args.env != "both" else ENVIRONMENTS

    rc = 0
    for env in envs:
        status = _put_template(token, env, template_bytes, dry_run=args.dry_run)
        if status != 0:
            rc = status
            # Keep going to the next env even if one fails so we report both — but
            # the process exit code will reflect the first failure.

    if rc == 0:
        print("[publish] all targets succeeded")
        if not args.dry_run:
            print(
                "[publish] NOTE: HubSpot's prerender cache may serve the previous version "
                "of the live page for up to 10 hours. To flush, click 'Update' on the "
                "landing page in HubSpot's UI. (See UPLOAD_TO_HUBSPOT.md.)"
            )
    return rc


if __name__ == "__main__":
    sys.exit(main())
