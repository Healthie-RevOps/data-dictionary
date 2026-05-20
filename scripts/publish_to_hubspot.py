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
import json
import mimetypes
import os
import re
import sys
import urllib.parse
import uuid
from pathlib import Path

# The two pieces of identity we need are well-defined by the existing project (see
# UPLOAD_TO_HUBSPOT.md). They are constants, not configuration.
HUBSPOT_HOST = "api.hubapi.com"
TEMPLATE_PATH_IN_HUBSPOT = "custom/pages/PP SFDC Data Dictionary.html"
ENVIRONMENTS = ("draft", "published")

_REPO_ROOT = Path(__file__).resolve().parent.parent
# Body fragment produced by generate_sfdc_dictionary_html.py
DEFAULT_BODY_FRAGMENT = _REPO_ROOT / "peapod_sfdc_data_dictionary.html"
# Wrapped template (body + HubL chrome) that we PUT to HubSpot
DEFAULT_LOCAL_TEMPLATE = _REPO_ROOT / "deploy_template_peapod_sfdc_data_dictionary.html"

# HubSpot CMS source-code files cap at 2.0 MiB. The unminified template runs
# ~2.36 MiB after the "All Other Fields" section was added, so we minify before
# pushing. The local rendered file stays human-readable for browser preview.
HUBSPOT_SOURCE_CODE_LIMIT = 2 * 1024 * 1024

# HubL/HTML chrome that wraps the body fragment for HubSpot Design Manager.
# The `templateType: "none"` annotation tells HubSpot this is a custom HTML
# file used by a landing page, not a drag-and-drop or email template.
_HUBL_HEADER = """<!--
  templateType: "none"
  isAvailableForNewContent: false
-->
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<meta name="robots" content="noindex,nofollow">
<title>Healthie Salesforce Data Dictionary</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>body{margin:0;padding:0;background:#f8fafc;font-family:"Inter",system-ui,sans-serif;}</style>
</head>
<body>
"""

_HUBL_FOOTER = """
</body>
</html>"""


def wrap_body_fragment(body_html: str) -> str:
    """Wrap the body fragment with HubL/HTML chrome for upload to HubSpot."""
    return _HUBL_HEADER + body_html + _HUBL_FOOTER


def _minify_style_block(body: str) -> str:
    body = re.sub(r"/\*.*?\*/", "", body, flags=re.DOTALL)
    body = re.sub(r"\s+", " ", body)
    body = re.sub(r"\s*([{};:,>])\s*", r"\1", body)
    return body.strip()


def _compact_json_record(rec):
    if isinstance(rec, dict):
        return {k: v for k, v in rec.items() if v not in ("", False, None, [], {})}
    return rec


def _minify_json_island(body: str) -> str:
    try:
        data = json.loads(body)
    except Exception:
        return body
    if isinstance(data, dict):
        data = {k: [_compact_json_record(r) for r in v] if isinstance(v, list) else v
                for k, v in data.items()}
    return json.dumps(data, separators=(",", ":"), ensure_ascii=False)


def minify_template(html: str) -> str:
    """Shrink the rendered HTML to fit under HubSpot's 2.0 MiB source-code limit.

    The unmodified body has lots of indentation and per-row class repetition.
    We minify in three passes — outer HTML whitespace, CSS, and the export
    JSON island — and leave the JS block alone (it's a raw-string regex
    minefield).
    """
    placeholders: list[str] = []

    def stash(m):
        placeholders.append(m.group(0))
        return f"\x00P{len(placeholders) - 1}\x00"

    html = re.sub(r"<(script|style)\b[^>]*>.*?</\1>", stash, html, flags=re.DOTALL | re.IGNORECASE)
    # Strip leading whitespace and blank lines, collapse inter-tag whitespace
    html = re.sub(r"^[ \t]+", "", html, flags=re.MULTILINE)
    html = re.sub(r">\s+<", "><", html)
    html = re.sub(r"\n\s*\n+", "\n", html)

    def minify_block(blk: str) -> str:
        m = re.match(r"(<(?P<tag>script|style)\b[^>]*>)(.*)(</(?P=tag)>)",
                     blk, flags=re.DOTALL | re.IGNORECASE)
        if not m:
            return blk
        opener, body, closer = m.group(1), m.group(3), m.group(4)
        opener_l = opener.lower()
        if 'application/json' in opener_l:
            return opener + _minify_json_island(body) + closer
        if opener_l.startswith('<style'):
            return opener + _minify_style_block(body) + closer
        return opener + body + closer  # leave JS alone

    for i, blk in enumerate(placeholders):
        html = html.replace(f"\x00P{i}\x00", minify_block(blk))
    return html


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
        "--body",
        type=Path,
        default=DEFAULT_BODY_FRAGMENT,
        help=f"Path to the body fragment from the generator (default: {DEFAULT_BODY_FRAGMENT.name})",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=None,
        help="Skip body wrapping and use this pre-wrapped template file directly.",
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

    if args.template is not None:
        template_path: Path = args.template
        if not template_path.is_file():
            print(f"ERROR: template file not found: {template_path}", file=sys.stderr)
            return 2
        template_bytes = template_path.read_bytes()
        if not template_bytes.strip():
            print(f"ERROR: template file is empty: {template_path}", file=sys.stderr)
            return 2
        print(f"[publish] loaded pre-wrapped template {template_path.name} ({len(template_bytes):,} bytes)")
    else:
        body_path: Path = args.body
        if not body_path.is_file():
            print(f"ERROR: body fragment not found: {body_path}", file=sys.stderr)
            print("       Run `python3 generate_sfdc_dictionary_html.py --render-only` first.", file=sys.stderr)
            return 2
        body_html = body_path.read_text(encoding="utf-8")
        if not body_html.strip():
            print(f"ERROR: body fragment is empty: {body_path}", file=sys.stderr)
            return 2
        print(f"[publish] loaded body fragment {body_path.name} ({len(body_html):,} chars)")
        wrapped = wrap_body_fragment(body_html)
        template_bytes = wrapped.encode("utf-8")
        DEFAULT_LOCAL_TEMPLATE.write_text(wrapped, encoding="utf-8")
        print(f"[publish] wrote wrapped template -> {DEFAULT_LOCAL_TEMPLATE.name} ({len(template_bytes):,} bytes)")

    minified = minify_template(template_bytes.decode("utf-8"))
    minified_bytes = minified.encode("utf-8")
    saved = len(template_bytes) - len(minified_bytes)
    print(f"[publish] minified -> {len(minified_bytes):,} bytes ({saved:+,} = "
          f"{(saved/len(template_bytes))*100:.1f}% smaller)")
    if len(minified_bytes) > HUBSPOT_SOURCE_CODE_LIMIT:
        print(f"ERROR: minified template is {len(minified_bytes):,} bytes, "
              f"over HubSpot's {HUBSPOT_SOURCE_CODE_LIMIT:,}-byte source-code limit.",
              file=sys.stderr)
        print("       HubSpot will reject this push. Reduce content before retrying.",
              file=sys.stderr)
        return 2
    template_bytes = minified_bytes

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
