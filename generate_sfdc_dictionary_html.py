#!/usr/bin/env python3
"""Generate the Healthie SFDC Data Dictionary v6 (Web Edition) and publish it
to the HubSpot CMS page at https://go.gethealthie.com/peapod-sfdc-data-dictionary.

See README.md for the overview and UPLOAD_TO_HUBSPOT.md for the publishing
walkthrough; this docstring is the source-of-truth for what the script does.

Pipeline:
  1. Parse Healthie_SFDC_Data_Dictionary_v5.docx (canonical field catalog).
  2. Augment with live Salesforce metadata via the `sf` CLI:
       - FieldDefinition (Tooling) for API names, DurableIds, types, descriptions
       - sf sobject describe for picklist active values
       - SELECT COUNT(<field>) FROM <object> aggregates for Fill %
       - CustomField (Tooling) for CreatedDate → NEW field chips (rolling 6mo)
       - Flow + FlowDefinition (Tooling) for the SFDC Flows section
     All results cached under .dictionary_cache/sfdc_*.
  3. Build a render model:
       - Per-row: ★ Key field (lookup/external-ID), NEW chip (rolling 180-day
         window), similar-field clusters with tooltip siblings.
       - Per-object: accent color for tab dot + section header band + sub-section
         left border.
       - Flows: with inferred triggering object from the developer-name prefix.
  4. Render a single HTML body fragment with inline CSS+JS, mirroring the visual
     model of /peapod-hubspot-data-dictionary while keeping the SFDC object
     taxonomy. Adds: global cross-object search, sticky Field Label column,
     SFDC Object Manager hyperlinks per field, and the Flows section with its
     own local search.
  5. (Optional) Push to HubSpot via /cms/v3/source-code (requires
     `cms.source_code.write` scope on the personal access key).

Stdlib-only (plus the `sf` CLI as a subprocess). No `requests`, no `PyYAML`.

CLI flags:
  --refresh             invalidate all caches and re-fetch from SFDC
  --refresh-fielddefs   re-fetch FieldDefinition only
  --refresh-fill        re-fetch Fill % counts only
  --refresh-picklists   re-fetch picklist active values only
  --skip-sfdc           skip all live SFDC calls (use cached data only)
  --render-only         skip SFDC + just regenerate the HTML from cache
"""
from __future__ import annotations

import argparse
import html
import json
import os
import re
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
CACHE_DIR = PROJECT_DIR / ".dictionary_cache"
V5_DOCX = PROJECT_DIR / "Healthie_SFDC_Data_Dictionary_v5.docx"
OUTPUT_HTML = PROJECT_DIR / "peapod_sfdc_data_dictionary.html"
DEFAULT_ORG = "bill.coffin@gethealthie.com"
SFDC_SETUP_BASE = "https://healthie.my.salesforce-setup.com/lightning/setup/ObjectManager"

W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Mapping from V5 H1 object label -> Salesforce API name.
OBJECT_API_NAME = {
    "Account": "Account",
    "Contact": "Contact",
    "Opportunity": "Opportunity",
    "Healthie Organization": "Healthie_Organization__c",
    "Stripe Customer": "Stripe_Customer__c",
    "Stripe Subscription": "Stripe_Subscription__c",
    "TaskRay Project": "taskray__Project__c",
}

# Object render order on the page (left to right tab order).
OBJECT_ORDER = [
    "Account", "Contact", "Opportunity", "Healthie Organization",
    "Stripe Customer", "Stripe Subscription", "TaskRay Project",
]

# Per-object accent color. Used for the active tab indicator, subsection
# title left border, and panel summary border so users can recognize which
# object's table they're scanning at a glance.
OBJECT_COLOR = {
    "Account":               "#4f46e5",  # indigo
    "Contact":               "#0891b2",  # cyan
    "Opportunity":           "#d97706",  # amber
    "Healthie Organization": "#10b981",  # emerald
    "Stripe Customer":       "#8b5cf6",  # violet
    "Stripe Subscription":   "#db2777",  # pink
    "TaskRay Project":       "#ea580c",  # orange
}

OBJECT_LABELS_LOWER = {k.lower(): v for k, v in OBJECT_API_NAME.items()}


# ---------- docx parsing ----------

def _text_of(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(W + "t"))


def parse_v5(docx_path: Path) -> dict:
    with zipfile.ZipFile(docx_path) as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")
    root = ET.fromstring(doc_xml)
    body = root.find(W + "body")
    if body is None:
        raise RuntimeError("No body element in V5 docx")

    out: dict = {
        "front_matter": {
            "how_to_use_paragraphs": [],
            "column_definitions": [],
            "governance": [],
            "data_sources": [],
            "glossary": [],
        },
        "objects": [],
        "appendix": {
            "picklist_values": [],
            "v4_additions": None,
            "changelog": None,
        },
        "meta": {
            "v5_mtime": datetime.fromtimestamp(docx_path.stat().st_mtime, tz=timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        },
    }

    current_h1 = ""
    current_h2 = ""
    appendix_started = False

    # We capture all paragraphs that belong under "How to Use This Dictionary" so
    # we can render the methodology text on the page. Tables encountered while
    # under specific H1/H2 contexts are dispatched accordingly.
    current_obj: dict | None = None
    current_section: dict | None = None

    for elem in body:
        tag = elem.tag.split("}")[-1]
        if tag == "p":
            style = elem.find("w:pPr/w:pStyle", W_NS)
            sval = style.get(W + "val") if style is not None else None
            txt = _text_of(elem).strip()

            if sval == "Heading1" and txt:
                current_h1 = txt
                current_h2 = ""
                if txt in OBJECT_API_NAME:
                    current_obj = {"label": txt, "sections": []}
                    out["objects"].append(current_obj)
                    current_section = None
                else:
                    current_obj = None
                    current_section = None
                if txt.startswith("Appendix"):
                    appendix_started = True
                continue

            if sval == "Heading2" and txt:
                current_h2 = txt
                if current_obj is not None:
                    current_section = {"title": txt, "fields": []}
                    current_obj["sections"].append(current_section)
                continue

            if current_h1 == "How to Use This Dictionary" and txt and sval not in ("Heading1", "Heading2"):
                out["front_matter"]["how_to_use_paragraphs"].append(txt)

        elif tag == "tbl":
            rows = elem.findall("w:tr", W_NS)
            if not rows:
                continue
            header = [_text_of(c).strip() for c in rows[0].findall("w:tc", W_NS)]
            data_rows = [
                [_text_of(c).strip() for c in r.findall("w:tc", W_NS)]
                for r in rows[1:]
            ]

            # Front-matter tables
            if current_h1 == "How to Use This Dictionary" and header == ["Column", "What it captures"]:
                out["front_matter"]["column_definitions"] = data_rows
                continue
            if current_h1 == "Governance" and header == ["Area", "Owner / Process"]:
                out["front_matter"]["governance"] = data_rows
                continue
            if current_h1 == "Data Sources & Integrations" and len(header) >= 3 and header[0] == "Source System":
                out["front_matter"]["data_sources"] = data_rows
                continue
            if current_h1 == "Glossary" and header == ["Term", "Definition"]:
                out["front_matter"]["glossary"] = data_rows
                continue

            # Object field tables (8 cols, first header "Field Label")
            if (
                current_obj is not None
                and current_section is not None
                and header[:1] == ["Field Label"]
                and len(header) == 8
            ):
                for row in data_rows:
                    if len(row) < 8:
                        row = row + [""] * (8 - len(row))
                    current_section["fields"].append({
                        "label": row[0],
                        "type_v5": row[1],
                        "pop_pct_v5": row[2],
                        "source": row[3],
                        "owner": row[4],
                        "sensitivity": row[5],
                        "description_v5": row[6],
                        "dependencies": row[7],
                    })
                continue

            # Appendix tables
            if appendix_started:
                if current_h1 == "Appendix: Picklist Values":
                    out["appendix"]["picklist_values"].append({
                        "title": current_h2 or "(unlabeled)",
                        "headers": header,
                        "rows": data_rows,
                    })
                    continue
                if current_h1 == "Appendix: V4 Field Additions":
                    out["appendix"]["v4_additions"] = {
                        "headers": header,
                        "rows": data_rows,
                    }
                    continue
                if current_h1 == "Changelog":
                    out["appendix"]["changelog"] = {
                        "headers": header,
                        "rows": data_rows,
                    }
                    continue

    return out


# ---------- cache ----------

def cache_path(name: str) -> Path:
    CACHE_DIR.mkdir(exist_ok=True)
    return CACHE_DIR / name


def cache_load(name: str):
    p = cache_path(name)
    if p.exists():
        try:
            return json.loads(p.read_text())
        except Exception:
            return None
    return None


def cache_save(name: str, value) -> None:
    cache_path(name).write_text(json.dumps(value, indent=2))


# ---------- Salesforce client (via `sf` CLI) ----------

class SfClient:
    def __init__(self, org: str = DEFAULT_ORG):
        self.org = org

    def soql(self, query: str, *, tooling: bool = False) -> dict:
        cmd = [
            "sf", "data", "query",
            "--query", query,
            "--target-org", self.org,
            "--json",
        ]
        if tooling:
            cmd.append("--use-tooling-api")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        # The CLI emits an update-available warning on stderr; the real result
        # (success or failure) is always JSON on stdout. Parse stdout first.
        try:
            payload = json.loads(result.stdout)
        except json.JSONDecodeError as e:
            raise RuntimeError(
                f"sf data query produced unparseable stdout: {e}; "
                f"stderr={result.stderr.strip()}; stdout={result.stdout[:400]}"
            )
        if payload.get("status") != 0:
            msg = payload.get("message") or payload.get("data", {}).get("message") or str(payload)
            raise RuntimeError(f"SOQL error: {msg.strip()}")
        return payload["result"]

    def describe(self, object_api: str) -> dict:
        """Run `sf sobject describe` for the given object and return parsed JSON."""
        cmd = [
            "sf", "sobject", "describe",
            "--sobject", object_api,
            "--target-org", self.org,
            "--json",
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        try:
            payload = json.loads(result.stdout)
        except json.JSONDecodeError as e:
            raise RuntimeError(
                f"sf sobject describe produced unparseable stdout: {e}; "
                f"stderr={result.stderr.strip()}; stdout={result.stdout[:400]}"
            )
        if payload.get("status") != 0:
            msg = payload.get("message") or str(payload)
            raise RuntimeError(f"describe error: {msg.strip()}")
        return payload["result"]


# ---------- Phase 2/3: SFDC field metadata + Fill % ----------

def fetch_fielddef_for_object(sf: SfClient, object_api: str) -> list[dict]:
    """Return all FieldDefinition records for a single object.

    Includes DurableId — its format `<ObjectAPI>.<URLSegment>` gives the
    exact URL path segment SFDC's Object Manager expects, so we don't have
    to guess (it covers custom field 15-char IDs, relationship names, etc.).
    """
    cache_name = f"sfdc_fielddef_{object_api}.json"
    cached = cache_load(cache_name)
    if cached is not None and cached and "DurableId" in cached[0] and "LastModifiedDate" in cached[0]:
        return cached
    query = (
        "SELECT QualifiedApiName, DurableId, Label, DataType, Description, "
        "IsCalculated, LastModifiedDate "
        "FROM FieldDefinition "
        f"WHERE EntityDefinition.QualifiedApiName = '{object_api}'"
    )
    res = sf.soql(query, tooling=True)
    records = res.get("records", [])
    cache_save(cache_name, records)
    return records


def fetch_total_count(sf: SfClient, object_api: str) -> int | None:
    cache_name = f"sfdc_total_{object_api}.json"
    cached = cache_load(cache_name)
    if cached is not None:
        return cached.get("total")
    res = sf.soql(f"SELECT COUNT() FROM {object_api}")
    total = res.get("totalSize")
    cache_save(cache_name, {"total": total})
    return total


NON_AGGREGATABLE_DATATYPE_PATTERN = re.compile(
    r"^(Address|Geolocation|Location|Base64|EncryptedText|TextArea)\b",
    re.I,
)


def is_aggregatable(field_def: dict) -> bool:
    dt = (field_def.get("DataType") or "").strip()
    # SOQL COUNT(field) does not support compound and a few special types.
    if NON_AGGREGATABLE_DATATYPE_PATTERN.search(dt):
        return False
    # Address compound fields show as just "Address"; their components
    # (BillingStreet etc.) are aggregatable normal Text.
    return True


def fetch_fill_counts(sf: SfClient, object_api: str, fielddefs_by_api: dict[str, dict],
                       api_names: list[str], chunk: int = 80) -> dict[str, int]:
    """Compute COUNT(field) for many fields in one query (SOQL aggregates).

    Returns {api_name: non_null_count}. Skips compound/non-aggregatable fields.
    Falls back to per-field queries if a chunk fails so a single bad field
    doesn't lose the rest.
    """
    cache_name = f"sfdc_fill_{object_api}.json"
    cached = cache_load(cache_name) or {}
    out = dict(cached)
    # Filter out non-aggregatable + already-cached.
    todo: list[str] = []
    for a in api_names:
        if a in out:
            continue
        fd = fielddefs_by_api.get(a)
        if fd and not is_aggregatable(fd):
            out[a] = None  # explicitly mark — won't render Fill % chip
            continue
        todo.append(a)
    cache_save(cache_name, out)
    if not todo:
        return out
    i = 0
    while i < len(todo):
        batch = todo[i:i + chunk]
        count_exprs = ", ".join(f"COUNT({a}) c_{n}" for n, a in enumerate(batch))
        query = f"SELECT {count_exprs} FROM {object_api}"
        try:
            res = sf.soql(query)
            records = res.get("records") or []
            if records:
                row = records[0]
                for n, a in enumerate(batch):
                    val = row.get(f"c_{n}")
                    out[a] = val if isinstance(val, int) else None
            else:
                for a in batch:
                    out[a] = None
            cache_save(cache_name, out)
            print(f"[fill] {object_api}: {min(i + chunk, len(todo))}/{len(todo)}")
            i += chunk
        except Exception as e:
            if len(batch) == 1:
                # Single bad field: skip it and move on.
                print(f"[fill] {object_api}: skipping {batch[0]} ({e})")
                out[batch[0]] = None
                cache_save(cache_name, out)
                i += 1
            else:
                # Split chunk in half and retry to isolate the bad field.
                half = max(1, len(batch) // 2)
                print(f"[fill] {object_api}: chunk {i}-{i+len(batch)} failed, splitting to {half}")
                chunk = half
    return out


def fetch_custom_field_dates(sf: SfClient, refresh: bool = False) -> dict[str, str]:
    """Fetch CustomField.CreatedDate keyed by the 15-char custom-field ID.

    FieldDefinition lacks a real CreatedDate, so we join via the underlying
    CustomField metadata object. CustomField.Id is 18 chars; DurableId on
    FieldDefinition is the 15-char prefix. We key by the 15-char form so the
    lookup is straightforward in build_model.
    """
    cache_name = "sfdc_custom_field_dates.json"
    if not refresh:
        cached = cache_load(cache_name)
        if cached is not None:
            return cached
    # TableEnumOrId is the API name for STANDARD objects but a CustomObject Id
    # (01I…/01IRc…) for CUSTOM objects, so we can't filter by our object names
    # cleanly. Query everything and match by ID prefix during lookup.
    query = "SELECT Id, CreatedDate FROM CustomField"
    out: dict[str, str] = {}
    try:
        res = sf.soql(query, tooling=True)
        for r in res.get("records", []):
            full_id = r.get("Id") or ""
            if not full_id:
                continue
            # 15-char ID prefix matches FieldDefinition.DurableId's id suffix.
            out[full_id[:15]] = r.get("CreatedDate") or ""
    except Exception as e:
        print(f"[customfield-dates] query failed: {e}")
    cache_save(cache_name, out)
    return out


def fetch_active_flows(sf: SfClient, refresh: bool = False) -> list[dict]:
    """Return all Active flows via the Tooling API.

    Joins Flow to its FlowDefinition for the developer name + def id (used to
    build the SFDC Setup URL).
    """
    cache_name = "sfdc_flows.json"
    if not refresh:
        cached = cache_load(cache_name)
        # Existing cache predates CreatedDate; re-fetch when missing.
        if cached is not None and cached and "CreatedDate" in cached[0]:
            return cached
    query = (
        "SELECT Id, MasterLabel, Definition.Id, Definition.DeveloperName, "
        "ProcessType, Status, ApiVersion, VersionNumber, Description, "
        "CreatedDate, LastModifiedDate "
        "FROM Flow WHERE Status = 'Active' ORDER BY MasterLabel"
    )
    try:
        res = sf.soql(query, tooling=True)
    except Exception as e:
        print(f"[flows] query failed: {e}")
        cache_save(cache_name, [])
        return []
    records = res.get("records", [])
    cache_save(cache_name, records)
    return records


def fetch_picklist_values(sf: SfClient, object_api: str) -> dict[str, list[dict]]:
    """Return {field_api_name: [{value,label,active}, ...]} for picklist fields
    on the given object, sourced from `sf sobject describe`."""
    cache_name = f"sfdc_picklist_{object_api}.json"
    cached = cache_load(cache_name)
    if cached is not None:
        return cached
    try:
        descr = sf.describe(object_api)
    except Exception as e:
        print(f"[picklist] {object_api} describe failed: {e}")
        cache_save(cache_name, {})
        return {}
    out: dict[str, list[dict]] = {}
    for f in descr.get("fields", []) or []:
        # picklist / multipicklist fields
        opts = f.get("picklistValues") or []
        if not opts:
            continue
        out[f.get("name")] = [
            {
                "value": o.get("value"),
                "label": o.get("label"),
                "active": bool(o.get("active")),
            }
            for o in opts
        ]
    cache_save(cache_name, out)
    return out


# ---------- Label → API name matching ----------

def _label_norm(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def match_label_to_field(label: str, fielddefs: list[dict]) -> dict | None:
    """Find the FieldDefinition record whose Label best matches `label`.

    Strategy: exact label (case-insensitive normalized) first; fall back to
    QualifiedApiName equality; fall back to label-norm substring match.
    """
    if not label:
        return None
    target = _label_norm(label)
    # exact normalized
    for f in fielddefs:
        if _label_norm(f.get("Label", "")) == target:
            return f
    # exact api name
    for f in fielddefs:
        if (f.get("QualifiedApiName") or "").lower() == label.lower():
            return f
    # contains
    for f in fielddefs:
        nl = _label_norm(f.get("Label", ""))
        if target and (target in nl or nl in target):
            return f
    return None


# ---------- model assembly ----------

def fill_chip(pct: int | None, total: int | None) -> dict:
    if pct is None:
        return {"pct_text": "—", "band": "na", "title": ""}
    if total is not None and total < 100:
        return {"pct_text": f"{pct}%*", "band": _band(pct), "title": f"based on {total} records"}
    return {"pct_text": f"{pct}%", "band": _band(pct), "title": ""}


def _band(pct: int) -> str:
    if pct >= 80:
        return "high"
    if pct >= 50:
        return "mid"
    return "low"


SYSTEM_AUDIT_FIELDS = {
    "CreatedById", "LastModifiedById", "MasterRecordId",
    "UserRecordAccessId", "RecordVisibilityId", "OperatingHoursId",
    "ConnectionReceivedId", "ConnectionSentId", "Jigsaw", "JigsawCompanyId",
    "PhotoUrl", "IsDeleted", "SystemModstamp",
}

# Fields without a real `/FieldsAndRelationships/<X>/view` page in Salesforce
# Object Manager. Hitting these URLs returns "Insufficient Privileges". We
# render the label as plain text (no link) for these.
NO_SETUP_PAGE_FIELDS = {
    # Audit / system metadata
    "Id", "IsDeleted", "SystemModstamp",
    "CreatedDate", "CreatedById",
    "LastModifiedDate", "LastModifiedById",
    "LastViewedDate", "LastReferencedDate", "LastActivityDate",
    "MasterRecordId",
    "ConnectionReceivedId", "ConnectionSentId",
    "UserRecordAccessId", "RecordVisibilityId", "OperatingHoursId",
    "Jigsaw", "JigsawCompanyId", "DandbCompanyId",
    "PhotoUrl", "CompareName", "CleanStatus",
    # Note: Address compound parents (BillingAddress, ShippingAddress,
    # MailingAddress) DO have setup pages — confirmed by URL test. Don't add.
}
SYSTEM_LOOKUP_TARGETS = {
    "User Record Access", "Record Visibility", "Operating Hours",
    "Connection", "Group",
}
EXTERNAL_ID_SUFFIX_RE = re.compile(r"_ID__c$", re.I)
LOOKUP_TARGET_RE = re.compile(r"^(Lookup|MasterDetail)\(([^)]*)\)$")


_TOKEN_STOPWORDS = {
    "id", "the", "and", "of", "or", "for", "to", "from", "in", "on",
    "by", "with", "name", "date", "field", "count", "total",
    "data", "info", "status", "type",  # too common to use as a similarity signal
}
_TOKEN_SPLIT_RE = re.compile(r"[\s_\-/]+")
_CAMEL_BOUNDARY_RE = re.compile(r"(?<=[a-z])(?=[A-Z])")


def _label_tokens(label: str) -> set[str]:
    if not label:
        return set()
    # Split CamelCase, then split on common separators.
    s = _CAMEL_BOUNDARY_RE.sub(" ", label)
    s = _TOKEN_SPLIT_RE.sub(" ", s)
    return {
        t.lower()
        for t in s.split()
        if len(t) >= 3 and t.lower() not in _TOKEN_STOPWORDS and not t.isdigit()
    }


def find_similar_fields(rows: list[dict]) -> dict[str, list[str]]:
    """Group fields by overlapping label tokens. Returns {field_label: [other_labels...]}.

    A field has similar siblings if it shares ≥2 meaningful tokens with another field
    in the same object, OR shares ≥1 token AND that token is a distinctive product/system
    name (Stripe, Vitally, Healthie, Campaign, Calendly, AskNicely, Chameleon, Trial,
    Renewal, Onboarding).
    """
    distinctive = {
        "stripe", "vitally", "healthie", "campaign", "calendly", "asknicely",
        "chameleon", "trial", "renewal", "onboarding", "csat", "nps",
        "implementation", "subscription", "plan", "churn", "intacct",
        "taskray", "mobile", "marketplace", "mrr", "arr",
    }
    token_index: list[tuple[str, set[str]]] = []
    for r in rows:
        toks = _label_tokens(r["label"])
        token_index.append((r["label"], toks))
    similar: dict[str, set[str]] = {}
    for i in range(len(token_index)):
        a_label, a_toks = token_index[i]
        if not a_toks:
            continue
        for j in range(i + 1, len(token_index)):
            b_label, b_toks = token_index[j]
            shared = a_toks & b_toks
            if not shared:
                continue
            if len(shared) >= 2 or (len(shared) >= 1 and (shared & distinctive)):
                similar.setdefault(a_label, set()).add(b_label)
                similar.setdefault(b_label, set()).add(a_label)
    # Sort sibling lists by label for stable output.
    return {k: sorted(v) for k, v in similar.items()}


def setup_url_segment(durable_id: str, api_name: str, data_type: str) -> str:
    """Compute the URL path segment for /FieldsAndRelationships/<X>/view.

    SFDC's FieldDefinition.DurableId is in the form `<Object>.<Segment>` and
    the `<Segment>` is the EXACT path SFDC's Object Manager uses. Examples:
      - Account.Name              → "Name"            (standard field)
      - Account.RecordType        → "RecordType"      (Record Type id)
      - Account.Owner             → "Owner"           (Lookup(User) id)
      - Account.00NRc00000rWl8X   → "00NRc00000rWl8X" (custom field id)
    Using DurableId removes all guesswork. Falls back to heuristic when
    DurableId is missing (e.g., older cache files).
    """
    if durable_id and "." in durable_id:
        return durable_id.split(".", 1)[1]
    # Heuristic fallback when DurableId isn't available.
    if not api_name:
        return ""
    if api_name.endswith("__c"):
        return api_name
    if api_name.endswith("Id") and api_name != "Id":
        return api_name[:-2]
    return api_name


def classify_key(api_name: str, data_type: str) -> tuple[str | None, str | None]:
    """Return (kind, target) when the field is a cross-object key, else (None, None).
    kind ∈ {'pk', 'lookup', 'external_id'}.
    """
    if not api_name:
        return None, None
    if api_name in SYSTEM_AUDIT_FIELDS:
        return None, None
    if api_name == "Id":
        return "pk", "self"
    if EXTERNAL_ID_SUFFIX_RE.search(api_name):
        return "external_id", api_name
    m = LOOKUP_TARGET_RE.match((data_type or "").strip())
    if m:
        target = m.group(2).strip()
        # Exclude lookups whose ONLY target is a system meta-object.
        # Lookup(User,Group) on OwnerId IS meaningful (ownership) — keep it.
        targets = [t.strip() for t in target.split(",") if t.strip()]
        if not targets:
            return None, None
        if all(t in SYSTEM_LOOKUP_TARGETS for t in targets):
            return None, None
        return "lookup", target or "unknown"
    return None, None


def sens_class(sensitivity: str) -> str:
    s = (sensitivity or "").strip().lower()
    if "pii" in s:
        return "pii"
    if "financial id" in s:
        return "fin-id"
    if "financial" in s:
        return "fin"
    if s == "internal":
        return "internal"
    return "other"


_INLINE_PICKLIST_SENTENCE = re.compile(
    r"\s*Active picklist values:[^.]*\.\s*",
    re.IGNORECASE,
)

# Free-text replacements applied to every description so renamed-but-still-referenced
# field names normalize across V5 docx and SFDC FieldDefinition.Description text.
DESCRIPTION_TEXT_REPLACEMENTS = [
    ("Product Health Score", "Health Score"),
]


def strip_inline_picklist_values(description: str) -> str:
    """V5 docx baked picklist values into descriptions ('Active picklist values: ...').
    We render live SFDC values in a separate row, so remove the embedded list."""
    if not description:
        return description
    return _INLINE_PICKLIST_SENTENCE.sub(" ", description).strip()


def normalize_description_text(description: str) -> str:
    """Apply free-text renames so SFDC's stale description prose tracks current
    field labels (e.g., "Product Health Score" was renamed to "Health Score")."""
    if not description:
        return description
    for old, new in DESCRIPTION_TEXT_REPLACEMENTS:
        description = description.replace(old, new)
    return description


from datetime import timedelta

NEW_FIELD_WINDOW_DAYS = 180  # rolling 6 months
NEW_FIELD_CUTOFF = datetime.now(timezone.utc) - timedelta(days=NEW_FIELD_WINDOW_DAYS)
NEW_FIELD_CUTOFF_LABEL = NEW_FIELD_CUTOFF.strftime("%Y-%m-%d")


def lookup_created_date(durable_id: str, custom_field_dates: dict) -> str:
    """Return the ISO 8601 CreatedDate for a custom field, or '' if not custom."""
    if not durable_id or "." not in durable_id:
        return ""
    cf_id = durable_id.split(".", 1)[1]
    if not cf_id.startswith("00N"):
        return ""
    return custom_field_dates.get(cf_id, "") or ""


def parse_iso_created(created: str) -> datetime | None:
    """Parse SFDC's ISO 8601 CreatedDate string into a tz-aware datetime."""
    if not created:
        return None
    try:
        # Normalize the +0000 / -0000 offset that fromisoformat can't parse pre-3.11.
        normalized = created.replace("+0000", "+00:00").replace("-0000", "+00:00")
        if normalized.endswith("Z"):
            normalized = normalized[:-1] + "+00:00"
        dt = datetime.fromisoformat(normalized)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except (ValueError, TypeError):
        return None


def is_new_field(durable_id: str, custom_field_dates: dict) -> bool:
    """Field is "new" if its CustomField CreatedDate falls within the rolling
    NEW_FIELD_WINDOW_DAYS-day window. Drops off automatically as fields age."""
    created = lookup_created_date(durable_id, custom_field_dates)
    dt = parse_iso_created(created)
    return dt is not None and dt >= NEW_FIELD_CUTOFF


def build_flows_model(flow_records: list[dict]) -> list[dict]:
    """Shape raw Flow records into render-ready dicts with inferred trigger object."""
    out = []
    known_objects = set(OBJECT_API_NAME)  # labels: 'Account', 'Contact', ...
    for r in flow_records or []:
        defn = r.get("Definition") or {}
        dev = (defn.get("DeveloperName") or "").strip()
        # First underscore-separated token tends to be the triggering object.
        first = dev.split("_", 1)[0] if dev else ""
        inferred = ""
        for obj_label in known_objects:
            if first.lower() == obj_label.lower().split()[0]:
                inferred = obj_label
                break
        modified = (r.get("LastModifiedDate") or "")[:10]
        created = (r.get("CreatedDate") or "")[:10]
        out.append({
            "label": r.get("MasterLabel") or dev or "—",
            "dev_name": dev or "—",
            "def_id": defn.get("Id") or "",
            "process_type": r.get("ProcessType") or "—",
            "version": r.get("VersionNumber"),
            "api_version": r.get("ApiVersion"),
            "modified": modified,
            "created": created,
            "description": (r.get("Description") or "").strip(),
            "inferred_object": inferred,
            "inferred_object_color": OBJECT_COLOR.get(inferred, "#94a3b8"),
            "similar_to": [],  # filled below
        })
    out.sort(key=lambda f: f["label"].lower())
    # Cluster flows with similar developer-name tokens within the same inferred object
    # (so we don't end up with 145-flow "everyone is similar" tooltips).
    by_object: dict[str, list[dict]] = {}
    for f in out:
        by_object.setdefault(f["inferred_object"] or "(other)", []).append(f)
    for group in by_object.values():
        # Reuse the field similarity logic; it tokenizes labels and compares.
        # Build a label-keyed similar map then attach to each flow.
        sims = find_similar_fields([{"label": f["label"]} for f in group])
        for f in group:
            f["similar_to"] = sims.get(f["label"], [])
    return out


def build_model(v5: dict, sf_data: dict) -> dict:
    """Combine V5 parse with SFDC metadata into a render-ready dict."""
    objects = []
    custom_field_dates = sf_data.get("custom_field_dates", {})
    for obj in v5["objects"]:
        api = OBJECT_API_NAME[obj["label"]]
        fielddefs = sf_data["fielddefs"].get(api, [])
        fill_counts = sf_data["fills"].get(api, {})
        total = sf_data["totals"].get(api)
        picklists = sf_data["picklists"].get(api, {})
        sections = []
        unmatched = []
        # First pass — assemble rows so we can compute object-wide similarity.
        for sec in obj["sections"]:
            rows = []
            for f in sec["fields"]:
                fd = match_label_to_field(f["label"], fielddefs)
                api_name = (fd or {}).get("QualifiedApiName") or ""
                if not api_name:
                    unmatched.append(f["label"])
                sf_data_type = (fd or {}).get("DataType") or ""
                sf_description = (fd or {}).get("Description") or ""
                description = sf_description.strip() or f["description_v5"]
                description = strip_inline_picklist_values(description)
                description = normalize_description_text(description)
                non_null = fill_counts.get(api_name) if api_name else None
                if non_null is not None and total:
                    pct = round(100.0 * non_null / total)
                else:
                    pct = None
                fchip = fill_chip(pct, total)
                opts = picklists.get(api_name, [])
                active_values = [o for o in opts if o.get("active")]
                key_kind, key_target = classify_key(api_name, sf_data_type)
                durable_id = (fd or {}).get("DurableId") or ""
                # Only render a link when Object Manager actually has a setup page
                # for this field. Compound parents and audit fields get plain text.
                if api_name and api_name not in NO_SETUP_PAGE_FIELDS:
                    segment = setup_url_segment(durable_id, api_name, sf_data_type)
                    setup_url = f"{SFDC_SETUP_BASE}/{api}/FieldsAndRelationships/{segment}/view"
                else:
                    setup_url = ""
                is_new = is_new_field(durable_id, custom_field_dates)
                created_iso = lookup_created_date(durable_id, custom_field_dates)
                created_date = (parse_iso_created(created_iso) or "")
                created_display = created_date.strftime("%Y-%m-%d") if created_date else ""
                # SFDC's live FieldDefinition.Label is authoritative for display
                # when we matched. Falls back to V5's label only when unmatched.
                sf_label = ((fd or {}).get("Label") or "").strip()
                display_label = sf_label or f["label"]
                rows.append({
                    "label": display_label,
                    "api_name": api_name or "—",
                    "setup_url": setup_url,
                    "type": sf_data_type or f["type_v5"] or "—",
                    "fill": fchip,
                    "source": f["source"] or "—",
                    "owner": f["owner"] or "—",
                    "sensitivity": f["sensitivity"] or "—",
                    "sensitivity_class": sens_class(f["sensitivity"]),
                    "description": description or "—",
                    "active_values": active_values,
                    "dependencies": f["dependencies"] or "—",
                    "is_key": key_kind is not None,
                    "key_kind": key_kind,
                    "key_target": key_target,
                    "is_new": is_new,
                    "created": created_display,
                    "similar_to": [],  # filled in below
                })
            sections.append({
                "title": sec["title"],
                "rows": rows,
            })
        # Second pass — compute similarity over ALL rows in this object.
        all_rows = [r for s in sections for r in s["rows"]]
        similar = find_similar_fields(all_rows)
        for r in all_rows:
            r["similar_to"] = similar.get(r["label"], [])
        # Build the "All Other Fields" list — everything in FieldDefinition that
        # the V5 dictionary doesn't already document. These rows carry fewer
        # attributes because we don't have V5 prose (Source/Owner/Sensitivity);
        # we just surface what SFDC's metadata gives us.
        curated_apis = {r["api_name"] for r in all_rows if r["api_name"] and r["api_name"] != "—"}
        other_rows = []
        for fd in fielddefs:
            api_name = (fd.get("QualifiedApiName") or "").strip()
            if not api_name or api_name in curated_apis:
                continue
            sf_label = (fd.get("Label") or "").strip() or api_name
            sf_data_type = (fd.get("DataType") or "").strip()
            description = (fd.get("Description") or "").strip()
            description = strip_inline_picklist_values(description)
            description = normalize_description_text(description)
            durable_id = fd.get("DurableId") or ""
            if api_name not in NO_SETUP_PAGE_FIELDS:
                segment = setup_url_segment(durable_id, api_name, sf_data_type)
                setup_url = f"{SFDC_SETUP_BASE}/{api}/FieldsAndRelationships/{segment}/view"
            else:
                setup_url = ""
            key_kind, key_target = classify_key(api_name, sf_data_type)
            is_new = is_new_field(durable_id, custom_field_dates)
            created_iso = lookup_created_date(durable_id, custom_field_dates)
            created_date = parse_iso_created(created_iso) or ""
            created_display = created_date.strftime("%Y-%m-%d") if created_date else ""
            is_custom = api_name.endswith("__c")
            other_rows.append({
                "label": sf_label,
                "api_name": api_name,
                "type": sf_data_type or "—",
                "description": description,
                "created": created_display,
                "is_custom": is_custom,
                "is_new": is_new,
                "is_key": key_kind is not None,
                "key_kind": key_kind,
                "key_target": key_target,
                "setup_url": setup_url,
            })
        other_rows.sort(key=lambda r: r["label"].lower())

        objects.append({
            "label": obj["label"],
            "api_name": api,
            "color": OBJECT_COLOR.get(obj["label"], "#4f46e5"),
            "total_records": total,
            "sections": sections,
            "other_fields": other_rows,
            "field_count": sum(len(s["rows"]) for s in sections),
            "picklist_count": sum(1 for s in sections for r in s["rows"] if r["active_values"]),
            "key_count": sum(1 for r in all_rows if r["is_key"]),
            "new_count": sum(1 for r in all_rows if r["is_new"]),
            "similar_count": sum(1 for r in all_rows if r["similar_to"]),
            "other_count": len(other_rows),
            "unmatched": unmatched,
        })
    return {
        "objects": objects,
        "flows": build_flows_model(sf_data.get("flows") or []),
        "front_matter": v5["front_matter"],
        "appendix": v5["appendix"],
        "meta": v5["meta"],
    }


# ---------- HTML rendering ----------

def H(s: str) -> str:
    return html.escape(s or "", quote=True)


def _human_count(n: int | None) -> str:
    if not isinstance(n, int):
        return "—"
    if n >= 1_000_000:
        return f"{n/1_000_000:.1f}M"
    if n >= 1_000:
        return f"{n/1_000:.1f}K"
    return f"{n}"


# JSON-island key compression. The export data is repeated 292x (fields) + 145x
# (flows), and each full key like "field_label" or "similar_fields" eats real
# bytes against HubSpot's 1.5 MiB published source-code limit. We emit records
# with short keys and the JS unmaps them on parse via FIELD_KEY_MAP /
# FLOW_KEY_MAP (kept in sync below in _JS).
EXPORT_FIELD_KEY_MAP = {
    "object": "o", "subsection": "ss", "field_label": "l", "api_name": "a",
    "type": "t", "fill_pct": "f", "created_date": "c", "is_new": "n",
    "is_key": "k", "key_kind": "kk", "key_target": "kt", "definition": "d",
    "active_values": "v", "source_update": "su", "owner": "ow",
    "sensitivity": "se", "dependencies": "de", "similar_fields": "sf",
}
EXPORT_FLOW_KEY_MAP = {
    "flow_label": "l", "developer_name": "dn", "triggering_object": "to",
    "process_type": "pt", "version": "ver", "api_version": "av",
    "created_date": "c", "last_modified": "lm", "description": "d",
    "similar_flows": "sf",
}


def _compact_rec(rec: dict, key_map: dict) -> dict:
    """Apply key_map and drop empty values to keep the JSON island small."""
    out = {}
    for k, v in rec.items():
        if v in ("", False, None, [], {}):
            continue
        out[key_map.get(k, k)] = v
    return out


def _build_export_data(model: dict) -> dict:
    """Flatten the render model into export-friendly arrays: one entry per
    field row and one per flow row. Used by the Export dropdown to produce
    CSV / XLSX / TSV downloads from a single JSON data island."""
    fields = []
    for obj in model["objects"]:
        for sec in obj["sections"]:
            for r in sec["rows"]:
                active_values = ", ".join(
                    (v.get("label") or v.get("value") or "")
                    for v in (r.get("active_values") or [])
                    if (v.get("label") or v.get("value"))
                )
                fields.append(_compact_rec({
                    "object": obj["label"],
                    "subsection": sec["title"],
                    "field_label": r["label"],
                    "api_name": r["api_name"] if r["api_name"] != "—" else "",
                    "type": r["type"] if r["type"] != "—" else "",
                    "fill_pct": r["fill"]["pct_text"] if r["fill"]["pct_text"] != "—" else "",
                    "created_date": r.get("created", ""),
                    "is_new": bool(r.get("is_new")),
                    "is_key": bool(r.get("is_key")),
                    "key_kind": r.get("key_kind") or "",
                    "key_target": r.get("key_target") or "",
                    "definition": r["description"] if r["description"] != "—" else "",
                    "active_values": active_values,
                    "source_update": r["source"] if r["source"] != "—" else "",
                    "owner": r["owner"] if r["owner"] != "—" else "",
                    "sensitivity": r["sensitivity"] if r["sensitivity"] != "—" else "",
                    "dependencies": r["dependencies"] if r["dependencies"] != "—" else "",
                    "similar_fields": "; ".join(r.get("similar_to") or []),
                }, EXPORT_FIELD_KEY_MAP))
    # Note: "All Other Fields" data is NOT included in the JSON island — the
    # JSON island already approaches HubSpot's 1.5 MiB template limit. The
    # All Other Fields export walks the rendered DOM rows instead (each row
    # carries enough data attributes for that to be straightforward).
    flows = []
    for f in (model.get("flows") or []):
        flows.append(_compact_rec({
            "flow_label": f["label"],
            "developer_name": f["dev_name"] if f["dev_name"] != "—" else "",
            "triggering_object": f.get("inferred_object") or "",
            "process_type": f["process_type"] if f["process_type"] != "—" else "",
            "version": f.get("version") if f.get("version") is not None else "",
            "api_version": f.get("api_version") if f.get("api_version") is not None else "",
            "created_date": f.get("created") or "",
            "last_modified": f.get("modified") or "",
            "description": f.get("description") or "",
            "similar_flows": "; ".join(f.get("similar_to") or []),
        }, EXPORT_FLOW_KEY_MAP))
    return {"fields": fields, "flows": flows}


def render_html(model: dict, run_meta: dict) -> str:
    total_fields = sum(o["field_count"] for o in model["objects"])
    data_sources_count = len(model["front_matter"].get("data_sources") or [])
    total_records = sum(o["total_records"] for o in model["objects"] if isinstance(o["total_records"], int))
    export_data = _build_export_data(model)

    objects_nav = "".join(
        f'<button class="dd-tab" data-tab="obj-{i}" data-obj="{H(o["label"])}" '
        f'style="--obj-color: {o["color"]};">'
        f'<span class="dd-tab-dot" style="background:{o["color"]};"></span>'
        f'<span class="dd-tab-label">{H(o["label"])}</span>'
        f'<span class="dd-tab-count" data-count-for="obj-{i}">{o["field_count"]}</span>'
        f"</button>"
        for i, o in enumerate(model["objects"])
    )

    objects_panels = "".join(_render_object_panel(i, o) for i, o in enumerate(model["objects"]))

    front_matter_html = _render_front_matter(model["front_matter"])
    appendix_html = _render_appendix(model["appendix"])
    flows_html = _render_flows_section(model.get("flows") or [])
    other_html = _render_other_fields_section(model["objects"])
    total_other = sum(o.get("other_count", 0) for o in model["objects"])

    css = _CSS
    js = _JS
    return f"""<div class="dd-root" id="dd-root">
<style>{css}</style>
<nav class="dd-navbar">
  <div class="dd-nav-brand">Healthie</div>
  <div class="dd-nav-breadcrumb">/ Salesforce Data Dictionary</div>
  <div class="dd-nav-right">
    <div class="dd-export-menu">
      <button type="button" class="dd-nav-link dd-nav-link-slate" id="dd-export-btn" aria-haspopup="true" aria-expanded="false">Export ↓</button>
      <div class="dd-export-pop" id="dd-export-pop" hidden>
        <button type="button" class="dd-export-item" data-export="xlsx">
          <span class="dd-export-title">Excel workbook (.xlsx)</span>
          <span class="dd-export-sub">Multi-sheet — one sheet per object plus Flows.</span>
        </button>
        <button type="button" class="dd-export-item" data-export="csv">
          <span class="dd-export-title">CSV (fields)</span>
          <span class="dd-export-sub">Flat file with Object + Subsection columns.</span>
        </button>
        <button type="button" class="dd-export-item" data-export="csv-flows">
          <span class="dd-export-title">CSV (flows)</span>
          <span class="dd-export-sub">All active SFDC flows.</span>
        </button>
        <button type="button" class="dd-export-item" data-export="csv-other">
          <span class="dd-export-title">CSV (all other fields)</span>
          <span class="dd-export-sub">Every FieldDefinition not in the curated dictionary.</span>
        </button>
        <button type="button" class="dd-export-item" data-export="tsv-clipboard">
          <span class="dd-export-title">Copy for Google Sheets</span>
          <span class="dd-export-sub">TSV to clipboard — paste into a blank sheet.</span>
        </button>
      </div>
    </div>
    <a href="#dd-other-fields" class="dd-nav-link dd-nav-link-slate-2">All Fields ({total_other}) ↓</a>
    <a href="#dd-flows" class="dd-nav-link dd-nav-link-amber">Flows ({len(model['flows'])}) ↓</a>
    <a href="https://go.gethealthie.com/peapod-hubspot-data-dictionary?hsLang=en" class="dd-nav-link dd-nav-link-purple">← HubSpot Dictionary</a>
    <a href="https://go.gethealthie.com/peapod-data-governance?hsLang=en" class="dd-nav-link dd-nav-link-green">Data Governance →</a>
    <span class="dd-nav-stats">{total_fields} fields &nbsp;·&nbsp; {len(model['objects'])} objects &nbsp;·&nbsp; {_human_count(total_records)} records &nbsp;·&nbsp; v6 updated {H(run_meta['generated_at_date'])}</span>
  </div>
</nav>
<script type="application/json" id="dd-export-data">{json.dumps(export_data, separators=(",", ":"), ensure_ascii=False)}</script>

<section class="dd-hero">
  <div class="dd-hero-inner">
    <h1 class="dd-hero-title">Salesforce Data Dictionary</h1>
    <p class="dd-hero-sub">Live field-level reference for the 7 Salesforce objects RevOps maintains, sourced from <code>FieldDefinition</code> and pulled at run time.</p>
    <div class="dd-kpis">
      <div class="dd-kpi"><div class="dd-kpi-num">{total_fields}</div><div class="dd-kpi-label">Fields documented</div></div>
      <div class="dd-kpi"><div class="dd-kpi-num">{len(model['objects'])}</div><div class="dd-kpi-label">Objects</div></div>
      <div class="dd-kpi"><div class="dd-kpi-num">{data_sources_count}</div><div class="dd-kpi-label">Integrated sources</div></div>
      <div class="dd-kpi"><div class="dd-kpi-num">v6</div><div class="dd-kpi-label">{H(run_meta['generated_at_date'])}</div></div>
    </div>
  </div>
</section>

<section class="dd-controls">
  <div class="dd-controls-inner">
    <input id="dd-search" class="dd-search" type="search" placeholder="Search every field across every object (Field Label, API Name, Description, Active Values…)" autocomplete="off" />
    <div class="dd-filters">
      <div class="dd-filter-group" role="group" aria-label="Similar fields filter">
        <span class="dd-filter-label">Similar:</span>
        <button type="button" class="dd-seg-btn is-active" data-similar-filter="all">All</button>
        <button type="button" class="dd-seg-btn" data-similar-filter="similar">⚠ Has similar</button>
        <button type="button" class="dd-seg-btn" data-similar-filter="unique">Unique</button>
      </div>
      <div class="dd-filter-group">
        <span class="dd-filter-label">Created in last:</span>
        <input id="dd-created-days" class="dd-days-input" type="number" min="0" placeholder="N" inputmode="numeric" />
        <span class="dd-filter-label">days</span>
        <button type="button" id="dd-created-clear" class="dd-link-btn">clear</button>
      </div>
    </div>
    <div class="dd-search-meta" id="dd-search-meta">Type to filter. Tab badges update live.</div>
  </div>
</section>

<section class="dd-legend">
  <div class="dd-legend-inner">
    <div class="dd-legend-item"><span class="dd-legend-chip dd-legend-chip-key">★</span><span>Key field — referenced by other objects (lookups, foreign keys, external IDs). Row is shaded.</span></div>
    <div class="dd-legend-item"><span class="dd-legend-chip dd-legend-chip-new">NEW</span><span>Field created in the past 6 months (after {NEW_FIELD_CUTOFF_LABEL}).</span></div>
    <div class="dd-legend-item"><span class="dd-legend-chip dd-legend-chip-warn">⚠</span><span>Similar fields exist in this object — hover for the related field names.</span></div>
    <div class="dd-legend-item"><span class="dd-chip dd-chip-high">≥80%</span><span>High fill</span></div>
    <div class="dd-legend-item"><span class="dd-chip dd-chip-mid">50–79%</span><span>Mid fill</span></div>
    <div class="dd-legend-item"><span class="dd-chip dd-chip-low">&lt;50%</span><span>Low fill</span></div>
  </div>
</section>

<nav class="dd-tabs" id="dd-tabs">
  {objects_nav}
</nav>

<main class="dd-panels" id="dd-panels">
  {objects_panels}
</main>

{other_html}

{flows_html}

<section class="dd-accordion-group">
  <h2 class="dd-section-title">How to Use This Dictionary</h2>
  {front_matter_html}
</section>

<section class="dd-accordion-group">
  <h2 class="dd-section-title">Appendix</h2>
  {appendix_html}
</section>

<footer class="dd-footer">
  Generated {H(run_meta['generated_at'])} from <code>Healthie_SFDC_Data_Dictionary_v5.docx</code> (last modified {H(model['meta']['v5_mtime'])}). Live Salesforce metadata pulled via <code>FieldDefinition</code>, <code>PicklistValueInfo</code>, and <code>COUNT()</code> queries against {H(DEFAULT_ORG)}.
</footer>

<script>{js}</script>
</div>
"""


def _render_object_panel(idx: int, obj: dict) -> str:
    color = obj["color"]
    sections_html = ""
    for sec in obj["sections"]:
        rows_html = "".join(_render_field_row(r) for r in sec["rows"])
        sections_html += f"""
    <section class="dd-subsection" data-subsection="{H(sec['title'])}" style="--obj-color: {color};">
      <h3 class="dd-subsection-title">{H(sec['title'])} <span class="dd-muted">({len(sec['rows'])})</span></h3>
      <div class="dd-table-wrap">
        <table class="dd-table">
          <thead>
            <tr>
              <th class="c-label">Field Label</th>
              <th class="c-api">API Name</th>
              <th class="c-type">Type</th>
              <th class="c-fill">Fill %</th>
              <th class="c-created">Created</th>
              <th class="c-desc">Definition</th>
              <th class="c-values">Active Values</th>
              <th class="c-source">Source / Update</th>
              <th class="c-owner">Owner</th>
              <th class="c-sens">Sensitivity</th>
              <th class="c-deps">Dependencies / Notes</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>
      </div>
    </section>"""
    total = obj["total_records"]
    record_count_text = (f"{total:,} records" if isinstance(total, int) else "record count unavailable")
    return f"""
  <div class="dd-panel" id="obj-{idx}" data-obj="{H(obj['label'])}" style="--obj-color: {color};">
    <header class="dd-object-header">
      <div class="dd-object-band"></div>
      <div class="dd-object-header-body">
        <h2 class="dd-object-title">{H(obj['label'])}</h2>
        <div class="dd-object-meta">
          <strong>{obj['field_count']}</strong> fields
          &nbsp;·&nbsp; <strong>{obj['picklist_count']}</strong> picklists
          &nbsp;·&nbsp; <strong>{obj['key_count']}</strong> key/ID fields
          &nbsp;·&nbsp; {H(record_count_text)}
          &nbsp;·&nbsp; API: <code>{H(obj['api_name'])}</code>
        </div>
      </div>
    </header>
    {sections_html}
  </div>"""


def _render_field_row(r: dict) -> str:
    sens_html = '<span class="dd-muted">—</span>'
    if r["sensitivity"] and r["sensitivity"] != "—":
        sens_html = f'<span class="dd-pill dd-pill-{r["sensitivity_class"]}">{H(r["sensitivity"])}</span>'

    fill_chip = r["fill"]
    fill_html = f'<span class="dd-chip dd-chip-{fill_chip["band"]}" title="{H(fill_chip["title"])}">{H(fill_chip["pct_text"])}</span>'

    desc_html = H(r["description"]) if r["description"] and r["description"] != "—" else '<span class="dd-muted">—</span>'

    values_html = '<span class="dd-muted">—</span>'
    if r["active_values"]:
        vals = ", ".join(H(v.get("label") or v.get("value") or "") for v in r["active_values"] if (v.get("label") or v.get("value")))
        if vals:
            values_html = f'<span class="dd-values">{vals}</span>'

    # NEW chip — green badge for fields created in the rolling 6-month window.
    new_chip = ""
    if r.get("is_new"):
        new_chip = (
            ' <span class="dd-new-chip dd-tip" '
            f'data-tip="New field: created in the past 6 months (after {NEW_FIELD_CUTOFF_LABEL})">'
            'NEW</span>'
        )

    # Key-field star (yellow row). Uses our CSS tooltip, not native `title`.
    key_star = ""
    if r.get("is_key"):
        kind = r.get("key_kind") or ""
        target = r.get("key_target") or ""
        tip_by_kind = {
            "pk": "Primary key — referenced by other objects via Lookup.",
            "lookup": f"Cross-object reference → {target}",
            "external_id": f"External identifier — {target}",
        }
        tip = tip_by_kind.get(kind, "Key field — referenced across objects.")
        key_star = f' <span class="dd-tip dd-key-star" data-tip="{H(tip)}">★</span>'

    # Similar-fields warning — newline-separated bullet list for the tooltip body.
    warn_icon = ""
    if r.get("similar_to"):
        bullets = "\n• ".join(r["similar_to"])
        tip = f"Similar fields in this object:\n• {bullets}"
        warn_icon = f' <span class="dd-tip dd-warn-icon" data-tip="{H(tip)}">⚠</span>'

    row_classes = ["dd-row"]
    if r.get("is_key"):
        row_classes.append("is-key")

    blob = " ".join([
        r["label"], r["api_name"], r["type"], r["source"], r["owner"],
        r["sensitivity"], r["description"],
        ", ".join((v.get("label") or v.get("value") or "") for v in r["active_values"]),
        r["dependencies"],
    ]).lower()

    label_inner = H(r["label"])
    if r.get("setup_url"):
        label_inner = (
            f'<a href="{H(r["setup_url"])}" target="_blank" rel="noopener" '
            f'class="dd-field-link" title="Open in Salesforce Object Manager">'
            f'{H(r["label"])}</a>'
        )

    created_cell = H(r["created"]) if r.get("created") else '<span class="dd-muted">—</span>'
    similar_flag = "1" if r.get("similar_to") else "0"

    return f"""
            <tr class="{' '.join(row_classes)}" data-search="{H(blob)}" data-similar="{similar_flag}" data-created="{H(r.get('created') or '')}">
              <td class="c-label">{label_inner}{new_chip}{warn_icon}</td>
              <td class="c-api"><code>{H(r['api_name'])}</code>{key_star}</td>
              <td class="c-type">{H(r['type'])}</td>
              <td class="c-fill">{fill_html}</td>
              <td class="c-created">{created_cell}</td>
              <td class="c-desc">{desc_html}</td>
              <td class="c-values">{values_html}</td>
              <td class="c-source">{H(r['source'])}</td>
              <td class="c-owner">{H(r['owner'])}</td>
              <td class="c-sens">{sens_html}</td>
              <td class="c-deps">{H(r['dependencies'])}</td>
            </tr>"""


def _render_two_col_table(headers: list[str], rows: list[list[str]]) -> str:
    head = "".join(f"<th>{H(h)}</th>" for h in headers)
    body = "".join(
        "<tr>" + "".join(f"<td>{H(c)}</td>" for c in row) + "</tr>"
        for row in rows
    )
    return f'<table class="dd-table dd-table-narrow"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'


def _render_other_fields_section(objects: list[dict]) -> str:
    """The "All Other Fields" section — every FieldDefinition on each curated
    object that the V5 dictionary doesn't already document. Surfaces only what
    SFDC's metadata gives us (no V5 prose), so columns are simpler than the
    curated tables."""
    total = sum(o.get("other_count", 0) for o in objects)
    if total == 0:
        return ""
    # Tabs for each object that has other_fields
    tabs_html = "".join(
        f'<button class="dd-tab dd-other-tab{" is-active" if i == 0 else ""}" '
        f'data-other-tab="other-{i}" '
        f'style="--obj-color: {o["color"]};">'
        f'<span class="dd-tab-dot" style="background:{o["color"]};"></span>'
        f'<span class="dd-tab-label">{H(o["label"])}</span>'
        f'<span class="dd-tab-count" data-other-count-for="other-{i}">{o.get("other_count", 0)}</span>'
        f"</button>"
        for i, o in enumerate(objects)
    )
    panels_html = "".join(_render_other_panel(i, o) for i, o in enumerate(objects))
    return f"""
<section class="dd-others" id="dd-other-fields">
  <header class="dd-object-header" style="--obj-color: var(--slate-700);">
    <div class="dd-object-band" style="background: var(--slate-700);"></div>
    <div class="dd-object-header-body">
      <h2 class="dd-object-title" style="color: var(--slate-700);">All Other Fields</h2>
      <div class="dd-object-meta">
        <strong>{total}</strong> additional fields across {len(objects)} objects &nbsp;·&nbsp;
        Every <code>FieldDefinition</code> not already documented in the curated dictionary above.
        These rows surface only what SFDC metadata gives us — no V5 prose, no Source/Owner/Sensitivity context.
      </div>
    </div>
  </header>
  <div class="dd-flows-controls">
    <input id="dd-other-search" class="dd-search dd-flow-search" type="search" placeholder="Filter all-other fields by label, API name, type, or description…" autocomplete="off" />
    <div class="dd-filters">
      <div class="dd-filter-group">
        <span class="dd-filter-label">Type:</span>
        <button type="button" class="dd-seg-btn is-active" data-other-type-filter="all">All</button>
        <button type="button" class="dd-seg-btn" data-other-type-filter="custom">Custom (__c)</button>
        <button type="button" class="dd-seg-btn" data-other-type-filter="standard">Standard</button>
      </div>
      <div class="dd-filter-group">
        <span class="dd-filter-label">Created in last:</span>
        <input id="dd-other-created-days" class="dd-days-input" type="number" min="0" placeholder="N" inputmode="numeric" />
        <span class="dd-filter-label">days</span>
        <button type="button" id="dd-other-created-clear" class="dd-link-btn">clear</button>
      </div>
    </div>
    <span id="dd-other-search-meta" class="dd-search-meta">Type to filter. Tab badges update live.</span>
  </div>
  <nav class="dd-tabs dd-other-tabs">{tabs_html}</nav>
  <div class="dd-panels dd-other-panels">{panels_html}</div>
</section>"""


def _render_other_panel(idx: int, obj: dict) -> str:
    rows = obj.get("other_fields") or []
    rows_html = "".join(_render_other_row(r) for r in rows)
    is_active = " is-active" if idx == 0 else ""
    return f"""
<div class="dd-panel dd-other-panel{is_active}" id="other-{idx}" data-obj="{H(obj['label'])}" style="--obj-color: {obj['color']};">
  <div class="dd-panel-summary">
    <div><strong>{obj.get('other_count', 0)}</strong> additional fields on <code>{H(obj['api_name'])}</code> &nbsp;·&nbsp; (also <strong>{obj['field_count']}</strong> curated in the dictionary above)</div>
  </div>
  <div class="dd-table-wrap">
    <table class="dd-table dd-others-table">
      <thead>
        <tr>
          <th class="c-other-label">Field Label</th>
          <th class="c-other-api">API Name</th>
          <th class="c-other-type">Type</th>
          <th class="c-other-flag">Source</th>
          <th class="c-other-created">Created</th>
          <th class="c-other-desc">Description</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
  </div>
</div>"""


def _render_other_row(r: dict) -> str:
    # All Other Fields renders 1,290 rows. HubSpot's published-env source-code
    # limit is 1.5 MiB, so byte efficiency matters: per-<td> classes are
    # omitted (CSS uses nth-child), and pill/muted wrappers use shorter class
    # names. JS reads cells positionally via tr.cells[N].
    label_inner = H(r["label"])
    if r.get("setup_url"):
        # The class="dd-field-link" is intentionally omitted — `.dd-other-row a`
        # in CSS picks up the styling.
        label_inner = (
            f'<a href="{H(r["setup_url"])}" target="_blank" rel="noopener">'
            f'{H(r["label"])}</a>'
        )
    new_chip = ""
    if r.get("is_new"):
        new_chip = (
            ' <span class="dd-new-chip dd-tip" '
            f'data-tip="New field: created in the past 6 months (after {NEW_FIELD_CUTOFF_LABEL})">NEW</span>'
        )
    key_star = ""
    if r.get("is_key"):
        target = r.get("key_target") or ""
        tip = f"Cross-object reference → {target}" if r.get("key_kind") == "lookup" else (
            "Primary key — referenced by other objects via Lookup." if r.get("key_kind") == "pk" else
            f"External identifier — {target}"
        )
        key_star = f' <span class="dd-tip dd-key-star" data-tip="{H(tip)}">★</span>'
    # Pill class names shortened: "dd-pill dd-pill-custom" -> "p p-c"
    source_chip = (
        '<span class="p p-c">Custom</span>' if r.get("is_custom")
        else '<span class="p p-s">Standard</span>'
    )
    created_cell = H(r["created"]) if r.get("created") else "—"
    desc_cell = H(r["description"]) if r["description"] else "—"
    blob = " ".join([r["label"], r["api_name"], r["type"], r["description"]]).lower()
    type_class = "custom" if r.get("is_custom") else "standard"
    return f"""
            <tr class="dd-other-row{' is-key' if r.get('is_key') else ''}" data-search="{H(blob)}" data-type-class="{type_class}" data-created="{H(r.get('created') or '')}">
              <td>{label_inner}{new_chip}</td>
              <td><code>{H(r['api_name'])}</code>{key_star}</td>
              <td>{H(r['type'])}</td>
              <td>{source_chip}</td>
              <td>{created_cell}</td>
              <td>{desc_cell}</td>
            </tr>"""


def _render_flows_section(flows: list[dict]) -> str:
    """Render the SFDC Flows section that lives below the object panels."""
    if not flows:
        return ""
    process_type_counts: dict[str, int] = {}
    for f in flows:
        process_type_counts[f["process_type"]] = process_type_counts.get(f["process_type"], 0) + 1
    process_summary = ", ".join(
        f"{count} {ptype}" for ptype, count in sorted(process_type_counts.items(), key=lambda kv: -kv[1])
    )
    rows_html = "".join(_render_flow_row(f) for f in flows)
    return f"""
<section class="dd-flows" id="dd-flows" style="--obj-color: #d97706;">
  <header class="dd-object-header">
    <div class="dd-object-band"></div>
    <div class="dd-object-header-body">
      <h2 class="dd-object-title">SFDC Flows</h2>
      <div class="dd-object-meta">
        <strong>{len(flows)}</strong> active flows &nbsp;·&nbsp; {H(process_summary)}
      </div>
    </div>
  </header>
  <div class="dd-flows-controls">
    <input id="dd-flow-search" class="dd-search dd-flow-search" type="search" placeholder="Filter flows by label, developer name, object, or description…" autocomplete="off" />
    <div class="dd-filters">
      <div class="dd-filter-group" role="group" aria-label="Similar flows filter">
        <span class="dd-filter-label">Similar:</span>
        <button type="button" class="dd-seg-btn is-active" data-flow-similar-filter="all">All</button>
        <button type="button" class="dd-seg-btn" data-flow-similar-filter="similar">⚠ Has similar</button>
        <button type="button" class="dd-seg-btn" data-flow-similar-filter="unique">Unique</button>
      </div>
      <div class="dd-filter-group">
        <span class="dd-filter-label">Created in last:</span>
        <input id="dd-flow-created-days" class="dd-days-input" type="number" min="0" placeholder="N" inputmode="numeric" />
        <span class="dd-filter-label">days</span>
        <button type="button" id="dd-flow-created-clear" class="dd-link-btn">clear</button>
      </div>
    </div>
    <span id="dd-flow-search-meta" class="dd-search-meta">Type to filter the flow list.</span>
  </div>
  <div class="dd-table-wrap">
    <table class="dd-table dd-flows-table">
      <thead>
        <tr>
          <th class="c-flow-label">Flow Label</th>
          <th class="c-flow-dev">Developer Name</th>
          <th class="c-flow-obj">Object</th>
          <th class="c-flow-type">Process Type</th>
          <th class="c-flow-version">Version</th>
          <th class="c-flow-created">Created</th>
          <th class="c-flow-modified">Last Modified</th>
          <th class="c-flow-desc">Description</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
  </div>
</section>"""


def _render_flow_row(f: dict) -> str:
    flow_url = f"https://healthie.my.salesforce-setup.com/lightning/setup/Flows/page?address=%2F{f['def_id']}" if f.get("def_id") else ""
    label_inner = H(f["label"])
    if flow_url:
        label_inner = (
            f'<a href="{H(flow_url)}" target="_blank" rel="noopener" '
            f'class="dd-field-link">{H(f["label"])}</a>'
        )
    warn_icon = ""
    if f.get("similar_to"):
        bullets = "\n• ".join(f["similar_to"])
        tip = f"Similar flows in this object:\n• {bullets}"
        warn_icon = f' <span class="dd-tip dd-warn-icon" data-tip="{H(tip)}">⚠</span>'
    obj_chip = ""
    if f["inferred_object"]:
        obj_chip = (
            f'<span class="dd-obj-chip" style="--obj-color: {f["inferred_object_color"]};">'
            f'<span class="dd-obj-chip-dot"></span>{H(f["inferred_object"])}'
            "</span>"
        )
    else:
        obj_chip = '<span class="dd-muted">—</span>'
    version_text = f"v{f['version']}" if f.get("version") is not None else "—"
    api_version = f"API {f['api_version']}" if f.get("api_version") else ""
    version_combined = f"{version_text} {H(api_version)}".strip()
    created_cell = H(f["created"]) if f.get("created") else '<span class="dd-muted">—</span>'
    desc = H(f["description"]) if f["description"] else '<span class="dd-muted">—</span>'
    similar_flag = "1" if f.get("similar_to") else "0"
    blob = " ".join([
        f["label"], f["dev_name"], f["process_type"], f["inferred_object"], f["description"]
    ]).lower()
    return f"""
            <tr class="dd-flow-row" data-search="{H(blob)}" data-similar="{similar_flag}" data-created="{H(f.get('created') or '')}">
              <td class="c-flow-label">{label_inner}{warn_icon}</td>
              <td class="c-flow-dev"><code>{H(f['dev_name'])}</code></td>
              <td class="c-flow-obj">{obj_chip}</td>
              <td class="c-flow-type">{H(f['process_type'])}</td>
              <td class="c-flow-version">{version_combined}</td>
              <td class="c-flow-created">{created_cell}</td>
              <td class="c-flow-modified">{H(f['modified'])}</td>
              <td class="c-flow-desc">{desc}</td>
            </tr>"""


def _render_front_matter(fm: dict) -> str:
    blocks = []
    if fm.get("how_to_use_paragraphs"):
        paras = "".join(f"<p>{H(p)}</p>" for p in fm["how_to_use_paragraphs"])
        blocks.append(_accordion("Methodology", paras))
    if fm.get("column_definitions"):
        blocks.append(_accordion(
            "Column Definitions",
            _render_two_col_table(["Column", "What it captures"], fm["column_definitions"]),
        ))
    if fm.get("governance"):
        blocks.append(_accordion(
            "Governance",
            _render_two_col_table(["Area", "Owner / Process"], fm["governance"]),
        ))
    if fm.get("data_sources"):
        # data_sources rows may be 3 wide
        headers = ["Source System", "Writes to", "Notes"]
        blocks.append(_accordion(
            "Data Sources & Integrations",
            _render_two_col_table(headers, fm["data_sources"]),
        ))
    if fm.get("glossary"):
        blocks.append(_accordion(
            "Glossary",
            _render_two_col_table(["Term", "Definition"], fm["glossary"]),
        ))
    return "\n".join(blocks)


def _render_appendix(ap: dict) -> str:
    parts = []
    if ap.get("picklist_values"):
        inner = ""
        for pl in ap["picklist_values"]:
            inner += f'<h4 class="dd-appendix-subtitle">{H(pl["title"])}</h4>'
            inner += _render_two_col_table(pl["headers"], pl["rows"])
        parts.append(_accordion("Picklist Values", inner))
    if ap.get("v4_additions"):
        parts.append(_accordion(
            "V4 Field Additions",
            _render_two_col_table(ap["v4_additions"]["headers"], ap["v4_additions"]["rows"]),
        ))
    if ap.get("changelog"):
        parts.append(_accordion(
            "Changelog",
            _render_two_col_table(ap["changelog"]["headers"], ap["changelog"]["rows"]),
        ))
    return "\n".join(parts)


def _accordion(title: str, inner_html: str) -> str:
    return f"""
<details class="dd-accordion">
  <summary class="dd-accordion-title">{H(title)}</summary>
  <div class="dd-accordion-body">{inner_html}</div>
</details>"""


# ---------- inline assets ----------

_CSS = """
.dd-root {
  --indigo: #4f46e5;
  --indigo-deep: #312e81;
  --slate-900: #0f172a;
  --slate-800: #1e293b;
  --slate-700: #334155;
  --slate-600: #475569;
  --slate-500: #64748b;
  --slate-400: #94a3b8;
  --slate-300: #cbd5e1;
  --slate-200: #e2e8f0;
  --slate-100: #f1f5f9;
  --slate-50: #f8fafc;
  --green: #16a34a;
  --green-soft: #dcfce7;
  --amber: #d97706;
  --amber-soft: #ffedd5;
  --red: #dc2626;
  --red-soft: #fee2e2;
  --pii: #db2777;
  --pii-soft: #fce7f3;
  --fin: #0e7490;
  --fin-soft: #cffafe;
  --finid: #6d28d9;
  --finid-soft: #ede9fe;
  --internal: #475569;
  --internal-soft: #e2e8f0;
  font-family: 'Inter', system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  color: var(--slate-800);
  font-size: 13px;
  line-height: 1.55;
  background: white;
}
.dd-root * { box-sizing: border-box; }
.dd-root a { color: var(--indigo); text-decoration: none; }
.dd-root code { font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace; font-size: 12px; background: var(--slate-100); border-radius: 4px; padding: 1px 5px; color: var(--slate-800); }
.dd-root h1, .dd-root h2, .dd-root h3, .dd-root h4 { margin: 0; color: var(--slate-900); font-weight: 600; }

.dd-navbar { position: sticky; top: 0; z-index: 100; background: white; border-bottom: 1px solid var(--slate-200); display: flex; align-items: center; gap: 12px; padding: 12px 24px; font-size: 13px; flex-wrap: wrap; }
.dd-nav-brand { font-size: 14px; font-weight: 700; color: var(--indigo); letter-spacing: -0.005em; }
.dd-nav-breadcrumb { color: var(--slate-500); font-size: 13px; }
.dd-nav-right { margin-left: auto; display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
.dd-nav-link { font-size: 12px; font-weight: 600; text-decoration: none; padding: 4px 10px; border-radius: 6px; border: 1px solid; white-space: nowrap; transition: background-color 120ms; }
.dd-nav-link-purple { color: var(--indigo); border-color: #ddd6fe; }
.dd-nav-link-purple:hover { background: #f5f3ff; }
.dd-nav-link-green { color: #10b981; border-color: #a7f3d0; }
.dd-nav-link-green:hover { background: #ecfdf5; }
.dd-nav-link-amber { color: #d97706; border-color: #fde68a; }
.dd-nav-link-amber:hover { background: #fffbeb; }
.dd-nav-link-slate-2 { color: var(--slate-700); border-color: var(--slate-300); }
.dd-nav-link-slate-2:hover { background: var(--slate-50); }
.dd-nav-link-slate { color: var(--slate-700); border-color: var(--slate-300); background: white; font: inherit; font-size: 12px; font-weight: 600; padding: 4px 10px; border-style: solid; border-width: 1px; border-radius: 6px; cursor: pointer; }
.dd-nav-link-slate:hover { background: var(--slate-50); color: var(--slate-900); }
.dd-export-menu { position: relative; }
.dd-export-pop { position: absolute; top: calc(100% + 6px); right: 0; background: white; border: 1px solid var(--slate-200); border-radius: 10px; box-shadow: 0 12px 28px rgba(15, 23, 42, 0.18); padding: 6px; min-width: 280px; z-index: 200; }
.dd-export-pop[hidden] { display: none; }
.dd-export-item { display: block; width: 100%; text-align: left; background: white; border: 0; padding: 9px 12px; border-radius: 6px; cursor: pointer; font: inherit; color: var(--slate-800); }
.dd-export-item:hover { background: var(--slate-50); }
.dd-export-title { display: block; font-weight: 600; font-size: 13px; color: var(--slate-900); }
.dd-export-sub { display: block; font-size: 11.5px; color: var(--slate-500); margin-top: 1px; }
.dd-export-toast { position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%); background: #0f172a; color: white; padding: 10px 18px; border-radius: 8px; font-size: 13px; box-shadow: 0 8px 20px rgba(15, 23, 42, 0.25); z-index: 9999; opacity: 0; transition: opacity 160ms; pointer-events: none; }
.dd-export-toast.is-visible { opacity: 1; }
.dd-nav-stats { font-size: 11px; color: var(--slate-400); }

.dd-hero { background: linear-gradient(180deg, #f5f3ff 0%, #ffffff 100%); border-bottom: 1px solid var(--slate-200); }
.dd-hero-inner { max-width: 1240px; margin: 0 auto; padding: 36px 20px 28px; }
.dd-hero-title { font-size: 28px; font-weight: 700; letter-spacing: -0.01em; margin-bottom: 6px; }
.dd-hero-sub { color: var(--slate-600); margin: 0 0 20px; max-width: 720px; }
.dd-kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; }
.dd-kpi { background: white; border: 1px solid var(--slate-200); border-radius: 10px; padding: 14px 16px; box-shadow: 0 1px 2px rgba(15,23,42,0.04); }
.dd-kpi-num { font-size: 22px; font-weight: 700; color: var(--slate-900); }
.dd-kpi-label { color: var(--slate-500); font-size: 12px; margin-top: 2px; }

.dd-controls { background: white; position: sticky; top: 42px; z-index: 90; border-bottom: 1px solid var(--slate-200); }
.dd-controls-inner { max-width: 1240px; margin: 0 auto; padding: 14px 20px; }
.dd-search { width: 100%; padding: 11px 14px; font-size: 14px; border: 1px solid var(--slate-300); border-radius: 8px; outline: none; transition: border-color 120ms, box-shadow 120ms; }
.dd-search:focus { border-color: var(--indigo); box-shadow: 0 0 0 3px rgba(79,70,229,0.18); }
.dd-search-meta { margin-top: 6px; color: var(--slate-500); font-size: 12px; }
.dd-filters { display: flex; gap: 18px; flex-wrap: wrap; align-items: center; margin-top: 10px; }
.dd-filter-group { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; }
.dd-filter-label { color: var(--slate-600); font-weight: 500; }
.dd-seg-btn { background: white; border: 1px solid var(--slate-300); color: var(--slate-600); font: inherit; font-size: 12px; padding: 4px 10px; cursor: pointer; transition: background-color 120ms, border-color 120ms, color 120ms; }
.dd-seg-btn:first-of-type { border-radius: 6px 0 0 6px; }
.dd-seg-btn:last-of-type { border-radius: 0 6px 6px 0; }
.dd-seg-btn + .dd-seg-btn { border-left: 0; }
.dd-seg-btn:hover { background: var(--slate-50); color: var(--slate-900); }
.dd-seg-btn.is-active { background: var(--indigo); color: white; border-color: var(--indigo); }
.dd-days-input { width: 64px; padding: 4px 8px; border: 1px solid var(--slate-300); border-radius: 6px; font-size: 12px; font-family: inherit; font-variant-numeric: tabular-nums; }
.dd-days-input:focus { outline: none; border-color: var(--indigo); box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.18); }
.dd-link-btn { background: none; border: none; color: var(--indigo); font: inherit; font-size: 11px; padding: 0; cursor: pointer; text-decoration: underline; }
.dd-link-btn:hover { color: var(--indigo-deep); }

.dd-legend { background: var(--slate-50); border-bottom: 1px solid var(--slate-200); }
.dd-legend-inner { max-width: 1240px; margin: 0 auto; padding: 10px 20px; display: flex; gap: 16px; flex-wrap: wrap; align-items: center; font-size: 12px; color: var(--slate-600); }
.dd-legend-item { display: inline-flex; align-items: center; gap: 6px; }
.dd-legend-chip { display: inline-flex; align-items: center; justify-content: center; min-width: 22px; height: 22px; padding: 0 8px; border-radius: 6px; font-weight: 700; font-size: 12px; }
.dd-legend-chip-key { background: #fef9c3; color: #b45309; border: 1px solid #fde68a; }
.dd-legend-chip-new { background: #dcfce7; color: #166534; border: 1px solid #86efac; font-size: 10.5px; letter-spacing: 0.04em; }
.dd-legend-chip-warn { background: #fff7ed; color: #c2410c; border: 1px solid #fdba74; }
.dd-new-chip { display: inline-block; background: #dcfce7; color: #166534; border: 1px solid #86efac; padding: 1px 6px; border-radius: 4px; font-size: 9.5px; font-weight: 700; letter-spacing: 0.05em; margin-left: 6px; vertical-align: middle; cursor: help; }

.dd-tabs { position: sticky; top: 196px; z-index: 80; background: white; border-bottom: 1px solid var(--slate-200); overflow-x: auto; }
.dd-tabs { display: flex; gap: 0; max-width: 1240px; margin: 0 auto; padding: 0 20px; }
.dd-tab { background: transparent; border: 0; border-bottom: 2px solid transparent; padding: 12px 14px; font: inherit; cursor: pointer; color: var(--slate-600); display: inline-flex; align-items: center; gap: 8px; white-space: nowrap; }
.dd-tab:hover { color: var(--slate-900); }
.dd-tab-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--slate-300); display: inline-block; }
.dd-tab.is-active { color: var(--slate-900); border-bottom-color: var(--obj-color, var(--indigo)); font-weight: 600; }
.dd-tab-count { background: var(--slate-100); color: var(--slate-700); border-radius: 999px; padding: 1px 8px; font-size: 11px; font-weight: 500; }
.dd-tab.is-active .dd-tab-count { background: var(--obj-color, var(--indigo-deep)); color: white; }
.dd-tab.is-empty { opacity: 0.45; }

.dd-panels { max-width: 1240px; margin: 0 auto; padding: 18px 20px 40px; }
.dd-panel { display: none; }
.dd-panel.is-active { display: block; }

/* Object header: colored top band + colored title so each object reads as a
   distinct section. The band uses the object's accent color (--obj-color). */
.dd-object-header { margin-bottom: 18px; border: 1px solid var(--slate-200); border-radius: 10px; background: white; overflow: hidden; }
.dd-object-band { height: 8px; background: var(--obj-color, var(--indigo)); }
.dd-object-header-body { padding: 14px 18px 16px; }
.dd-root h2.dd-object-title { font-size: 20px; font-weight: 700; color: var(--obj-color, var(--indigo)); letter-spacing: -0.01em; margin-bottom: 4px; }
.dd-object-meta { color: var(--slate-600); font-size: 13px; }
.dd-subsection { margin: 22px 0; }
.dd-subsection-title { font-size: 15px; margin-bottom: 8px; padding-left: 10px; border-left: 4px solid var(--obj-color, var(--indigo)); }
.dd-muted { color: var(--slate-400); font-weight: normal; }

.dd-key-star { color: #ca8a04; margin-left: 4px; font-size: 12px; }
.dd-warn-icon { color: #c2410c; margin-left: 6px; font-size: 12px; }
.dd-field-link { color: var(--slate-900); text-decoration: none; border-bottom: 1px dashed transparent; transition: border-color 120ms, color 120ms; }
.dd-field-link:hover { color: var(--indigo); border-bottom-color: var(--indigo); }
.dd-field-link::after { content: ""; display: inline-block; width: 10px; height: 10px; margin-left: 4px; background: currentColor; opacity: 0; transition: opacity 120ms; -webkit-mask: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M7 17 17 7"/><path d="M7 7h10v10"/></svg>') center/contain no-repeat; mask: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M7 17 17 7"/><path d="M7 7h10v10"/></svg>') center/contain no-repeat; vertical-align: -1px; }
.dd-field-link:hover::after { opacity: 0.6; }
.dd-tip { cursor: help; display: inline-block; }
.dd-tooltip-pop {
  position: fixed;
  background: #0f172a;
  color: #f8fafc;
  padding: 8px 12px;
  border-radius: 6px;
  font-size: 12px;
  font-weight: 400;
  line-height: 1.5;
  white-space: pre-line;
  max-width: 360px;
  z-index: 9999;
  box-shadow: 0 6px 18px rgba(15, 23, 42, 0.22);
  pointer-events: none;
}
.dd-row.is-key td { background: #fef9c3; }
.dd-row.is-key:hover td { background: #fef3c7; }
.dd-row.is-key.is-match td { background: #fde68a; }
.dd-row.is-key td.c-label { background: #fef9c3; }
.dd-table tr.is-key:hover td.c-label { background: #fef3c7; }

.dd-table-wrap { overflow-x: auto; border: 1px solid var(--slate-200); border-radius: 10px; background: white; position: relative; scroll-behavior: smooth; }
.dd-table-wrap::-webkit-scrollbar { height: 12px; background: var(--slate-50); }
.dd-table-wrap::-webkit-scrollbar-track { background: var(--slate-100); border-radius: 6px; margin: 0 4px; }
.dd-table-wrap::-webkit-scrollbar-thumb { background: var(--slate-400); border-radius: 6px; border: 2px solid var(--slate-100); min-width: 60px; }
.dd-table-wrap::-webkit-scrollbar-thumb:hover { background: var(--slate-500); }
.dd-table-wrap { scrollbar-color: var(--slate-400) var(--slate-100); scrollbar-width: thin; }
/* Edge fade shadows — inset box-shadow stays at the visible edge of the scroll
   viewport. JS toggles .has-scroll-left / .has-scroll-right based on position. */
.dd-table-wrap { transition: box-shadow 140ms ease-out; }
.dd-table-wrap.has-scroll-right { box-shadow: inset -22px 0 16px -14px rgba(15, 23, 42, 0.25); }
.dd-table-wrap.has-scroll-left { box-shadow: inset 22px 0 16px -14px rgba(15, 23, 42, 0.25); }
.dd-table-wrap.has-scroll-left.has-scroll-right { box-shadow: inset 22px 0 16px -14px rgba(15, 23, 42, 0.25), inset -22px 0 16px -14px rgba(15, 23, 42, 0.25); }
.dd-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 12.5px; }
.dd-table thead th { position: sticky; top: 0; background: var(--slate-50); color: var(--slate-700); font-weight: 600; text-align: left; padding: 10px 12px; border-bottom: 2px solid var(--slate-200); white-space: nowrap; z-index: 1; }
.dd-table tbody td { padding: 11px 12px; border-bottom: 1px solid var(--slate-100); vertical-align: top; color: var(--slate-800); word-break: break-word; overflow-wrap: anywhere; }
.dd-table tbody tr:hover td { background: #f5f3ff; }
.dd-row.is-hidden { display: none; }
.dd-row.is-match td { background: #fffbeb; }
/* Sticky Field Label column so users always see which field they're on while
   scrolling horizontally to reach the right-side columns. */
.dd-table th.c-label, .dd-table td.c-label { position: sticky; left: 0; background: white; z-index: 2; box-shadow: 1px 0 0 0 var(--slate-200); }
.dd-table thead th.c-label { background: var(--slate-50); z-index: 3; }
.dd-table tbody tr:hover td.c-label { background: #f5f3ff; }
.dd-table tr.is-match td.c-label { background: #fffbeb; }
.dd-table .c-label { min-width: 180px; max-width: 240px; font-weight: 500; }
.dd-table .c-api { min-width: 200px; max-width: 260px; }
.dd-table .c-type { min-width: 130px; max-width: 170px; }
.dd-table .c-fill { min-width: 70px; max-width: 90px; }
.dd-table .c-desc { min-width: 360px; max-width: 460px; }
.dd-table .c-values { min-width: 260px; max-width: 380px; color: var(--slate-700); }
.dd-table .c-source { min-width: 140px; max-width: 180px; }
.dd-table .c-created { min-width: 110px; max-width: 130px; white-space: nowrap; color: var(--slate-600); font-variant-numeric: tabular-nums; }
.dd-table .c-owner { min-width: 110px; max-width: 150px; }
.dd-table .c-sens { min-width: 100px; max-width: 130px; }
.dd-table .c-deps { min-width: 180px; max-width: 240px; }
.dd-values { color: var(--slate-700); }

.dd-chip { display: inline-block; padding: 2px 9px; border-radius: 999px; font-weight: 600; font-size: 11.5px; }
.dd-chip-high { background: var(--green-soft); color: var(--green); }
.dd-chip-mid { background: var(--amber-soft); color: var(--amber); }
.dd-chip-low { background: var(--red-soft); color: var(--red); }
.dd-chip-na { background: var(--slate-100); color: var(--slate-500); }

.dd-pill { display: inline-block; padding: 2px 9px; border-radius: 6px; font-size: 11.5px; font-weight: 500; }
.dd-pill-pii { background: var(--pii-soft); color: var(--pii); }
.dd-pill-fin { background: var(--fin-soft); color: var(--fin); }
.dd-pill-fin-id { background: var(--finid-soft); color: var(--finid); }
.dd-pill-internal { background: var(--internal-soft); color: var(--internal); }
.dd-pill-other { background: var(--slate-100); color: var(--slate-600); }

.dd-values { color: var(--slate-700); }
.dd-section-title { font-size: 18px; margin: 36px 0 12px; max-width: 1240px; margin-left: auto; margin-right: auto; padding: 0 20px; }
.dd-accordion-group { max-width: 1240px; margin: 0 auto; padding: 0 20px; }
.dd-accordion { background: white; border: 1px solid var(--slate-200); border-radius: 10px; margin: 10px 0; overflow: hidden; }
.dd-accordion[open] .dd-accordion-title::after { transform: rotate(90deg); }
.dd-accordion-title { cursor: pointer; padding: 12px 16px; font-weight: 600; color: var(--slate-900); display: flex; align-items: center; gap: 10px; list-style: none; }
.dd-accordion-title::-webkit-details-marker { display: none; }
.dd-accordion-title::after { content: "›"; font-size: 18px; color: var(--slate-400); transition: transform 120ms; margin-left: auto; }
.dd-accordion-body { padding: 6px 16px 16px; border-top: 1px solid var(--slate-100); }
.dd-appendix-subtitle { font-size: 13px; margin: 12px 0 6px; color: var(--slate-700); }
.dd-table-narrow th, .dd-table-narrow td { padding: 8px 10px; font-size: 12.5px; }

.dd-footer { max-width: 1240px; margin: 30px auto 60px; padding: 0 20px; color: var(--slate-500); font-size: 12px; }

.dd-others { max-width: 1240px; margin: 24px auto 0; padding: 0 20px; }
.dd-other-tabs { position: static; margin-top: 14px; }
.dd-other-panels { padding: 0; }
.dd-other-panel { display: none; }
.dd-other-panel.is-active { display: block; }
.dd-other-row.is-hidden { display: none; }
/* Sticky first column (label) — addressed by position to keep the <td>s class-free. */
.dd-others-table thead th.c-other-label, .dd-others-table tbody td:nth-child(1) { position: sticky; left: 0; background: white; box-shadow: 1px 0 0 0 var(--slate-200); }
.dd-others-table thead th.c-other-label { background: var(--slate-50); }
.dd-others-table tbody tr:hover td:nth-child(1) { background: #f5f3ff; }
.dd-others-table th.c-other-label, .dd-others-table tbody td:nth-child(1) { min-width: 220px; max-width: 320px; font-weight: 500; }
.dd-others-table th.c-other-api, .dd-others-table tbody td:nth-child(2) { min-width: 220px; max-width: 320px; }
.dd-others-table th.c-other-type, .dd-others-table tbody td:nth-child(3) { min-width: 130px; max-width: 180px; }
.dd-others-table th.c-other-flag, .dd-others-table tbody td:nth-child(4) { min-width: 90px; max-width: 110px; }
.dd-others-table th.c-other-created, .dd-others-table tbody td:nth-child(5) { min-width: 110px; max-width: 130px; white-space: nowrap; color: var(--slate-600); font-variant-numeric: tabular-nums; }
.dd-others-table th.c-other-desc, .dd-others-table tbody td:nth-child(6) { min-width: 360px; max-width: 600px; color: var(--slate-700); }
/* Anchors inside All Other Fields rows inherit dd-field-link styling without the per-anchor class. */
.dd-other-row a { color: var(--indigo); text-decoration: none; border-bottom: 1px dashed transparent; }
.dd-other-row a:hover { border-bottom-color: var(--indigo); }
.dd-pill-custom { background: #ede9fe; color: #6d28d9; }
.dd-pill-standard { background: var(--slate-100); color: var(--slate-600); }
/* Short pill aliases used only by the All Other Fields rows (1,290 occurrences). Match .dd-pill shape. */
.p { display: inline-block; padding: 2px 9px; border-radius: 6px; font-size: 11.5px; font-weight: 500; }
.p-c { background: #ede9fe; color: #6d28d9; }
.p-s { background: var(--slate-100); color: var(--slate-600); }

.dd-flows { max-width: 1240px; margin: 24px auto 0; padding: 0 20px; }
.dd-flows-controls { display: flex; gap: 10px; align-items: center; margin: 14px 0; }
.dd-flow-search { flex: 1; }
.dd-flow-row.is-hidden { display: none; }
.dd-flow-row.is-match td { background: #fffbeb; }
.dd-flows-table thead th.c-flow-label, .dd-flows-table tbody td.c-flow-label { position: sticky; left: 0; background: white; box-shadow: 1px 0 0 0 var(--slate-200); }
.dd-flows-table thead th.c-flow-label { background: var(--slate-50); }
.dd-flows-table tbody tr:hover td.c-flow-label { background: #f5f3ff; }
.dd-flows-table .c-flow-label { min-width: 240px; max-width: 320px; font-weight: 500; }
.dd-flows-table .c-flow-dev { min-width: 240px; max-width: 320px; }
.dd-flows-table .c-flow-obj { min-width: 130px; max-width: 170px; }
.dd-flows-table .c-flow-type { min-width: 140px; max-width: 180px; }
.dd-flows-table .c-flow-version { min-width: 90px; max-width: 120px; white-space: nowrap; }
.dd-flows-table .c-flow-created { min-width: 100px; max-width: 120px; white-space: nowrap; color: var(--slate-600); font-variant-numeric: tabular-nums; }
.dd-flows-table .c-flow-modified { min-width: 100px; max-width: 120px; white-space: nowrap; color: var(--slate-500); font-variant-numeric: tabular-nums; }
.dd-flows-table .c-flow-desc { min-width: 320px; max-width: 480px; color: var(--slate-700); }
.dd-obj-chip { display: inline-flex; align-items: center; gap: 6px; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 600; color: var(--obj-color); background: color-mix(in srgb, var(--obj-color) 12%, white); border: 1px solid color-mix(in srgb, var(--obj-color) 30%, white); }
.dd-obj-chip-dot { width: 6px; height: 6px; border-radius: 50%; background: var(--obj-color); display: inline-block; }

@media (max-width: 700px) {
  .dd-navbar { padding: 10px 14px; }
  .dd-nav-stats { display: none; }
  .dd-controls { top: 70px; }
  .dd-legend-inner { padding: 8px 14px; gap: 10px; }
  .dd-tabs { top: 188px; }
  .dd-controls-inner { padding: 12px; }
  .dd-hero-inner { padding: 24px 14px 18px; }
  .dd-hero-title { font-size: 22px; }
}
"""

_JS = r"""
(function(){
  var root = document.getElementById('dd-root');
  if (!root) return;
  // Scoped to the main object tab nav only — the All Other Fields section
  // has its own .dd-other-tab + .dd-other-panel classes handled separately.
  var tabs = Array.prototype.slice.call(root.querySelectorAll('#dd-tabs .dd-tab'));
  var panels = Array.prototype.slice.call(root.querySelectorAll('#dd-panels > .dd-panel'));
  var search = root.querySelector('#dd-search');
  var meta = root.querySelector('#dd-search-meta');

  function activateTab(id){
    tabs.forEach(function(t){
      var on = t.getAttribute('data-tab') === id;
      t.classList.toggle('is-active', on);
    });
    panels.forEach(function(p){
      p.classList.toggle('is-active', p.id === id);
    });
  }

  // Initialize first tab.
  if (tabs.length){ activateTab(tabs[0].getAttribute('data-tab')); }

  tabs.forEach(function(t){
    t.addEventListener('click', function(){
      activateTab(t.getAttribute('data-tab'));
    });
  });

  // Filter state — composed of three independent constraints, AND'd together.
  var filterState = { q: '', similar: 'all', daysWindow: null };

  function applyFilter(){
    var q = (filterState.q || '').trim().toLowerCase();
    var hasQuery = q.length > 0;
    var simMode = filterState.similar;
    var hasSimFilter = simMode !== 'all';
    var days = filterState.daysWindow;
    var hasDayFilter = typeof days === 'number' && days > 0;
    var cutoffMs = hasDayFilter ? (Date.now() - days * 86400000) : null;
    var anyExtraFilter = hasQuery || hasSimFilter || hasDayFilter;
    var totalMatches = 0;
    var perTabCounts = {};

    panels.forEach(function(p){
      var tabId = p.id;
      var rows = Array.prototype.slice.call(p.querySelectorAll('tr.dd-row'));
      var visibleCount = 0;
      rows.forEach(function(r){
        var hit = true;
        if (hasQuery){
          var blob = r.getAttribute('data-search') || '';
          if (blob.indexOf(q) === -1) hit = false;
        }
        if (hit && hasSimFilter){
          var hasSim = r.getAttribute('data-similar') === '1';
          if (simMode === 'similar' && !hasSim) hit = false;
          if (simMode === 'unique' && hasSim) hit = false;
        }
        if (hit && hasDayFilter){
          var created = r.getAttribute('data-created') || '';
          if (!created) { hit = false; }
          else {
            // Compare as midnight UTC; tolerant of YYYY-MM-DD format.
            var t = Date.parse(created + 'T00:00:00Z');
            if (isNaN(t) || t < cutoffMs) hit = false;
          }
        }
        r.classList.toggle('is-hidden', !hit);
        r.classList.toggle('is-match', hit && hasQuery);
        if (hit) visibleCount++;
      });
      perTabCounts[tabId] = visibleCount;
      totalMatches += visibleCount;
    });

    tabs.forEach(function(t){
      var id = t.getAttribute('data-tab');
      var c = perTabCounts[id] || 0;
      var count = t.querySelector('[data-count-for="' + id + '"]');
      if (count) count.textContent = c;
      t.classList.toggle('is-empty', anyExtraFilter && c === 0);
    });

    if (meta){
      if (anyExtraFilter){
        var nObjs = Object.keys(perTabCounts).filter(function(k){ return perTabCounts[k] > 0; }).length;
        var bits = [totalMatches + ' field' + (totalMatches === 1 ? '' : 's') + ' across ' + nObjs + ' object' + (nObjs === 1 ? '' : 's')];
        if (hasQuery) bits.push('matching "' + q + '"');
        if (hasSimFilter) bits.push(simMode === 'similar' ? 'with similar siblings' : 'unique within object');
        if (hasDayFilter) bits.push('created in last ' + days + ' days');
        // When a global query is active it also filters the All Other Fields
        // section (handled by the search input handler below). Surface its
        // count here so the user knows there are uncurated matches to scroll to.
        if (hasQuery){
          var otherMatches = root.querySelectorAll('#dd-other-fields tr.dd-other-row:not(.is-hidden)').length;
          if (otherMatches > 0) bits.push(otherMatches + ' in All Other Fields');
        }
        meta.textContent = bits.join(' · ') + '.';
      } else {
        meta.textContent = 'Type to filter. Tab badges update live.';
      }
    }
  }

  if (search){
    search.addEventListener('input', function(){
      filterState.q = search.value;
      // Also drive the All Other Fields section so global search surfaces
      // matches there too. Visually sync the section's local search input
      // so the user sees it's filtered.
      otherFilterState.q = search.value;
      var otherSearchInput = root.querySelector('#dd-other-search');
      if (otherSearchInput) otherSearchInput.value = search.value;
      applyOtherFilter();
      applyFilter();
    });
  }

  // Similar-fields segmented control
  Array.prototype.slice.call(root.querySelectorAll('[data-similar-filter]')).forEach(function(btn){
    btn.addEventListener('click', function(){
      filterState.similar = btn.getAttribute('data-similar-filter');
      root.querySelectorAll('[data-similar-filter]').forEach(function(b){
        b.classList.toggle('is-active', b === btn);
      });
      applyFilter();
    });
  });

  // Created-within-N-days filter
  var daysInput = root.querySelector('#dd-created-days');
  var daysClear = root.querySelector('#dd-created-clear');
  if (daysInput){
    daysInput.addEventListener('input', function(){
      var v = parseInt(daysInput.value, 10);
      filterState.daysWindow = (isNaN(v) || v <= 0) ? null : v;
      applyFilter();
    });
  }
  if (daysClear){
    daysClear.addEventListener('click', function(){
      daysInput.value = '';
      filterState.daysWindow = null;
      applyFilter();
    });
  }

  // ---- Flows section filters (search + similar + days window) ----
  var flowSearch = root.querySelector('#dd-flow-search');
  var flowMeta = root.querySelector('#dd-flow-search-meta');
  var flowDaysInput = root.querySelector('#dd-flow-created-days');
  var flowDaysClear = root.querySelector('#dd-flow-created-clear');
  var flowFilterState = { q: '', similar: 'all', daysWindow: null };

  function applyFlowFilter(){
    var q = (flowFilterState.q || '').trim().toLowerCase();
    var hasQ = q.length > 0;
    var simMode = flowFilterState.similar;
    var hasSim = simMode !== 'all';
    var days = flowFilterState.daysWindow;
    var hasDays = typeof days === 'number' && days > 0;
    var cutoffMs = hasDays ? (Date.now() - days * 86400000) : null;
    var rows = Array.prototype.slice.call(root.querySelectorAll('tr.dd-flow-row'));
    var visible = 0;
    rows.forEach(function(r){
      var hit = true;
      if (hasQ){
        var blob = r.getAttribute('data-search') || '';
        if (blob.indexOf(q) === -1) hit = false;
      }
      if (hit && hasSim){
        var sim = r.getAttribute('data-similar') === '1';
        if (simMode === 'similar' && !sim) hit = false;
        if (simMode === 'unique' && sim) hit = false;
      }
      if (hit && hasDays){
        var created = r.getAttribute('data-created') || '';
        if (!created) { hit = false; }
        else {
          var t = Date.parse(created + 'T00:00:00Z');
          if (isNaN(t) || t < cutoffMs) hit = false;
        }
      }
      r.classList.toggle('is-hidden', !hit);
      r.classList.toggle('is-match', hit && hasQ);
      if (hit) visible++;
    });
    if (flowMeta){
      if (hasQ || hasSim || hasDays){
        var bits = [visible + ' flow' + (visible === 1 ? '' : 's')];
        if (hasQ) bits.push('matching "' + q + '"');
        if (hasSim) bits.push(simMode === 'similar' ? 'with similar siblings' : 'unique within object');
        if (hasDays) bits.push('created in last ' + days + ' days');
        flowMeta.textContent = bits.join(' · ') + '.';
      } else {
        flowMeta.textContent = 'Type to filter the flow list.';
      }
    }
  }

  if (flowSearch){
    flowSearch.addEventListener('input', function(){
      flowFilterState.q = flowSearch.value;
      applyFlowFilter();
    });
  }
  Array.prototype.slice.call(root.querySelectorAll('[data-flow-similar-filter]')).forEach(function(btn){
    btn.addEventListener('click', function(){
      flowFilterState.similar = btn.getAttribute('data-flow-similar-filter');
      root.querySelectorAll('[data-flow-similar-filter]').forEach(function(b){
        b.classList.toggle('is-active', b === btn);
      });
      applyFlowFilter();
    });
  });
  if (flowDaysInput){
    flowDaysInput.addEventListener('input', function(){
      var v = parseInt(flowDaysInput.value, 10);
      flowFilterState.daysWindow = (isNaN(v) || v <= 0) ? null : v;
      applyFlowFilter();
    });
  }
  if (flowDaysClear){
    flowDaysClear.addEventListener('click', function(){
      flowDaysInput.value = '';
      flowFilterState.daysWindow = null;
      applyFlowFilter();
    });
  }

  // ---- Tooltips: hover anything with .dd-tip and [data-tip] ----
  var tipEl = null;
  function showTip(target){
    var text = target.getAttribute('data-tip');
    if (!text) return;
    if (!tipEl){
      tipEl = document.createElement('div');
      tipEl.className = 'dd-tooltip-pop';
      document.body.appendChild(tipEl);
    }
    tipEl.textContent = text;
    var r = target.getBoundingClientRect();
    // Position above the icon, clamp inside the viewport.
    var tipRect;
    tipEl.style.left = '-9999px';
    tipEl.style.top = '0px';
    tipEl.style.display = 'block';
    tipRect = tipEl.getBoundingClientRect();
    var left = r.left;
    if (left + tipRect.width > window.innerWidth - 8) {
      left = Math.max(8, window.innerWidth - tipRect.width - 8);
    }
    var top = r.top - tipRect.height - 8;
    if (top < 8) {
      // Flip below the icon if there's no room above.
      top = r.bottom + 8;
    }
    tipEl.style.left = left + 'px';
    tipEl.style.top = top + 'px';
  }
  function hideTip(){ if (tipEl) tipEl.style.display = 'none'; }
  root.addEventListener('mouseover', function(e){
    var t = e.target.closest && e.target.closest('.dd-tip');
    if (t) showTip(t);
  });
  root.addEventListener('mouseout', function(e){
    var t = e.target.closest && e.target.closest('.dd-tip');
    if (t) hideTip();
  });
  document.addEventListener('scroll', hideTip, true);

  // ---- Horizontal-scroll UX ----
  // (a) Edge shadow indicators — show inset box-shadow on whichever side has
  //     more content to reveal. Update on scroll, resize, and tab change.
  // (b) Mouse-wheel converts vertical motion to horizontal motion inside any
  //     table wrap that has overflow. Lets the user scroll wide tables
  //     without shift-wheel or hunting for the scrollbar.
  function updateScrollEdges(wrap){
    if (!wrap) return;
    var hasRight = wrap.scrollLeft + wrap.clientWidth < wrap.scrollWidth - 1;
    var hasLeft = wrap.scrollLeft > 0;
    wrap.classList.toggle('has-scroll-right', hasRight);
    wrap.classList.toggle('has-scroll-left', hasLeft);
  }
  function refreshAllScrollEdges(){
    root.querySelectorAll('.dd-table-wrap').forEach(updateScrollEdges);
  }
  root.querySelectorAll('.dd-table-wrap').forEach(function(wrap){
    updateScrollEdges(wrap);
    wrap.addEventListener('scroll', function(){ updateScrollEdges(wrap); }, { passive: true });
    // Wheel-to-horizontal: only intercept dominant vertical motion, and only
    // when the wrap actually has horizontal overflow + room to scroll in the
    // wheeled direction. Lets natural vertical page scroll pass through when
    // the wrap is fully scrolled to the relevant edge.
    wrap.addEventListener('wheel', function(e){
      if (e.shiftKey) return;  // user already requested horizontal explicitly
      if (Math.abs(e.deltaX) > Math.abs(e.deltaY)) return;  // already horizontal
      if (wrap.scrollWidth <= wrap.clientWidth) return;
      var atLeft = wrap.scrollLeft <= 0;
      var atRight = wrap.scrollLeft + wrap.clientWidth >= wrap.scrollWidth - 1;
      if ((e.deltaY < 0 && atLeft) || (e.deltaY > 0 && atRight)) return;
      e.preventDefault();
      wrap.scrollLeft += e.deltaY;
    }, { passive: false });
  });
  window.addEventListener('resize', refreshAllScrollEdges);
  // Tabs hide/show panels, which changes which wraps are measurable.
  tabs.forEach(function(t){
    t.addEventListener('click', function(){
      // After the panel becomes visible the wraps inside need a re-measure.
      setTimeout(refreshAllScrollEdges, 0);
    });
  });

  // ---- "All Other Fields" section: tabs + filters ----
  var otherTabs = Array.prototype.slice.call(root.querySelectorAll('.dd-other-tab'));
  var otherPanels = Array.prototype.slice.call(root.querySelectorAll('.dd-other-panel'));
  function activateOtherTab(id){
    otherTabs.forEach(function(t){ t.classList.toggle('is-active', t.getAttribute('data-other-tab') === id); });
    otherPanels.forEach(function(p){ p.classList.toggle('is-active', p.id === id); });
    setTimeout(refreshAllScrollEdges, 0);
  }
  otherTabs.forEach(function(t){
    t.addEventListener('click', function(){ activateOtherTab(t.getAttribute('data-other-tab')); });
  });

  var otherFilterState = { q: '', typeClass: 'all', daysWindow: null };
  function applyOtherFilter(){
    var q = (otherFilterState.q || '').trim().toLowerCase();
    var hasQ = q.length > 0;
    var typeMode = otherFilterState.typeClass;
    var hasType = typeMode !== 'all';
    var days = otherFilterState.daysWindow;
    var hasDays = typeof days === 'number' && days > 0;
    var cutoffMs = hasDays ? (Date.now() - days * 86400000) : null;
    var perTab = {};
    otherPanels.forEach(function(p){
      var rows = Array.prototype.slice.call(p.querySelectorAll('tr.dd-other-row'));
      var visible = 0;
      rows.forEach(function(r){
        var hit = true;
        if (hasQ){
          var blob = r.getAttribute('data-search') || '';
          if (blob.indexOf(q) === -1) hit = false;
        }
        if (hit && hasType){
          var tc = r.getAttribute('data-type-class') || '';
          if (tc !== typeMode) hit = false;
        }
        if (hit && hasDays){
          var created = r.getAttribute('data-created') || '';
          if (!created) { hit = false; }
          else {
            var t = Date.parse(created + 'T00:00:00Z');
            if (isNaN(t) || t < cutoffMs) hit = false;
          }
        }
        r.classList.toggle('is-hidden', !hit);
        if (hit) visible++;
      });
      perTab[p.id] = visible;
    });
    otherTabs.forEach(function(t){
      var id = t.getAttribute('data-other-tab');
      var c = perTab[id] || 0;
      var count = t.querySelector('[data-other-count-for="' + id + '"]');
      if (count) count.textContent = c;
      t.classList.toggle('is-empty', (hasQ || hasType || hasDays) && c === 0);
    });
    var meta = root.querySelector('#dd-other-search-meta');
    if (meta){
      if (hasQ || hasType || hasDays){
        var totalVis = Object.keys(perTab).reduce(function(s, k){ return s + perTab[k]; }, 0);
        var bits = [totalVis + ' field' + (totalVis === 1 ? '' : 's')];
        if (hasQ) bits.push('matching "' + q + '"');
        if (hasType) bits.push(typeMode === 'custom' ? 'custom only' : 'standard only');
        if (hasDays) bits.push('created in last ' + days + ' days');
        meta.textContent = bits.join(' · ') + '.';
      } else {
        meta.textContent = 'Type to filter. Tab badges update live.';
      }
    }
  }
  var otherSearch = root.querySelector('#dd-other-search');
  if (otherSearch){
    otherSearch.addEventListener('input', function(){ otherFilterState.q = otherSearch.value; applyOtherFilter(); });
  }
  root.querySelectorAll('[data-other-type-filter]').forEach(function(btn){
    btn.addEventListener('click', function(){
      otherFilterState.typeClass = btn.getAttribute('data-other-type-filter');
      root.querySelectorAll('[data-other-type-filter]').forEach(function(b){ b.classList.toggle('is-active', b === btn); });
      applyOtherFilter();
    });
  });
  var otherDaysInput = root.querySelector('#dd-other-created-days');
  var otherDaysClear = root.querySelector('#dd-other-created-clear');
  if (otherDaysInput){
    otherDaysInput.addEventListener('input', function(){
      var v = parseInt(otherDaysInput.value, 10);
      otherFilterState.daysWindow = (isNaN(v) || v <= 0) ? null : v;
      applyOtherFilter();
    });
  }
  if (otherDaysClear){
    otherDaysClear.addEventListener('click', function(){
      otherDaysInput.value = '';
      otherFilterState.daysWindow = null;
      applyOtherFilter();
    });
  }

  // ---- Export: CSV / XLSX / Google Sheets TSV ----
  // The JSON island uses short keys (e.g. "fl" instead of "field_label") to
  // stay within HubSpot's published source-code size limit. We rehydrate to
  // long keys once on parse so the rest of the export code can read by name.
  // Maps must stay in sync with EXPORT_FIELD_KEY_MAP / EXPORT_FLOW_KEY_MAP
  // in generate_sfdc_dictionary_html.py.
  var FIELD_KEY_UNMAP = {o:'object',ss:'subsection',l:'field_label',a:'api_name',t:'type',f:'fill_pct',c:'created_date',n:'is_new',k:'is_key',kk:'key_kind',kt:'key_target',d:'definition',v:'active_values',su:'source_update',ow:'owner',se:'sensitivity',de:'dependencies',sf:'similar_fields'};
  var FLOW_KEY_UNMAP = {l:'flow_label',dn:'developer_name',to:'triggering_object',pt:'process_type',ver:'version',av:'api_version',c:'created_date',lm:'last_modified',d:'description',sf:'similar_flows'};
  function rehydrate(items, unmap){
    return items.map(function(item){
      var out = {};
      for (var k in item){ if (Object.prototype.hasOwnProperty.call(item, k)) out[unmap[k] || k] = item[k]; }
      return out;
    });
  }
  var exportData = null;
  function getExportData(){
    if (exportData) return exportData;
    var node = document.getElementById('dd-export-data');
    if (!node) return { fields: [], flows: [] };
    try {
      var raw = JSON.parse(node.textContent);
      exportData = {
        fields: rehydrate(raw.fields || [], FIELD_KEY_UNMAP),
        flows: rehydrate(raw.flows || [], FLOW_KEY_UNMAP)
      };
    } catch (e) { exportData = { fields: [], flows: [] }; }
    return exportData;
  }

  var FIELD_HEADERS = ['Object', 'Subsection', 'Field Label', 'API Name', 'Type', 'Fill %', 'Created Date', 'NEW', 'Key Field', 'Key Kind', 'Key Target', 'Definition', 'Active Values', 'Source / Update', 'Owner', 'Sensitivity', 'Dependencies / Notes', 'Similar Fields'];
  var FIELD_KEYS = ['object', 'subsection', 'field_label', 'api_name', 'type', 'fill_pct', 'created_date', 'is_new', 'is_key', 'key_kind', 'key_target', 'definition', 'active_values', 'source_update', 'owner', 'sensitivity', 'dependencies', 'similar_fields'];
  var FLOW_HEADERS = ['Flow Label', 'Developer Name', 'Triggering Object', 'Process Type', 'Version', 'API Version', 'Created Date', 'Last Modified', 'Description', 'Similar Flows'];
  var FLOW_KEYS = ['flow_label', 'developer_name', 'triggering_object', 'process_type', 'version', 'api_version', 'created_date', 'last_modified', 'description', 'similar_flows'];
  var OTHER_HEADERS = ['Object', 'Field Label', 'API Name', 'Type', 'Source Type', 'Key Field', 'NEW', 'Created Date', 'Description'];
  var OTHER_KEYS = ['object', 'field_label', 'api_name', 'type', 'source_type', 'is_key', 'is_new', 'created_date', 'description'];

  function rowsForExport(items, keys){
    return items.map(function(item){
      return keys.map(function(k){
        var v = item[k];
        if (v === true) return 'TRUE';
        if (v === false) return 'FALSE';
        if (v === null || v === undefined) return '';
        return v;
      });
    });
  }
  function csvEscape(v){
    if (v === null || v === undefined) return '';
    var s = String(v);
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }
  function toCSV(headers, rows){
    var lines = [headers.map(csvEscape).join(',')];
    rows.forEach(function(r){ lines.push(r.map(csvEscape).join(',')); });
    return lines.join('\r\n');
  }
  function toTSV(headers, rows){
    var clean = function(v){ return String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' '); };
    var lines = [headers.map(clean).join('\t')];
    rows.forEach(function(r){ lines.push(r.map(clean).join('\t')); });
    return lines.join('\r\n');
  }
  function downloadBlob(text, filename, mime){
    var blob = new Blob([text], { type: mime + ';charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(function(){ URL.revokeObjectURL(url); a.remove(); }, 250);
  }

  function showToast(msg){
    var t = document.querySelector('.dd-export-toast');
    if (!t){ t = document.createElement('div'); t.className = 'dd-export-toast'; document.body.appendChild(t); }
    t.textContent = msg;
    requestAnimationFrame(function(){ t.classList.add('is-visible'); });
    clearTimeout(t._h);
    t._h = setTimeout(function(){ t.classList.remove('is-visible'); }, 2500);
  }

  var sheetJsLoading = null;
  function loadSheetJS(){
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (sheetJsLoading) return sheetJsLoading;
    sheetJsLoading = new Promise(function(resolve, reject){
      var s = document.createElement('script');
      s.src = 'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js';
      s.onload = function(){ resolve(window.XLSX); };
      s.onerror = function(){ reject(new Error('failed to load SheetJS')); };
      document.head.appendChild(s);
    });
    return sheetJsLoading;
  }

  function exportCSVFields(){
    var data = getExportData();
    var rows = rowsForExport(data.fields, FIELD_KEYS);
    downloadBlob(toCSV(FIELD_HEADERS, rows), 'Healthie_SFDC_Data_Dictionary_fields.csv', 'text/csv');
    showToast('CSV downloaded · ' + data.fields.length + ' fields');
  }
  function exportCSVFlows(){
    var data = getExportData();
    var rows = rowsForExport(data.flows, FLOW_KEYS);
    downloadBlob(toCSV(FLOW_HEADERS, rows), 'Healthie_SFDC_Data_Dictionary_flows.csv', 'text/csv');
    showToast('CSV downloaded · ' + data.flows.length + ' flows');
  }
  // Other-fields data is NOT in the JSON island (too big). Walk the rendered
  // DOM rows directly. The <td> cells are class-free to save bytes, so we
  // address them positionally: 0=label, 1=api, 2=type, 3=flag, 4=created, 5=desc.
  function collectOtherFields(){
    var rows = [];
    Array.prototype.slice.call(document.querySelectorAll('#dd-other-fields .dd-other-panel')).forEach(function(panel){
      var obj = panel.getAttribute('data-obj') || '';
      Array.prototype.slice.call(panel.querySelectorAll('tr.dd-other-row')).forEach(function(tr){
        var labelCell = tr.cells[0];
        var apiCell = tr.cells[1];
        var typeCell = tr.cells[2];
        var sourceType = (tr.getAttribute('data-type-class') || '').replace(/^./, function(c){ return c.toUpperCase(); });
        var createdCell = tr.cells[4];
        var descCell = tr.cells[5];
        var isKey = tr.classList.contains('is-key');
        var isNew = !!tr.querySelector('.dd-new-chip');
        // Strip the chip text (NEW / ★) from cell text via firstChild/textContent.
        var labelText = labelCell ? (labelCell.querySelector('a') ? labelCell.querySelector('a').textContent : labelCell.firstChild && labelCell.firstChild.textContent || labelCell.textContent).trim() : '';
        var apiText = apiCell && apiCell.querySelector('code') ? apiCell.querySelector('code').textContent.trim() : '';
        var typeText = typeCell ? typeCell.textContent.trim() : '';
        var createdText = (tr.getAttribute('data-created') || '').trim();
        var descText = descCell ? (descCell.textContent.trim() === '—' ? '' : descCell.textContent.trim()) : '';
        rows.push({
          object: obj,
          field_label: labelText,
          api_name: apiText,
          type: typeText,
          source_type: sourceType,
          is_key: isKey,
          is_new: isNew,
          created_date: createdText,
          description: descText,
        });
      });
    });
    return rows;
  }
  function exportCSVOther(){
    var rows = collectOtherFields();
    var rowsArr = rowsForExport(rows, OTHER_KEYS);
    downloadBlob(toCSV(OTHER_HEADERS, rowsArr), 'Healthie_SFDC_Data_Dictionary_all_other_fields.csv', 'text/csv');
    showToast('CSV downloaded · ' + rows.length + ' additional fields');
  }
  function exportTSVClipboard(){
    var data = getExportData();
    var rows = rowsForExport(data.fields, FIELD_KEYS);
    var tsv = toTSV(FIELD_HEADERS, rows);
    if (navigator.clipboard && navigator.clipboard.writeText){
      navigator.clipboard.writeText(tsv).then(function(){
        showToast('Copied ' + data.fields.length + ' field rows. Paste into Google Sheets.');
      }, function(){ fallbackCopy(tsv); });
    } else {
      fallbackCopy(tsv);
    }
  }
  function fallbackCopy(text){
    var ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); showToast('Copied — paste into Google Sheets.'); }
    catch (e) { showToast('Copy failed — please use the CSV download instead.'); }
    ta.remove();
  }
  function exportXLSX(){
    showToast('Building Excel workbook…');
    loadSheetJS().then(function(XLSX){
      var data = getExportData();
      var wb = XLSX.utils.book_new();
      // One sheet per object (group fields by their object).
      var byObject = {};
      data.fields.forEach(function(f){
        (byObject[f.object] = byObject[f.object] || []).push(f);
      });
      // Preserve the on-page object tab order using the DOM tabs.
      var tabOrder = Array.prototype.slice.call(document.querySelectorAll('.dd-tab')).map(function(t){
        return (t.querySelector('.dd-tab-label') || {}).textContent || '';
      });
      tabOrder.forEach(function(objLabel){
        if (!byObject[objLabel]) return;
        var rows = rowsForExport(byObject[objLabel], FIELD_KEYS);
        var aoa = [FIELD_HEADERS].concat(rows);
        var ws = XLSX.utils.aoa_to_sheet(aoa);
        // Sheet names: ≤31 chars, no special chars
        var safe = objLabel.replace(/[\/\\?*:\[\]]/g, '-').slice(0, 31);
        XLSX.utils.book_append_sheet(wb, ws, safe);
      });
      // Flows sheet
      if (data.flows.length){
        var flowRows = rowsForExport(data.flows, FLOW_KEYS);
        var flowsWs = XLSX.utils.aoa_to_sheet([FLOW_HEADERS].concat(flowRows));
        XLSX.utils.book_append_sheet(wb, flowsWs, 'Flows');
      }
      // All Other Fields sheet — pulled from the DOM (not the JSON island, which
      // omits this dataset to stay within HubSpot's 2 MiB template limit).
      var otherRows = collectOtherFields();
      if (otherRows.length){
        var otherArr = rowsForExport(otherRows, OTHER_KEYS);
        var otherWs = XLSX.utils.aoa_to_sheet([OTHER_HEADERS].concat(otherArr));
        XLSX.utils.book_append_sheet(wb, otherWs, 'All Other Fields');
      }
      XLSX.writeFile(wb, 'Healthie_SFDC_Data_Dictionary.xlsx');
      showToast('Excel workbook downloaded · ' + data.fields.length + ' fields, ' + data.flows.length + ' flows');
    }, function(err){
      showToast('Failed to load Excel exporter. Try CSV instead.');
    });
  }

  // Wire up the dropdown
  var exportBtn = root.querySelector('#dd-export-btn');
  var exportPop = root.querySelector('#dd-export-pop');
  if (exportBtn && exportPop){
    exportBtn.addEventListener('click', function(e){
      e.stopPropagation();
      var open = !exportPop.hasAttribute('hidden');
      if (open) { exportPop.setAttribute('hidden', ''); exportBtn.setAttribute('aria-expanded', 'false'); }
      else { exportPop.removeAttribute('hidden'); exportBtn.setAttribute('aria-expanded', 'true'); }
    });
    document.addEventListener('click', function(e){
      if (!exportPop.hasAttribute('hidden') && !exportPop.contains(e.target) && e.target !== exportBtn){
        exportPop.setAttribute('hidden', '');
        exportBtn.setAttribute('aria-expanded', 'false');
      }
    });
    exportPop.querySelectorAll('[data-export]').forEach(function(item){
      item.addEventListener('click', function(){
        var kind = item.getAttribute('data-export');
        exportPop.setAttribute('hidden', '');
        exportBtn.setAttribute('aria-expanded', 'false');
        if (kind === 'csv') exportCSVFields();
        else if (kind === 'csv-flows') exportCSVFlows();
        else if (kind === 'csv-other') exportCSVOther();
        else if (kind === 'tsv-clipboard') exportTSVClipboard();
        else if (kind === 'xlsx') exportXLSX();
      });
    });
  }
})();
"""


# ---------- main ----------

def main() -> int:
    parser = argparse.ArgumentParser(description="Generate the SFDC Data Dictionary v6 web page.")
    parser.add_argument("--refresh", action="store_true", help="Invalidate all caches.")
    parser.add_argument("--refresh-fielddefs", action="store_true", help="Re-fetch FieldDefinitions only.")
    parser.add_argument("--refresh-fill", action="store_true", help="Re-fetch live Fill % counts only.")
    parser.add_argument("--refresh-picklists", action="store_true", help="Re-fetch picklist values only.")
    parser.add_argument("--skip-sfdc", action="store_true", help="Skip all live SFDC calls (V5 docx only).")
    parser.add_argument("--render-only", action="store_true", help="Render only from cache.")
    args = parser.parse_args()

    print(f"[parse] reading {V5_DOCX.name}")
    v5 = parse_v5(V5_DOCX)
    print(f"[parse] objects: {len(v5['objects'])}, fields: {sum(len(s['fields']) for o in v5['objects'] for s in o['sections'])}")

    if args.refresh:
        for f in CACHE_DIR.glob("sfdc_*.json"):
            f.unlink()

    sf = SfClient(DEFAULT_ORG)
    sf_data = {"fielddefs": {}, "fills": {}, "totals": {}, "picklists": {}, "flows": [], "custom_field_dates": {}}

    if args.skip_sfdc or args.render_only:
        for api in OBJECT_API_NAME.values():
            sf_data["fielddefs"][api] = cache_load(f"sfdc_fielddef_{api}.json") or []
            sf_data["fills"][api] = cache_load(f"sfdc_fill_{api}.json") or {}
            tot = cache_load(f"sfdc_total_{api}.json") or {}
            sf_data["totals"][api] = tot.get("total")
            sf_data["picklists"][api] = cache_load(f"sfdc_picklist_{api}.json") or {}
        sf_data["flows"] = cache_load("sfdc_flows.json") or []
        sf_data["custom_field_dates"] = cache_load("sfdc_custom_field_dates.json") or {}
    else:
        for label in OBJECT_ORDER:
            api = OBJECT_API_NAME[label]
            print(f"[sfdc] {label} ({api})")
            try:
                if args.refresh_fielddefs:
                    p = cache_path(f"sfdc_fielddef_{api}.json")
                    if p.exists():
                        p.unlink()
                sf_data["fielddefs"][api] = fetch_fielddef_for_object(sf, api)
                print(f"  fielddefs: {len(sf_data['fielddefs'][api])}")
                try:
                    sf_data["totals"][api] = fetch_total_count(sf, api)
                    print(f"  total records: {sf_data['totals'][api]}")
                except Exception as e:
                    print(f"  total records failed: {e}")
                    sf_data["totals"][api] = None
                if args.refresh_picklists:
                    p = cache_path(f"sfdc_picklist_{api}.json")
                    if p.exists():
                        p.unlink()
                sf_data["picklists"][api] = fetch_picklist_values(sf, api)
                print(f"  picklist fields: {len(sf_data['picklists'][api])}")
                # collect all api names that match V5 labels for fill counts
                v5_labels = []
                for obj in v5["objects"]:
                    if obj["label"] == label:
                        for sec in obj["sections"]:
                            for f in sec["fields"]:
                                v5_labels.append(f["label"])
                fielddefs_by_api = {fd["QualifiedApiName"]: fd for fd in sf_data["fielddefs"][api]}
                resolved = []
                for lbl in v5_labels:
                    fd = match_label_to_field(lbl, sf_data["fielddefs"][api])
                    if fd and fd.get("QualifiedApiName"):
                        resolved.append(fd["QualifiedApiName"])
                resolved = sorted(set(resolved))
                print(f"  resolving fill for {len(resolved)} fields")
                if args.refresh_fill:
                    p = cache_path(f"sfdc_fill_{api}.json")
                    if p.exists():
                        p.unlink()
                sf_data["fills"][api] = fetch_fill_counts(sf, api, fielddefs_by_api, resolved)
            except Exception as e:
                print(f"[sfdc] {label} failed mid-way: {e}; continuing")
        print(f"[flows] fetching active flows")
        sf_data["flows"] = fetch_active_flows(sf, refresh=args.refresh)
        print(f"[flows] {len(sf_data['flows'])} active flows")
        print(f"[customfield-dates] fetching")
        sf_data["custom_field_dates"] = fetch_custom_field_dates(sf, refresh=args.refresh)
        print(f"[customfield-dates] {len(sf_data['custom_field_dates'])} custom fields with dates")

    model = build_model(v5, sf_data)
    for o in model["objects"]:
        if o["unmatched"]:
            print(f"[match] {o['label']}: {len(o['unmatched'])} unmatched labels: {o['unmatched'][:5]}{'...' if len(o['unmatched']) > 5 else ''}")
    now = datetime.now(timezone.utc)
    run_meta = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M UTC"),
        "generated_at_date": now.strftime("%Y-%m-%d"),
    }
    html_out = render_html(model, run_meta)
    OUTPUT_HTML.write_text(html_out)

    # Also produce a fully-wrapped standalone HTML doc for local preview.
    standalone_path = OUTPUT_HTML.with_name(OUTPUT_HTML.stem + "_standalone.html")
    standalone = (
        "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\">"
        "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">"
        "<title>Healthie SFDC Data Dictionary v6 — Preview</title>"
        "<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">"
        "<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>"
        "<link href=\"https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap\" rel=\"stylesheet\">"
        "<style>body{margin:0;padding:0;background:#f8fafc;font-family:'Inter',system-ui,sans-serif;}</style>"
        "</head><body>"
        + html_out
        + "</body></html>"
    )
    standalone_path.write_text(standalone)

    total_fields = sum(o["field_count"] for o in model["objects"])
    print(f"[done] wrote {OUTPUT_HTML} ({len(html_out):,} bytes) — {total_fields} fields, {len(model['objects'])} objects")
    print(f"[done] wrote {standalone_path} for local preview")
    return 0


if __name__ == "__main__":
    sys.exit(main())
