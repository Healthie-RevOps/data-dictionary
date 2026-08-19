# Handoff notes for the Healthie Data Dictionary project

This file is the canonical orientation for any session picking up this work. Read it before touching code.

## What this project is

Tooling that generates and publishes Healthie's RevOps Salesforce Data Dictionary as a HubSpot CMS landing page.

> **⚠ 2026-08-18 — the dictionary moved off HubSpot.** The new home is
> https://reports-production-9595.up.railway.app/peapod-sfdc-data-dictionary (Railway-hosted
> reports app, Google OAuth gated to @gethealthie.com). Both HubSpot landing pages
> (SFDC + HubSpot dictionaries) were unpublished and now 404. The HubSpot template
> `custom/pages/PP SFDC Data Dictionary.html` was replaced (draft + published envs) with a
> small meta-refresh/JS redirect to the Railway URL — source in
> [redirect_template_peapod_sfdc_data_dictionary.html](redirect_template_peapod_sfdc_data_dictionary.html).
> The redirect only serves if the landing page is republished in the HubSpot UI (page id
> `212676525863`, currently `state: DRAFT`); alternatively add a HubSpot URL redirect in
> Settings → Website → Domains & URLs. The access key lacks `content` /
> `content.landing_pages.write` scopes, so neither can be automated. The full dictionary
> template that used to live at that path is preserved locally in
> `deploy_template_peapod_sfdc_data_dictionary.html` — push it back with
> `scripts/publish_to_hubspot.py` to restore.

- **Old HubSpot page (unpublished, 404):** https://go.gethealthie.com/peapod-sfdc-data-dictionary
- **Sibling page (reference only, not generated here):** https://go.gethealthie.com/peapod-hubspot-data-dictionary
- **HubSpot portal:** healthie-prod, id `43826161`
- **HubSpot landing page id:** `212676525863`
- **Design Manager template path:** `custom/pages/PP SFDC Data Dictionary.html`
- **Salesforce target org:** `bill.coffin@gethealthie.com` (set as default via `sf` CLI)

[README.md](README.md) has the long-form overview; this file is the operational handoff.

## Primary script

[generate_sfdc_dictionary_html.py](generate_sfdc_dictionary_html.py) — single Python script (stdlib only + `sf` CLI as subprocess) that does the full pipeline:

1. Parse [Healthie_SFDC_Data_Dictionary_v5.docx](Healthie_SFDC_Data_Dictionary_v5.docx) (canonical V5 dictionary — 292 fields across 7 objects, structured as Heading1 per object + Heading2 sub-sections + 8-col tables).
2. Fetch live SFDC metadata via `sf` CLI (Tooling API):
   - `FieldDefinition` per object — labels, API names, DurableId, DataType, Description, LastModifiedDate
   - `sf sobject describe` per object — picklist active values
   - `SELECT COUNT(<field>) FROM <object>` aggregates — Fill %
   - `CustomField` org-wide — CreatedDate (~28K records, keyed by 15-char Id prefix to match DurableId)
   - `Flow` + `FlowDefinition` (Tooling) — 145 active flows
3. All caches under [`.dictionary_cache/`](.dictionary_cache/) as `sfdc_*.json`.
4. Build render model + emit HTML.
5. Wrap in HubSpot template chrome + multipart PUT to `/cms/v3/source-code/{draft,published}/content/<path>`.

CLI flags:
- `--render-only` — skip SFDC, render from cache only
- `--skip-sfdc` — same effect, no cache writes
- `--refresh` — invalidate all caches
- `--refresh-fielddefs` / `--refresh-fill` / `--refresh-picklists` — targeted refreshes

Note: there's a **second**, older script [generate_hubspot_dictionary.py](generate_hubspot_dictionary.py) that produces a .docx (not HTML) for the HubSpot side. Not active. Don't confuse.

## Current state of the page

| Element | Source | Notes |
| --- | --- | --- |
| Top navbar | Hardcoded layout | "Healthie / Salesforce Data Dictionary" breadcrumb · Export ↓ dropdown · All Fields ({N}) ↓ · Flows ({N}) ↓ · ← HubSpot Dictionary · Data Governance → · stats |
| Hero KPI strip | Computed | Fields documented · Objects (7) · Integrated sources (13) · v6 + date |
| Search bar | JS | Filters curated fields across all object tabs |
| Filter row | JS | Similar (All / ⚠ Has similar / Unique) + Created in last N days |
| Legend | Static | ★ Key field · NEW (with rolling cutoff date) · ⚠ Similar fields · ≥80/50-79/<50 fill chips |
| Object tabs | Per-object | 7 tabs with colored dots: Account=indigo, Contact=cyan, Opportunity=amber, Healthie Org=emerald, Stripe Customer=violet, Stripe Sub=pink, TaskRay=orange |
| Per-object header | Per-object | Colored top band (8px) + object title in the object's accent color + summary line (N fields · M picklists · K key fields · records · API) |
| Field tables | V5 + SFDC | Sticky Field Label column. Columns: **Field Label · API Name · Type · Fill % · Created · Definition · Active Values · Source/Update · Owner · Sensitivity · Dependencies** |
| Each field row | Compose | ★ yellow background if cross-object key field · NEW green chip if created in last 180 days · ⚠ warning if has similar siblings in same object (with sibling-list tooltip on hover) · Object Manager link on Field Label (uses DurableId so custom fields use 15-char IDs, standard fields use API name or relationship name) |
| All Other Fields | SFDC | New section, 1,290 uncurated FieldDefinitions across 7 objects. Own tab nav. Columns: Field Label · API Name · Type · Custom/Standard pill · Created · Description. Filters: search + custom/standard + days-window. |
| SFDC Flows section | SFDC | 145 active flows. Own tab nav. Columns: Flow Label · Developer Name · Object · Process Type · Version · Created · Last Modified · Description. ⚠ similar-field clustering within object groups. |
| Front-matter accordions | V5 | How to Use, Governance, Data Sources, Glossary |
| Appendix accordions | V5 | Picklist Values, V4 Field Additions, Changelog |
| Footer | Computed | V5 docx mtime + generation timestamp |

## Export functionality (recently added — see "Open issue" below)

Top-right "Export ↓" dropdown:
- **Excel workbook (.xlsx)** — multi-sheet workbook. Lazy-loads SheetJS from `https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js` on click. One sheet per curated object + Flows sheet + All Other Fields sheet.
- **CSV (fields)** — flat file with Object + Subsection columns for the 292 curated fields.
- **CSV (flows)** — 145 flow records.
- **CSV (all other fields)** — 1,290 uncurated FieldDefinitions. Pulled from the DOM (not the JSON island — see below).
- **Copy for Google Sheets** — TSV to clipboard, paste into a blank sheet.

Implementation: a JSON island `<script type="application/json" id="dd-export-data">` carries `{fields: [...], flows: [...]}`. Curated fields + flows export from this island. **The All Other Fields export walks the rendered DOM** because including it in the JSON island puts us over HubSpot's 2 MiB template size limit (see Open Issue).

## How publishing works

1. Generator renders `peapod_sfdc_data_dictionary.html` (body fragment).
2. Wraps it with HubL header + full HTML doc → `deploy_template_peapod_sfdc_data_dictionary.html`.
3. Pushes to HubSpot via `PUT /cms/v3/source-code/{environment}/content/<path>` (multipart/form-data, `file` field). Pushes both `draft` and `published` environments.
4. Required scope: `cms.source_code.write` (currently granted on the personal access key in [hubspot.config.yml](hubspot.config.yml)).
5. **Manual republish step still required** after each push: HubSpot UI → Marketing → Website → Landing Pages → "PP SFDC Data Dictionary" → click **Update**. This flushes HubSpot's prerender cache (the live URL otherwise serves a cached snapshot for up to ~10 hours).

If we ever get the legacy `content` scope (or granular `cms.pages.landing_pages.write`), the manual step can be automated by a `PATCH /cms/v3/pages/landing-pages/212676525863`. Code path is sketched in [UPLOAD_TO_HUBSPOT.md](UPLOAD_TO_HUBSPOT.md).

## HubSpot source-code size limit — 1.5 MiB on **published**, 2.0 MiB on **draft**

The "All Other Fields" + Export dropdown work crossed the source-code threshold and triggered an asymmetric rejection: draft env accepted ~1.82 MiB; published env rejected with:

```
HTTP 400 TEMPLATE_VALIDATION_FAILED
"Coded files must be smaller than 1.5 MiB"  (size: 1,804,853)
```

So the **real published limit is 1.5 MiB** (1,572,864 bytes), even though earlier notes suggested 2.0. The minified output must be < 1.5 MiB to ship both envs. Current minified size: ~1.52 MiB (~47 KB headroom).

Mitigations live in code:
1. **JSON island uses short keys** — `_build_export_data` emits records with 1-3 char keys (e.g. `"l"` for `field_label`, `"ss"` for `subsection`) via `EXPORT_FIELD_KEY_MAP` / `EXPORT_FLOW_KEY_MAP`. The JS `getExportData()` rehydrates to long names via `FIELD_KEY_UNMAP` / `FLOW_KEY_UNMAP` (these maps must stay in sync). Saved ~70 KB.
2. **JSON island drops empty values** — `_compact_rec()` strips keys whose value is `""`, `False`, `None`, `[]`, `{}` before emit. Consumer treats missing keys as empty.
3. **All Other Fields rows have no per-`<td>` classes** — CSS uses `td:nth-child(N)` selectors. JS `collectOtherFields()` reads via `tr.cells[N]`. Saved ~155 KB.
4. **All Other Fields pill classes compressed** — `dd-pill dd-pill-custom` → `p p-c`, `dd-pill dd-pill-standard` → `p p-s`. CSS rules for `.p`, `.p-c`, `.p-s` mirror `.dd-pill` styling.
5. **All Other Fields anchors drop `class="dd-field-link"`** — `.dd-other-row a` in CSS provides the same styling.
6. **Empty descriptions render as bare `—`** — no `<span class="dd-muted">` wrapper. JS treats cell text `=== "—"` as empty for export.
7. **HTML minification at push time** — `scripts/publish_to_hubspot.py` strips leading whitespace, collapses inter-tag whitespace, minifies CSS (comments + collapse), and compacts JSON island. Saves ~240 KB on top of the source changes. The local body fragment stays readable for browser preview.

The publish pipeline now lives entirely in [scripts/publish_to_hubspot.py](scripts/publish_to_hubspot.py):
- Reads `peapod_sfdc_data_dictionary.html` (body fragment from the generator)
- Wraps with HubL chrome (the `<!-- templateType: "none" -->` header + standard HTML5 head)
- Writes the wrapped result to `deploy_template_peapod_sfdc_data_dictionary.html` (for inspection)
- Minifies
- PUTs to both `draft` and `published` environments

Run it as:
```bash
HUBSPOT_TOKEN="$(python3 -c "import re; print(re.search(r'accessToken:\s*>-\s*\n\s+(\S+)', open('hubspot.config.yml').read()).group(1))")" \
  python3 scripts/publish_to_hubspot.py
```

(Or `--dry-run` to skip the network call.)

If size creeps back over the limit, next levers:
- Drop the JSON island entirely; rebuild curated-fields export by walking the curated DOM rows the same way `collectOtherFields()` does for All Other Fields.
- Drop `other_fields` content from rendered HTML entirely; lazy-build the tables from a hosted JSON file using HubSpot Files API (`files` scope is granted).
- Shorten the row-level `data-search` attribute (currently ~80 bytes × 1290 rows = ~100 KB; could be computed lazily from cell text on first keystroke).
- Split the page across multiple HubL imports (would require changing the template architecture).

**After every successful push: click Update on the HubSpot landing page UI to flush the prerender cache.**

## Key implementation details to know

### Why the URL segments work for SFDC Object Manager links
SFDC's Lightning Object Manager URLs follow this pattern:
- Standard field (e.g. `Name`): `/FieldsAndRelationships/Name/view`
- Standard reference field (`OwnerId`, `RecordTypeId`, `AccountId`): uses the **relationship name** without `Id`: `/FieldsAndRelationships/Owner/view`
- Custom field (`Health_Score__c`): uses the **15-char CustomField record ID**: `/FieldsAndRelationships/00NRc00000rWl8X/view`

`FieldDefinition.DurableId` returns the format `<Object>.<segment>` and the segment is *exactly* what the URL needs. So `setup_url_segment(durable_id, ...)` just splits on the dot. No more heuristics.

`NO_SETUP_PAGE_FIELDS` is a deny list of standard fields with no setup page (CreatedDate, LastModifiedDate, IsDeleted, etc.) — they render as plain-text labels with no link.

### NEW chip is a rolling 6-month window
`NEW_FIELD_WINDOW_DAYS = 180`. `is_new_field(durable_id, custom_field_dates)` compares `CustomField.CreatedDate` to `now() - 180 days`. The cutoff date is shown in the legend and tooltip; updates automatically each generation.

Standard fields don't have a CustomField record, so they never get a NEW chip — that's correct (they predate the org).

### Similar-field clustering
`find_similar_fields(rows)` tokenizes labels (CamelCase + `_`/`-` split, drops stopwords + short tokens), then clusters fields that share **≥2 meaningful tokens** OR **≥1 token from a distinctive product list** (Stripe, Vitally, Healthie, Calendly, Chameleon, CSAT, NPS, ARR, MRR, etc.). Each field's `similar_to` list drives the ⚠ warning + sibling-list tooltip.

Applies to both fields (within each object) and flows (clustered within their inferred triggering object so we don't get 145-flow tooltips).

### Key field classification
`classify_key(api_name, data_type)` flags a field as Key if:
- It's `Id` (primary key)
- It ends in `_ID__c` (external IDs like `Healthie_User_ID__c`, `Stripe_ID__c`)
- It's a Lookup or MasterDetail to a non-system object

The deny list `SYSTEM_LOOKUP_TARGETS` excludes Lookup(User Record Access), Lookup(Record Visibility), etc. — those are SFDC infra, not business relationships.

Key rows get a yellow background + a ★ next to the API name with a tooltip explaining the target.

### Description text normalization
`normalize_description_text(text)` runs `DESCRIPTION_TEXT_REPLACEMENTS` over every description string. Currently `[("Product Health Score", "Health Score")]` because SFDC's stale FieldDefinition.Description text on some derived fields still references "Product Health Score" even after the field itself was renamed. Add more `(old, new)` tuples here when SFDC labels drift.

V5 docx label drift is auto-corrected at render time by preferring `FieldDefinition.Label`; the docx was also edited in-place to rename "Product Health Score" → "Health Score" (backup in `go_gethealthie_backups/Healthie_SFDC_Data_Dictionary_v5_PRE-HealthScore-fix_*.docx`).

### Horizontal scroll UX
Table wraps have three UX enhancements:
1. **Inset shadow on edges** — JS toggles `.has-scroll-left` / `.has-scroll-right` classes based on `scrollLeft`/`clientWidth`/`scrollWidth`. CSS uses `box-shadow: inset` which stays at the visible edge.
2. **Wheel-to-horizontal** — vertical mouse wheel motion inside a table wrap converts to horizontal scroll (unless Shift is held, deltaX dominates, or wheel direction would scroll past an edge).
3. **Visible scrollbars** — WebKit + Firefox custom styling, 12px tall, slate-400 thumb.

`refreshAllScrollEdges()` runs on resize and after tab clicks (since hidden panels measure as 0×0 and need re-measurement when they become visible).

### CSS/JS quoting gotcha
`_JS = r"""..."""` is a **raw string** (note the `r` prefix). This is required because `\t`, `\r`, `\n` inside regex literals must remain literal backslash-escape sequences (e.g., `/[\t\r\n]+/g`). Without the raw string, Python interprets them as actual tab/CR/LF characters, which breaks JS regex literals (those can't span multiple lines). Don't drop the `r`.

`_CSS = """..."""` is a regular triple-quoted string. We don't use backslash escapes there.

### Per-object color palette
`OBJECT_COLOR` constant near the top — change here to recolor everything object-related (tab dots, section headers, panel summary borders, subsection title borders, etc.).

## Important files inventory

| File | Purpose |
|---|---|
| [generate_sfdc_dictionary_html.py](generate_sfdc_dictionary_html.py) | The whole pipeline — single file (~2000+ lines). Top declares constants + helpers; bottom has `_CSS`, `_JS`, `render_html`, and `main`. |
| [Healthie_SFDC_Data_Dictionary_v5.docx](Healthie_SFDC_Data_Dictionary_v5.docx) | Canonical V5 dictionary source. Don't edit directly without backing up first; the generator parses this. |
| [hubspot.config.yml](hubspot.config.yml) | HubSpot CLI config; contains the personal access key + an access token cache. `hs accounts info` refreshes the token. |
| [peapod_sfdc_data_dictionary.html](peapod_sfdc_data_dictionary.html) | Latest rendered body fragment (output of last generator run). |
| [peapod_sfdc_data_dictionary_standalone.html](peapod_sfdc_data_dictionary_standalone.html) | Same body wrapped in `<html>`/`<head>` for local browser preview. |
| [deploy_template_peapod_sfdc_data_dictionary.html](deploy_template_peapod_sfdc_data_dictionary.html) | Last template pushed to HubSpot. **Currently stale** until the in-flight minified push lands. |
| [UPLOAD_TO_HUBSPOT.md](UPLOAD_TO_HUBSPOT.md) | Operational walkthrough. |
| [README.md](README.md) | Project overview (recently updated with GitHub Actions references — note those workflows don't exist yet locally; user added them as future-state). |
| [.dictionary_cache/](.dictionary_cache/) | Per-object SFDC metadata caches. Survives across runs unless `--refresh*` is used. |
| [go_gethealthie_backups/](go_gethealthie_backups/) | Timestamped snapshots of live pages + V5 docx pre-rename backup. |

## Known caveats

- **TaskRay Project SOQL is denied** — `SELECT COUNT() FROM taskray__Project__c` returns "sObject type not supported". User-permission gap on the managed package. TaskRay rows render with `—` for Fill %. Don't try to "fix" by adjusting casing — it's a permission issue.
- **One unmatched V5 label** — Opportunity → "Using Mobile White Label" doesn't resolve to a FieldDefinition. Likely renamed in SFDC; V5 hasn't reflected it. Renders as plain text (no link).
- **Token expiry** — HubSpot access tokens expire hourly. The push script auto-runs `hs accounts info` before pushing to refresh the token cached in hubspot.config.yml. If that's not running, push will 401.
- **HubSpot 2.0 MiB source-code size limit** — see Open Issue above. Don't add new content that would push us over without considering this.
- **JSON island has only fields + flows** — `other_fields` were dropped to stay under the size limit. If you need to add data, walk the DOM in JS rather than adding to the JSON island.

## Last-known-good push state

Last successful push (both draft + published): 2026-05-20, the **minified All Other Fields + Export dropdown** build. Minified size 1,525,024 bytes (~1.45 MiB), under the 1.5 MiB published-env limit by ~47 KB. After this push, the user still needs to click **Update** in HubSpot's landing page UI to flush the prerender cache (the live URL otherwise serves the previously-prerendered HTML for up to ~10 hours).

## Recently completed features (in rough order)

1. Initial port from V5 docx to web dictionary
2. HubSpot CMS publishing via `cms.source_code.write`
3. SFDC Object Manager hyperlinks per field (using DurableId for correct URL segments)
4. Per-object accent colors (tab dots, section headers, etc.)
5. ★ Key field highlighting + tooltip
6. ⚠ Similar field clustering + sibling-list tooltips
7. Two-row description layout, then reverted to columns with sticky Field Label
8. Top navbar with cross-links + stats
9. SFDC Flows section (145 active flows)
10. NEW field chip (rolling 6-month window via CustomField.CreatedDate)
11. Created Date column + Similar filter + Created-in-last-N-days filter (both for fields and flows)
12. Horizontal scroll UX: edge shadows, wheel-to-horizontal, visible scrollbars
13. JS-managed tooltips (replacing buggy native `title=`)
14. Export dropdown (XLSX / CSV / TSV-to-clipboard) — **PUSHED**
15. All Other Fields section (1,290 uncurated FieldDefinitions) — **PUSHED**
16. Compression pass for HubSpot's 1.5 MiB published source-code limit (short JSON keys, class-free `<td>`s in other-rows, push-time minification) — **PUSHED**

## Style preferences observed across the session

- The user wants **clean, professional UX** — they had specific opinions about column visibility (Created column was hidden behind horizontal scroll → reposition), tooltip mechanism (native `title` was unreliable → JS-managed), filter composability (filters should AND together).
- They mirror the HubSpot dictionary's visual style intentionally (white navbar, indigo accent, Inter font, similar nav-link pill style, similar legend).
- They prefer **direct fixes** over workarounds (e.g., when the SFDC Object Manager URLs were broken, they wanted a real audit + fix, not just a denylist of "fields that don't link").
- When labels drift between V5 docx and SFDC live, **SFDC's live label wins** as the authoritative source. We also fixed the docx in place where appropriate.
- They explicitly want me to keep pushing to HubSpot after each meaningful change.

## Next likely asks

The user has been steadily adding capabilities to this page. Possible directions:
- More export formats / scheduled exports
- Automated nightly regeneration (the README references `.github/workflows/sfdc-dictionary-sync.yml` and `github-action/README.md` — neither exists yet)
- Validation rules section (similar to Flows section)
- Apex triggers / classes section
- Per-field "where is this used" cross-references
- Dark mode
- Mobile/tablet layout improvements
