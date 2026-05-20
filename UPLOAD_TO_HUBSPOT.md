# Publishing the SFDC Data Dictionary to HubSpot

## TL;DR

The dictionary is now published end-to-end via API. The generator pushes the rendered template directly to HubSpot's Design Manager:

```bash
python3 generate_sfdc_dictionary_html.py
```

This regenerates [peapod_sfdc_data_dictionary.html](peapod_sfdc_data_dictionary.html) from V5 + live SFDC. To actually deploy it, run the push script (or invoke the deploy function from the generator — both push the wrapped template to `custom/pages/PP SFDC Data Dictionary.html`).

**One manual step still required:** after the push, click **Update** on the landing page in HubSpot's UI to flush its prerender cache. HubSpot stores a pre-rendered snapshot of landing pages independent of the underlying template; updating the template doesn't auto-rebuild the snapshot.

## What the page actually is

| Field | Value |
| --- | --- |
| Content type | LANDING_PAGE |
| Id | `212676525863` |
| Name | "PP SFDC Data Dictionary" |
| Slug | `peapod-sfdc-data-dictionary` |
| URL | https://go.gethealthie.com/peapod-sfdc-data-dictionary |
| Template path | `custom/pages/PP SFDC Data Dictionary.html` |
| Hosting | HubSpot CMS, portal `healthie-prod` (id 43826161) |

The template is a static custom HTML file (not a drag-and-drop modular page), so the publish target is the template file in Design Manager. The landing-page record itself points at this template.

## How the API push works

Required HubSpot scope on the personal access key: **`cms.source_code.write`** (plus `cms.source_code.read` for the round-trip verify). The token in [hubspot.config.yml](hubspot.config.yml) currently has this scope.

The endpoint:

```
PUT /cms/v3/source-code/{environment}/content/custom/pages/PP SFDC Data Dictionary.html
```

with `environment ∈ {draft, published}`. Body must be multipart/form-data with a `file` field — the API rejects raw `text/html` or JSON. The push routine in the generator handles this. Both environments are PUT on each push so draft and published stay in sync.

## Manual republish (HubSpot UI step)

After the API push lands, the live URL still serves the previously prerendered HTML for up to ~10 hours because of HubSpot's `s-maxage=36000` prerender cache. To trigger republish:

1. Log in to HubSpot, portal **healthie-prod**.
2. Top nav → **Marketing** → **Website** → **Landing Pages**.
3. Open "PP SFDC Data Dictionary".
4. Click **Update** (top right). HubSpot will pick up the new template content and regenerate the prerender snapshot. The live page updates within ~1 minute.

If you want to skip this step entirely, the missing piece is the legacy `content` scope (or granular `cms.pages.landing_pages.write`) on the access key — that would let us PATCH the page directly via API to trigger republish. The push routine in the generator is wired to detect this and add the page touch when the scope is present.

## Verification after publish

Spot-check on https://go.gethealthie.com/peapod-sfdc-data-dictionary:

- All 7 object tabs render (Account / Contact / Opportunity / Healthie Organization / Stripe Customer / Stripe Subscription / TaskRay Project)
- Hero KPI strip shows current numbers (e.g., **292 fields · 7 objects · 13 integrated sources · v6**)
- Global search filters fields across tabs in real time (try "stripe" → ~73 matches across 5 tabs)
- Fill % chips render as colored numerics
- A picklist row (e.g., Account → Type) shows its active values in the Active Values column
- **★ Key** field rows (yellow shaded) appear and link to SFDC Object Manager
- **NEW** green chips render on fields created in the past 6 months
- **⚠ Similar fields** tooltips appear on hover for clustered fields
- Object section headers have distinct colored top bands
- The "Flows (145) ↓" nav link scrolls to the SFDC Flows section at the bottom
- Footer shows the V5 docx mtime + generation date

## Rollback

If something goes wrong:

1. **Easiest** — HubSpot Page History. Marketing → Website → Landing Pages → "PP SFDC Data Dictionary" → gear icon → **Page history**. Restore the previous version.
2. **Source-code rollback** — the previous template version is the latest non-current entry in HubSpot's source-code history. Or push the backup template from `go_gethealthie_backups/template_PP_SFDC_Data_Dictionary_*.html` (timestamped, captured before any change was made).
3. **Local backups** — `go_gethealthie_backups/peapod-sfdc-data-dictionary_*.html` is the pre-change snapshot of the rendered live page (228 KB). Worst case, re-create the page from this HTML.
