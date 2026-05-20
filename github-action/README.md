# Weekly SFDC Data Dictionary sync — GitHub Action

**Repo:** https://github.com/healthiebc/data-dictionary

This is the replacement for the Cowork-hosted scheduled task. It runs the existing
`generate_sfdc_dictionary_html.py` against live Salesforce and pushes the resulting
template to HubSpot's CMS, every Monday at noon UTC. No laptop or sandbox involvement.

## Why we moved off Cowork

Cowork's sandbox blocks outbound connections to `login.salesforce.com`,
`*.my.salesforce.com`, `api.hubapi.com`, and `go.gethealthie.com`. Until that allowlist
changes, the sync has to run somewhere with normal internet access. GitHub Actions has
both the network access and the cron scheduling built in, so the work lives there now.

The Cowork scheduled task (`weekly-sfdc-data-dictionary-sync`) has been paused. Don't
re-enable it unless you've also gotten the four domains above added to Cowork's network
allowlist.

## File layout to drop into your repo

```
<repo root>/
├── .github/
│   └── workflows/
│       └── sfdc-dictionary-sync.yml    ← from this folder
├── scripts/
│   └── publish_to_hubspot.py            ← from this folder
├── generate_sfdc_dictionary_html.py     ← existing, unchanged
├── Healthie_SFDC_Data_Dictionary_v5.docx
├── .dictionary_cache/                   ← existing, kept under version control
├── deploy_template_peapod_sfdc_data_dictionary.html  ← generator output, kept under VC
├── peapod_sfdc_data_dictionary.html
├── peapod_sfdc_data_dictionary_standalone.html
└── hubspot.config.yml                   ← still useful locally; NOT needed by the Action
```

Copy `.github/workflows/sfdc-dictionary-sync.yml` and `scripts/publish_to_hubspot.py`
into a repo that already has the generator and the V5 docx. If you don't have such a
repo yet, `git init` your workspace folder and push it as a private repo — the workflow
expects everything to be in the repo root.

## One-time setup

### 1. Create the two repository secrets

In the repo, go to **[Settings → Secrets and variables → Actions](https://github.com/healthiebc/data-dictionary/settings/secrets/actions) → New repository secret**
and add:

#### `SFDX_AUTH_URL`

The same value you've been using in Cowork's `.sf-auth/auth-url.txt`. On your laptop:

```bash
sf org display --target-org bill.coffin@gethealthie.com --verbose --json
```

Copy the value of `result.sfdxAuthUrl` (starts with `force://`). Paste it as the secret
value. **Single line, no quotes.**

#### `HUBSPOT_TOKEN`

The personal access key from `hubspot.config.yml` — the value of `personalAccessKey`
under the `healthie-prod` portal. It already has the `cms.source_code.write` and
`cms.source_code.read` scopes the publisher needs.

If you'd rather generate a fresh token, in HubSpot: **Settings → Integrations →
Private Apps → Create private app → Scopes:** select `cms.source_code.read` and
`cms.source_code.write` → **Create**. Copy the access token. Paste it as the secret
value.

### 2. Sanity-check the workflow without publishing

Once the secrets are in, go to **Actions → SFDC Data Dictionary weekly sync → Run workflow**.
Set "Generate HTML but do not push to HubSpot" to `true` and run it. This exercises:

- CLI install
- Salesforce login via your auth URL
- Full schema pull and HTML generation
- Caches committed back to the repo

…but skips the actual HubSpot push, so if any wiring is off you'll see it without affecting
the live page. The generated HTML lands as a workflow artifact you can download from the run
page.

### 3. Run the real thing

Re-run with "skip_publish" left as `false`. The Monday cron will pick it up automatically
from then on; the manual button is there for off-cycle re-runs.

## What the workflow does, step by step

1. Checks out the repo.
2. Installs Python 3.11 and Node 20, then `npm install -g @salesforce/cli`.
3. Writes `$SFDX_AUTH_URL` to a temp file (never echoed) and runs
   `sf org login sfdx-url --sfdx-url-file ... --set-default --alias prod`.
4. Runs `python3 generate_sfdc_dictionary_html.py`. This:
   - Reads V5 docx.
   - Pulls FieldDefinitions, totals, fill counts, picklists, flows, and CustomField
     CreatedDates from Salesforce.
   - Re-renders the three HTML files (`peapod_sfdc_data_dictionary.html`, the standalone
     preview, and `deploy_template_peapod_sfdc_data_dictionary.html`).
   - Updates the `.dictionary_cache/` JSON files.
5. Runs `python3 scripts/publish_to_hubspot.py`, which PUTs the deploy template to
   `https://api.hubapi.com/cms/v3/source-code/{draft,published}/content/...`.
6. Commits any changes to the generated HTML files and `.dictionary_cache/*` back to the
   repo, so the next run has the prior snapshot to diff against.
7. Uploads the three HTML files as a workflow artifact (90-day retention).
8. Writes a markdown summary to the run page (cron, sizes, prerender-cache reminder).

## Manual step still required after each successful run

HubSpot caches a prerendered copy of the live landing page independent of the template
file. The API push updates the template but does not flush that cache, so for up to ten
hours the live page may still serve the previous content.

To flush immediately: open the **PP SFDC Data Dictionary** landing page in HubSpot's
UI and click **Update** in the top-right. This is the same manual step you've been
doing in the existing manual workflow — nothing new here.

## Rotating credentials

If `SFDX_AUTH_URL` ever stops working (refresh token revoked, password reset, MFA reset),
re-run `sf org display --target-org bill.coffin@gethealthie.com --verbose --json` on your
laptop, grab the new `sfdxAuthUrl`, and overwrite the GitHub secret.

If `HUBSPOT_TOKEN` ever stops working (token revoked or scopes changed), generate a fresh
private-app token with `cms.source_code.write` and overwrite the secret.

## Rollback / pause

To skip a week: **Actions → SFDC Data Dictionary weekly sync → ⋯ → Disable workflow.**
Re-enable when ready. The repo and its committed snapshots aren't touched.

To revert a bad sync: the HTML files and `.dictionary_cache/*` are committed by the
workflow, so each Monday's run is a single commit you can revert with `git revert`.
For the HubSpot side, the previous template version is recoverable from the
`go_gethealthie_backups/` directory and via HubSpot's own revision history on the
source-code file in Design Manager.

## What this workflow does NOT do

- It does not edit the V5 docx (`Healthie_SFDC_Data_Dictionary_v5.docx`). Source-of-truth
  content lives in that docx; the workflow is read-only against it. If you want a field's
  description, sensitivity, owner, or notes to change, edit the docx, commit it, and
  re-run.
- It does not flag rename/retype changes for review. The existing generator silently
  prefers `FieldDefinition.Label` over the docx, so renames just propagate. If you want
  a "human review required" gate, that's a follow-up task on top of this.
- It does not click **Update** in HubSpot's UI for you. That step is on a human.

## Smoke-testing locally

You can run the publish helper without touching HubSpot:

```bash
HUBSPOT_TOKEN=anything python3 scripts/publish_to_hubspot.py --dry-run
```

It will load the deploy template, build the multipart body, and print what it would send —
useful when you change the template path or HubSpot's source-code URL contract.

Without a token it refuses to run at all (exit code 2), so a misconfigured workflow can't
accidentally send an unauthenticated request.
