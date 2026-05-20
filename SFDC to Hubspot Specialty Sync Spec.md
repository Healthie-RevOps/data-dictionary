# Specialty Sync Strategy — Salesforce ↔ HubSpot

## Context

Specialty data is captured in three systems — Healthie (the product), Salesforce, and HubSpot — and the values drift across them. The goal is to design a systematic, automated sync that eliminates the discrepancies between **Salesforce Account.Healthie_Specialty_v2__c** (and the derived SFDC Contact formula) and the **HubSpot specialty property** on Contact and Company.

This document is the **strategy and architecture recommendation only**. No code is to be written and no system is to be touched until the open decisions in §4 are confirmed.

---

## 1. Verified Current State (corrections to the problem statement)

The original problem statement contained three assumptions that discovery did not bear out. The design must reflect the verified picture.

### Salesforce side

| Field | Object | Type | Calculated? | Formula | Populated | Notes |
|---|---|---|---|---|---|---|
| `Healthie_Specialty_v2__c` | Account | Text(255) | **No** | — | 11,838 / 335,992 | API-written. **Manually editable.** |
| `Specialty__c` | Account | Multipicklist | No | — | 139 | Legacy, deprecate. |
| `Healthie_Speciality__c` | Contact | Formula (Text) | **Yes** | `Account.Healthie_Specialty_v2__c` | 57,242 | **Note British spelling — extra "i".** Read-only by definition. |
| `Contact_Specialty__c` | Contact | Picklist | No | — | 0 | Unused. |
| `Provider_Specialty__c` | Contact | Multipicklist | No | — | 59 | Sales-entered, rare. |
| `Provider_Specialty__c` | Lead | Multipicklist | No | — | 160 | Pre-conversion. |
| `Healthie_Primary_Specialty__c` | Healthie_Organization__c | Text(255) | No | — | 10,881 / 193,178 | **Healthie product DB → SFDC.** True upstream. |

Key corrections:
- **The Account field is NOT a formula** — it's a plain text field that three integrations write to (Healthie Integration is the primary, dominant writer; Stripe and the HubSpot AppExchange package also touch it). It is also manually editable, which is the source of write-collision risk.
- **The Contact formula DOES exist**, but the API name is **`Healthie_Speciality__c`** (British "Speciality" — one extra `i`). It cannot be written to directly; writes must target the parent Account.
- **No flows, triggers, validation rules, or workflow field updates currently fire on `Healthie_Specialty_v2__c`.** New automation will not conflict, but there is also no existing audit trail.
- **The real upstream is the Healthie product DB**, which writes `Healthie_Organization__c.Healthie_Primary_Specialty__c` and (through the same integration) `Account.Healthie_Specialty_v2__c`. The Account field is itself downstream of Healthie, not a free-standing master record.

### HubSpot side

HubSpot has **six specialty-named properties on Contact**, none well-populated, with conflicting purposes:

| HubSpot Property | Type | Object | Population | Origin |
|---|---|---|---|---|
| `contact_specialty` | Dropdown | Contact | 0% | Already labeled "SFDC-synced" — but empty |
| `provider_specialty` | Multi-checkbox | Contact | 0% | Labeled SFDC-synced, multi-select |
| `provider_specialty_text` | Text | Contact | 1% | SFDC-synced companion |
| `specialtyii` | Dropdown | Contact | 2% | Form submissions |
| `healthie_specialty` | Text | Contact | 5% | Form submissions — currently the most populated |
| `specialty` | Text | Contact | 0% | Form submissions |
| `what_is_your_specialty` | Dropdown (60+ values) | Contact | 0% | Form submissions |

Equivalent properties exist on Company at similar (low) populations.

Volume: **~467k Contacts, ~343k Companies, ~15k Deals** in HubSpot portal `healthie-prod` (43826161).

### Cross-system gaps

- **30,291 SFDC Contacts (11.2%) are orphaned** (no `AccountId`) — the formula chain doesn't reach them.
- **HubSpot's existing "SFDC-synced" specialty properties are 0% populated**, despite a sync supposedly being in place. This is a major signal that the HubSpot AppExchange connector is not currently mapping specialty, or the mapping is mis-configured.
- HubSpot's actual populated specialty values come from forms landing in `healthie_specialty` (5%) and `specialtyii` (2%) — neither of which is wired to SFDC today.

---

## 2. Confirmed Source-of-Truth Decision

**Hybrid: Salesforce wins for existing customers, HubSpot wins for prospects.**

Customer/prospect classification signal:
- **Customer**: Account has at least one related `Healthie_Organization__c` record (always present for customers, since Healthie Integration creates one). Verify via `SELECT Id FROM Healthie_Organization__c WHERE Account__c = :acctId LIMIT 1`.
- **Prospect**: No related `Healthie_Organization__c`. Specialty data, if any, originates from HubSpot forms.
- **Edge case — Lead-stage**: HubSpot contact has no SFDC Contact yet, only a Lead. Sync lands on `Lead.Provider_Specialty__c` until conversion.

**On conversion (Lead → Account)**: Lead.Provider_Specialty__c → Account.Healthie_Specialty_v2__c, then the customer-flow rules take over.

**On customer transition (prospect Account → customer Account)** — i.e., when a Healthie_Organization__c is first created for an Account: Healthie_Specialty_v2__c is written by the Healthie Integration; this **overrides** any prior HubSpot-sourced value. This is acceptable because Healthie's value reflects the actual provisioned specialty.

---

## 3. Open Decisions Before Build

Two decisions must be locked before writing any code:

### 3.1 Which HubSpot property is canonical?

The team has six fragmented specialty properties. The plan recommends:

**Recommendation**: Designate `contact_specialty` (Dropdown, on Contact) and a matching `company_specialty` (or whichever exists on Company today; verify) as the **single canonical pair**. Why:
- It is already labeled as the SFDC-sync target — minimizes new property creation.
- Dropdown enforces value consistency; the text-based form properties drift.
- Existing 0% population means there is no data to lose by re-wiring.

Required preparation work — **before the sync goes live**:
1. **Value-domain consolidation**: collect distinct values from `specialtyii`, `healthie_specialty`, `specialty`, `what_is_your_specialty` (HubSpot side) and `Account.Healthie_Specialty_v2__c` (SFDC side). Build a single controlled list. Add missing values to the `contact_specialty` dropdown.
2. **Form rewiring**: change all HubSpot forms currently writing to other specialty properties to write to `contact_specialty` instead, with a dropdown mapping.
3. **Legacy property migration**: backfill `contact_specialty` from the form-sourced properties in priority order (`healthie_specialty` → `specialtyii` → `what_is_your_specialty` → `specialty`).
4. **Deprecation**: hide the legacy properties in HubSpot UI; do not delete until sync has run cleanly for 30 days.

**Marketing Ops sign-off required** on this consolidation before build.

### 3.2 What about the multi-select `provider_specialty` property?

This is a multi-checkbox in HubSpot, paired with `Contact.Provider_Specialty__c` (multipicklist) in SFDC. The semantic is different: this represents *areas of expertise* (potentially multiple), where `contact_specialty` represents *primary specialty* (single).

**Recommendation**: Treat `provider_specialty` as **out of scope** for this sync. It's a different concept and a different field architecture (multi vs single). Document this explicitly so the design doesn't get re-litigated mid-build.

---

## 4. Recommended Architecture

### 4.1 Architecture choice — Hybrid B+C (SFDC Flow outbound + HubSpot Workflow webhook inbound), with classification gating

| Option | Verdict | Reason |
|---|---|---|
| A. HubSpot native SFDC connector | **Rejected** | Cannot do conditional direction by customer/prospect status. Cannot handle formula-target safely. |
| B. SFDC Flow outbound only | Partial | Handles customer flow only. |
| C. HubSpot Workflow + webhook inbound only | Partial | Handles prospect flow only. |
| D. iPaaS (Workato/Zapier/Make) | Acceptable fallback | Polling latency, ongoing cost, vendor lock-in. Worth considering if internal API-callout maintenance is unwanted. |
| E. Custom Python service | Rejected for v1 | Highest build/maintenance cost. Reserve for v2 if v1 hits scaling limits. |
| F. Platform Events + Webhooks (event-driven) | Future state | Best long-term real-time design, but over-engineered for current volume (~12k populated Account records). |

**Recommended v1: hybrid B+C, gated by customer/prospect classification.**

### 4.2 Component design

**Customer flow — SFDC → HubSpot** (B):

1. **Trigger**: Record-triggered Flow on `Account` after-save, condition: `Healthie_Specialty_v2__c IS_CHANGED` AND `Account has related Healthie_Organization__c`.
2. **Action**: Invocable Apex callout to HubSpot CRM API (`PATCH /crm/v3/objects/companies/{hubspotId}`) updating `contact_specialty` (Company-side equivalent) and a separate callout to update the associated Contacts. SFDC↔HubSpot ID mapping comes from the existing HubSpot AppExchange package's `HubSpot_Inc__HubSpot_Id__c` field.
3. **Loop suppression**: Set `Sync_Source__c = 'SFDC'` and `Last_SFDC_Sync__c = NOW()` on Account before callout (single field, single update — does not re-trigger the flow because the gating condition is `IS_CHANGED(Healthie_Specialty_v2__c)`).
4. **Error handling**: Failures write to a new `Sync_Error_Log__c` custom object with record ID, payload, error message, retry count. Slack notification via existing `SlackApiService` Apex class for retry-exhausted records.

**Prospect flow — HubSpot → SFDC** (C):

1. **Trigger**: HubSpot Workflow on Contact, condition: `contact_specialty IS_KNOWN` AND `associated_company.salesforce_account_id IS_KNOWN` AND that SFDC Account has NO Healthie_Organization (lookup via webhook). For Contacts not yet associated with a SFDC Account, fall through to Lead path.
2. **Action**: HubSpot webhook → AWS Lambda (or named API endpoint) → SFDC Apex REST endpoint → updates `Account.Healthie_Specialty_v2__c` (which propagates to Contact via existing formula).
3. **Lead path**: If associated Contact has no SFDC Account but has a SFDC Lead (matched on email), update `Lead.Provider_Specialty__c` instead.
4. **Loop suppression**: SFDC REST endpoint sets `Sync_Source__c = 'HubSpot'` before write.

**Conflict resolution (the cross-direction guard)**:
- Before any write, the receiving side checks `Sync_Source__c` and `Last_*_Sync__c` timestamps.
- If a customer write arrives from HubSpot for an Account where `Healthie_Organization__c` exists → **reject silently, log to `Sync_Error_Log__c` as "ignored: customer record"**.
- If a prospect write arrives from SFDC for an Account where no `Healthie_Organization__c` exists → **proceed** (rare, but legitimate when a sales rep manually pre-populates).

### 4.3 Why not the HubSpot AppExchange connector for the customer flow?

The installed `HubSpot Integration` package (namespace `HubSpot_Inc`, Daiquiri version) does support property mapping between Account and Company. However:
- It does not natively support conditional sync direction by classification.
- It will not respect `Sync_Source__c`-style suppression flags.
- It already exists but is not currently mapping specialty (HubSpot side at 0%) — fixing this would require a property-mapping config change that we don't have visibility into.
- It's a black box for debugging.

**Recommendation**: Use the connector for ID mapping (already there) but bypass it for the actual specialty write — use Apex callouts so we own the logic.

---

## 5. Data Mapping Rules

| Concern | Rule |
|---|---|
| **Value mapping** | Build a `Specialty_Value_Mapping__mdt` custom metadata type with `HubSpot_Value__c` and `SFDC_Value__c`. Required before go-live. Populate from the value-domain consolidation in §3.1. |
| **Case sensitivity** | Normalize to title case on both sides. Trim whitespace. |
| **Blank/null** | If incoming value is null: skip write (do not overwrite a populated field with null). Exception: if `Sync_Source__c` matches the incoming direction and last-sync timestamp is newer than the existing value, allow null (explicit clear). |
| **Unknown values** | If HubSpot value has no SFDC mapping in `Specialty_Value_Mapping__mdt`: write to a staging field `Account.Pending_Specialty_Review__c` (new), notify via Slack `#data-quality`, do NOT touch `Healthie_Specialty_v2__c`. Same in reverse for SFDC→HubSpot. |
| **Length** | SFDC field is 255 chars; HubSpot dropdown values are short. Reject and log anything >255. |
| **Multipicklist vs single** | Out of scope per §3.2. The `provider_specialty` checkbox in HubSpot does NOT sync to `Account.Specialty__c` multipicklist. |

---

## 6. Risks & Mitigations

| Risk | Severity | Mitigation |
|---|---|---|
| **Healthie Integration races with the new sync** writing to `Healthie_Specialty_v2__c` | High | Confirm Healthie Integration's write schedule. If it writes on every Healthie change, the SFDC→HubSpot flow will fire on every Healthie update. Add a `Sync_Source__c = 'Healthie'` discriminator so the flow can deduplicate. |
| **Formula field write attempts** — engineer copies code path that writes to `Contact.Healthie_Speciality__c` | Medium | Document the formula's read-only nature loudly. Add a Salesforce validation rule that prevents Contact direct-edits (not applicable to formula, but useful warning in onboarding docs). |
| **Field naming typo** (`Speciality` vs `Specialty`) causes integration errors | Medium | Document everywhere. Do NOT rename the formula field (would break dependent reports/dashboards) — instead create a Description on the field and a CLAUDE.md/data-dict entry calling out the British spelling. |
| **30k orphaned Contacts** never get specialty | Medium | Run a one-time backfill: for orphaned Contacts with a matching email in HubSpot, attach to the appropriate Account. Out of scope for this sync but flag for follow-up. |
| **Circular overwrite / infinite loop** | High | `Sync_Source__c` flag, last-sync timestamps, and the "do not write if source matches recipient" gate prevent loops. Test on a sandbox before go-live. |
| **HubSpot AppExchange package writes specialty independently** | Medium | The package currently does NOT populate `contact_specialty` (0% across all HubSpot specialty properties from SFDC). Verify in package config that no other mapping is enabled. If enabled, disable. |
| **HubSpot forms continue submitting to legacy properties** post-cutover | High | Rewire forms BEFORE go-live (per §3.1). Audit all forms; one missed form = ongoing drift. |
| **Multi-currency / multi-language complications** | Low | Healthie is US-centric; not currently an issue. |
| **Historical drift never reconciles** | High | One-time reconciliation script before go-live (per §7, Phase 1). |
| **Customer/prospect classification is wrong** for some edge cases | Medium | Define an allow-list of Account.RecordTypeId values that count as "customer." Healthie_Organization__c presence is the primary signal; RecordType is the fallback. |
| **The HubSpot Integration package's own ID mapping breaks** | Medium | Use `HubSpot_Inc__HubSpot_Id__c` as primary key for HubSpot↔SFDC linkage. If null, fall back to email match (Contact) or domain match (Company). |

---

## 7. Implementation Sequence

The sequence is structured so each phase is reversible and verifiable.

### Phase 0 — Data audit & open decision resolution (estimated 1 week, no code)

- Export `Account` + `Healthie_Specialty_v2__c` + Healthie_Organization presence + RecordTypeId — 336k records.
- Export HubSpot Company + Contact + all six specialty properties via the existing `generate_hubspot_dictionary.py` (already idempotent, already cached).
- Side-by-side spreadsheet comparison: count mismatches, blank-vs-value gaps, value-domain divergence.
- Output: a Specialty Value Mapping table (the source data for `Specialty_Value_Mapping__mdt`).
- Marketing Ops sign-off on HubSpot canonical property choice (§3.1).
- **Gate**: do not proceed to Phase 1 until both the value mapping table and the marketing decision are in writing.

### Phase 1 — Foundation (1 week)

- Create the new SFDC custom metadata `Specialty_Value_Mapping__mdt` (already-active config object).
- Create new SFDC fields on Account: `Sync_Source__c` (picklist: Healthie/SFDC/HubSpot), `Last_SFDC_Sync__c` (DateTime), `Last_HubSpot_Sync__c` (DateTime), `Pending_Specialty_Review__c` (Text 255).
- Create `Sync_Error_Log__c` custom object.
- Build the value-mapping records in `Specialty_Value_Mapping__mdt`.

### Phase 2 — HubSpot consolidation (1–2 weeks, depends on form count)

- Backfill `contact_specialty` from legacy form properties (priority order in §3.1).
- Rewire each form to write to `contact_specialty`. Audit form submissions for 7 days to confirm.
- Hide legacy properties in UI (keep data; do not delete).

### Phase 3 — One-time reconciliation (1 week)

- Script (run from local dev env, not deployed): for every SFDC Account with `Healthie_Specialty_v2__c` populated, find the matched HubSpot Company by `HubSpot_Inc__HubSpot_Id__c` and write the SFDC value to `contact_specialty` (mapped via the metadata table). For prospects (Account without Healthie_Organization__c), do the reverse: if HubSpot has a value and SFDC doesn't, write to Account.
- Dry-run first, output diff CSV for human review.
- After approval, run the live reconciliation.

### Phase 4 — Customer flow (SFDC → HubSpot) (1 week)

- Build the SFDC record-triggered Flow + invocable Apex callout.
- Test in a sandbox: set up 20 test Accounts, half with Healthie_Organization, half without; toggle `Healthie_Specialty_v2__c`, verify HubSpot gets only the "customer" updates.
- Deploy via DevOps Center to production behind a Custom Permission gate (`Specialty_Sync_Enabled`) so it can be turned off without redeploy.

### Phase 5 — Prospect flow (HubSpot → SFDC) (1–2 weeks)

- Build the HubSpot Workflow, webhook receiver, and SFDC Apex REST endpoint.
- Mirror sandbox testing approach: prospect Contacts only, verify writes land on Account or Lead correctly, verify customer-Account writes are silently dropped.
- Deploy behind the same permission gate.

### Phase 6 — Observability (concurrent with 4–5)

- `Sync_Error_Log__c` records with retry counts.
- Slack `#data-quality-alerts` notification on retry-exhausted records (use existing `SlackApiService` from the in-house Slack integration).
- Daily digest job (Apex Scheduler): count of syncs by direction, error rate, top-5 unmapped values for the day.

### Phase 7 — Go-live & monitor (2 weeks)

- Enable Custom Permission for 10% of internal users (via Permission Set) for 3 days.
- Expand to 100% if error rate <2%.
- Monitor for 2 weeks; revisit value mapping table weekly.

---

## 8. Verification Plan

End-to-end test scenarios (must pass on sandbox before production go-live):

1. **Customer SFDC→HubSpot**: Account has Healthie_Organization. Manually update `Healthie_Specialty_v2__c` to "Cardiology". Verify within 30s:
   - Matched HubSpot Company `contact_specialty` = "Cardiology" (post-mapping).
   - Associated HubSpot Contacts' `contact_specialty` = "Cardiology".
   - Account `Sync_Source__c` = 'SFDC', `Last_SFDC_Sync__c` set.
   - SFDC Contact `Healthie_Speciality__c` formula = "Cardiology" (already automatic).

2. **Prospect HubSpot→SFDC**: HubSpot Contact (no SFDC Account or with a non-customer Account) submits a form setting `contact_specialty` = "Nutrition". Verify within 30s:
   - SFDC Account `Healthie_Specialty_v2__c` = "Nutrition" (mapped).
   - SFDC Contact `Healthie_Speciality__c` formula = "Nutrition".
   - `Sync_Source__c` = 'HubSpot', `Last_HubSpot_Sync__c` set.

3. **Lead path**: HubSpot Contact with no matching SFDC Account, but a matching SFDC Lead. Update HubSpot specialty. Verify Lead.Provider_Specialty__c is updated; no Account write. (Not applicable to us currently)

4. **Customer write rejection**: For an Account WITH Healthie_Organization, send a HubSpot→SFDC specialty update. Verify the write is rejected, an entry lands in `Sync_Error_Log__c` with reason "ignored: customer record", and the Account field is unchanged.

5. **Loop prevention**: Force-set `Sync_Source__c = 'HubSpot'` on an Account and update `Healthie_Specialty_v2__c` from a SFDC integration user. Verify the outbound flow does NOT fire (because the IS_CHANGED + classification gating should prevent it, OR if it does fire, the HubSpot endpoint should detect that HubSpot was the last-writer and no-op).

6. **Unknown value handling**: HubSpot form submits `contact_specialty` = "<value not in mapping table>". Verify SFDC writes to `Pending_Specialty_Review__c`, not to `Healthie_Specialty_v2__c`, and a Slack alert fires.

7. **Orphan Contact**: Backfill audit — confirm orphaned Contacts (no AccountId) are flagged in the reconciliation diff CSV. Sync does not attempt to write to them.

8. **Healthie Integration race**: Trigger a Healthie sync (via Healthie's own integration mechanism) that writes `Healthie_Specialty_v2__c`. Verify the outbound flow fires only once and `Sync_Source__c` reflects the actual chain.

---

## 9. Critical Files / Configuration Touchpoints

(For reference during execution; nothing edited until decisions in §3 are confirmed.)

- **Salesforce metadata** (eventual write targets — DO NOT modify yet):
  - `objects/Account.object` — new fields: `Sync_Source__c`, `Last_SFDC_Sync__c`, `Last_HubSpot_Sync__c`, `Pending_Specialty_Review__c`
  - `objects/Sync_Error_Log__c.object` — new custom object
  - `objects/Specialty_Value_Mapping__mdt.object` — new custom metadata type
  - `flows/Specialty_Customer_Outbound_Sync.flow` — new
  - `classes/HubSpotSpecialtyCallout.cls` — new
  - `classes/HubSpotSpecialtyWebhookReceiver.cls` — new (Apex REST endpoint)
  - `customPermissions/Specialty_Sync_Enabled.customPermission` — new
- **HubSpot config**:
  - Forms currently writing to `specialtyii`, `healthie_specialty`, `specialty`, `what_is_your_specialty` — need rewiring
  - Workflow `Specialty Prospect Outbound to SFDC` — new
  - Property settings on `contact_specialty` — value list expansion
- **Reused existing code**:
  - `SlackApiService` Apex class (from the in-house Slack integration) for `#data-quality-alerts` notifications
  - `Slack_Integration_Config__mdt.Default` config record
  - `SlackAPI` Remote Site Setting
  - `HubSpot_Inc__HubSpot_Id__c` field on Account/Contact for ID mapping
  - `generate_hubspot_dictionary.py` (local) for reconciliation extracts
  - `hubspot.config.yml` (local) for HubSpot API auth in scripts
- **Out of scope (explicitly)**:
  - `Account.Specialty__c` (legacy multipicklist) — deprecate separately
  - `Contact.Provider_Specialty__c` / HubSpot `provider_specialty` — different concept, different sync
  - Healthie product DB → SFDC sync — owned by the Healthie Integration team
  - Orphan Contact backfill — separate workstream

---

## 10. What Must Happen Before Build

1. **Confirm HubSpot canonical property** (§3.1) — Marketing Ops sign-off.
2. **Confirm customer/prospect classification rule** (§2) — Healthie_Organization presence as primary, RecordType as fallback. Sales Ops sign-off.
3. **Produce the value-mapping table** (§5, §7 Phase 0) — Data Ops + Marketing Ops jointly.
4. **Confirm Healthie Integration write cadence and field ownership** — Integration team sign-off; needed to size loop-prevention design.
5. **Confirm `provider_specialty` is out of scope** (§3.2) — explicit acknowledgement to avoid mid-build scope expansion.

Until all five are in writing, no SFDC metadata changes, no HubSpot property changes, no API code.
