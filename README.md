# INS_delphi

Delphi W1 data collection and reporting system for malaria/SMI/HIV/TB intervention optimization.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure `config.env`:
   - Set `KOBO_TOKEN` with your KoboToolbox API token
   - Set `KOBO_SERVER` and `FORM_ID`
   - Update `TOPIC_CODE` and `INPUT_FILE` as needed

## Scripts

### aggregate_results.py

Fetch submissions from KoboToolbox API and generate aggregated results file.

**Default behavior** (API fetch):
```bash
python code/aggregate_results.py
```

**CSV mode** (read from CSV file):
```bash
python code/aggregate_results.py --csv path/to/file.csv
```

**Outputs:**
- Excel file in `./results/` with timestamp
- Includes: Submissions, Responses (wide), QC_Summary, QC_Detail, Coverage, Coverage_detail, Catalogo (from dictionary)

### generate_w1_report.py

Generate HTML report from aggregated results.

**Usage:**
```bash
python code/generate_w1_report.py path/to/results.xlsx
```

**Optional arguments:**
- `--output-dir DIR` - Override output directory (default: `./reports/`)

**Outputs:**
- HTML file in `./reports/` with timestamp
- Includes collapsible response rate details and intervention metadata

### dashboard.py

Real-time Streamlit dashboard for monitoring expert response completion.

**Usage:**
```bash
streamlit run code/dashboard.py
```

**Features:**
- Auto-refresh (5-minute cache)
- Expert × intervention heatmap
- Expert × group coverage matrix  
- Response rate charts (by intervention and expert)
- Summary metrics

### generate_qrcode.py

Generate QR code for the gateway URL to share with experts.

**Usage:**
```bash
python code/generate_qrcode.py
```

**Optional arguments:**
- `--url URL` - Custom URL to encode (default: gateway URL)
- `--output FILE` - Output file path (default: `gateway_qr.png`)
- `--box-size N` - Size of each box in pixels (default: 10)
- `--border N` - Border size in boxes (default: 4)

**Example:**
```bash
python code/generate_qrcode.py --output qr_codes/malaria_gateway.png --box-size 15
```

### generate_kobo_and_pages.py

Generate Kobo XLSForms and intervention HTML pages from the dictionary workbook.

**Usage:**
```bash
python code/generate_kobo_and_pages.py
```

**Grouping configuration (`config.env`):**
- `SUBFORM_GROUP_BY` must be one of: `area`, `programa`, `componente`, `grupo`
- `SUBFORM_MAX_SIZE` optionally splits large groups into chunks

**Validation behavior:**
- Script fails fast if `SUBFORM_GROUP_BY` has an invalid value
- Script fails if any row has blank value in the selected grouping column
   (example: if `SUBFORM_GROUP_BY=grupo`, all `Grupo` cells must be filled)

### deploy_kobo_forms.py

Upload/redeploy generated sub-form XLSForms to KoboToolbox and persist
`SUBFORM_URL_*` + `SUBFORM_ASSET_*` mappings in `deployed_forms.env`.

**Usage:**
```bash
python code/deploy_kobo_forms.py            # fresh upload/create assets
python code/deploy_kobo_forms.py --redeploy # update existing mapped assets
python code/deploy_kobo_forms.py --list     # list account assets
```

**Important behavior:**
- `--redeploy` expects stable slug-to-asset mapping (`SUBFORM_ASSET_*`)
- after deploy, script rewrites `deployed_forms.env` and regenerates pages
- this is the preferred path when grouping/slugs are unchanged

## Workflow

1. **Before data collection:** Run `generate_qrcode.py` to create QR code for expert access
2. **During data collection:** Run `dashboard.py` to monitor progress in real-time
3. **Generate results:** Run `aggregate_results.py` to create archival results file
4. **Create report:** Run `generate_w1_report.py` on results file to generate formatted HTML

## Form deployment reset modes

Use one of the two reset modes below depending on whether form grouping changed.

### A) Soft reset (preferred)

Use when grouping output is unchanged (same subform slugs / same number of forms):

```bash
python code/generate_dictionaries.py --all
python code/generate_kobo_and_pages.py
python code/deploy_kobo_forms.py --redeploy
```

This preserves asset UIDs and updates form content in place.

### B) Hard reset (recommended when regrouping changes assets)

Use when you changed `SUBFORM_GROUP_BY`, `SUBFORM_MAX_SIZE`, dictionary structure,
or otherwise produced a different slug set (more/fewer/renamed subforms).

1. Archive or export any submissions you need to keep.
2. In KoboToolbox, delete old W1 subform assets for the topic.
3. Remove `deployed_forms.env` (or clear all `SUBFORM_URL_*` and `SUBFORM_ASSET_*` lines).
4. Regenerate files:
   ```bash
   python code/generate_dictionaries.py --all
   python code/generate_kobo_and_pages.py
   ```
5. Deploy fresh assets (without `--redeploy`):
   ```bash
   python code/deploy_kobo_forms.py
   ```
6. If using Streamlit Cloud secrets, regenerate and paste secrets:
   ```bash
   python code/make_secrets_toml.py
   ```

Why hard reset in this case: old asset mappings can become stale when slug names
or counts change, which can leave orphaned forms or mismatched URLs.

## Pending issues

ISSUE: Expert codes misinterpreted as times on Portuguese-locale browsers
Status: Open
Priority: High — must fix before workshop
Affected component: KoboToolbox Enketo form — Código do Especialista dropdown (Q01)
Description:
On machines with a Portuguese locale (pt-PT or pt-MZ), the expert codes in the dropdown are rendered as times rather than strings. The browser's locale-aware parser interprets the leading digits + PM suffix as a time expression. Examples observed: 001PM → 00:01, 002PM → 14h00, 005PM → 00h05, 006PM → 18h00. The XX-suffix codes appear unaffected. Only one machine confirmed affected so far but likely affects any pt-locale browser.
Root cause:
The code format [0-9]{3}PM is ambiguous — leading zeros + PM suffix is a valid time expression in Portuguese locale time parsing. The issue originates in the XLSForm choices list where these values are not explicitly typed as strings.
Recommended fix:
Prefix all expert codes with a letter to make them unambiguously non-numeric. Proposed new format: E[0-9]{3}PM and E[0-9]{3}XX (e.g. E001PM, E011XX).
Files requiring changes:

deploy_kobo_forms.py — choices list generation
generate_kobo_and_pages.py — EXPERT_CODES list and any regex validation
gateway.html — ALLOWED_HASHES (all hashes must be regenerated for new code format)
aggregate_results.py — EXPERT_CODES list, regex validation ^[0-9]{3}(PM|XX)$
delphi_w1_malaria_template.xlsx — Listas sheet codes and dropdown validation
dicionario_delphi_w1_malaria.xlsx — codes list if maintained there
Any already-submitted responses — codes used so far will need to be noted and remapped if resubmitting

Note: Codes already submitted to Kobo under the old format should be
documented before redeployment. If the workshop has not yet started,
clean redeployment is straightforward. If partial submissions exist,
the aggregator's expert_code field will need a remap step.

