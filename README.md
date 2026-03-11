# INS_delphi

Delphi W1 data collection and reporting system for malaria intervention optimization.

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

## Workflow

1. **Before data collection:** Run `generate_qrcode.py` to create QR code for expert access
2. **During data collection:** Run `dashboard.py` to monitor progress in real-time
3. **Generate results:** Run `aggregate_results.py` to create archival results file
4. **Create report:** Run `generate_w1_report.py` on results file to generate formatted HTML

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

