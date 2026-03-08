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
