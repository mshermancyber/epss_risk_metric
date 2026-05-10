# epss_risk_metric
KRI/KPI Vulnerability Risk Management metrics generator.  Calculates 15 metrics (8 KPIs + 7 KRIs) from vulnerability scan data.  Integrates EPSS CSV.GZ and CISA KEV feeds.  Generates professional Excel reports with status indicators. One-command setup, no hardcoded paths.

## Features

- **15 Metrics**: 8 Key Performance Indicators + 7 Key Risk Indicators
- **Live Threat Intelligence**: EPSS from CSV.GZ + CISA Known Exploited Vulnerabilities
- **Professional Reporting**: 5-sheet Excel output with formatting and status indicators
- **Automatic Setup**: One-command sandbox creation, works from any directory
- **No Hardcoded Paths**: Completely portable, ideal for sharing and collaboration
- **Error Handling**: Graceful fallbacks if external feeds fail

## Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/mshermancyber/kri-kpi-generator
cd kri-kpi-generator
```

### 2. Run Setup (creates `/tmp/krikpi_sandbox`)
```bash
bash setup.sh
```

### 3. Add Your Data Files
```bash
cp /path/to/Asset_Inventory.xlsx /tmp/krikpi_sandbox/data/
cp /path/to/Vulnerability_OPEN.xlsx /tmp/krikpi_sandbox/data/
cp /path/to/Vulnerability_CLOSED.xlsx /tmp/krikpi_sandbox/data/
cp kri_kpi_generator_v3.2.py /tmp/krikpi_sandbox/scripts/
```

### 4. Setup Virtual Environment (one-time)
```bash
bash /tmp/krikpi_sandbox/setup.sh
```

### 5. Run Generator
```bash
bash /tmp/krikpi_sandbox/run.sh
```

### 6. View Results
```bash
libreoffice /tmp/krikpi_sandbox/output/KRI_KPI_Metrics_*.xlsx
```

## Requirements

- Python 3.7+
- curl (for EPSS download)
- gunzip (for decompression)
- Dependencies: pandas, openpyxl, requests (installed automatically)

### Install Dependencies (CachyOS/Arch)
```bash
sudo pacman -S python curl gzip
```

### Install Dependencies (Ubuntu/Debian)
```bash
sudo apt-get install python3 curl gzip
```

## Input Files

Your vulnerability scan data must include three Excel files:

### 1. Asset_Inventory.xlsx
- **Columns**: Asset ID, Hostname, Asset Type, Environment, OS, Location, IP Address, AWS Account Name, Internet Facing, Data Classification, Business Criticality, Last Patched, Owner
- **Business Criticality**: Tier 1, Tier 2, or Tier 3

### 2. Vulnerability_Scan_Data_OPEN_REALISTIC.xlsx
- **Columns**: Asset ID, CVE, Severity, Description, Discovery Date, EPSS, Scanner, Last Scanned
- **Status**: Open vulnerabilities awaiting remediation

### 3. Vulnerability_Scan_Data_CLOSED_REALISTIC.xlsx
- **Columns**: Asset ID, CVE, Severity, Description, Discovery Date, Remediation Date, EPSS, Scanner, Last Scanned
- **Status**: Remediated vulnerabilities with closure dates

See `examples/` folder for format reference.

## Metrics

### Key Performance Indicators (KPIs)
1. **MTTR by Severity** - Mean time to remediate (days)
2. **SLA Compliance** - Percentage of findings within SLA targets
3. **Unpatched Critical/High** - Count of unpatched critical/high severity findings
4. **EPSS Distribution** - Vulnerability count by EPSS score range
5. **Monthly Discovery Rate** - New findings per month
6. **Patch Lag (KEV)** - Days since CISA exploitation announced
7. **False Positive Rate** - Percentage of false positives
8. **Detection Coverage** - Findings by scan type

### Key Risk Indicators (KRIs)
1. **Unmitigated Critical Risk Score** - Weighted risk from unpatched critical findings
2. **Total Days at Risk (DAR)** - Cumulative risk-adjusted exposure time
3. **Exploit-in-the-Wild Count** - Number of CISA KEV findings
4. **Control Effectiveness Gap** - Percentage of unmitigated findings
5. **Regression Rate** - Percentage of re-opened vulnerabilities
6. **Asset Risk Concentration** - Risk concentration among top assets
7. **Technical Debt** - Count of findings >120 days old

## Excel Output

### Sheet 1: Metrics Summary
- All 15 metrics with values
- Status indicators (🟢 OK, 🟡 WARN, 🔴 CRITICAL)
- Risk thresholds applied

### Sheet 2: Severity Breakdown
- Vulnerability count by severity (Critical, High, Medium, Low, Info)
- MTTR and SLA compliance by severity

### Sheet 3: KEV Analysis
- CISA Known Exploited Vulnerabilities count
- Coverage gap percentage
- Exploit lag in days

### Sheet 4: Coverage Analysis
- Detection coverage by scanner type
- Count and percentage for each scanner

### Sheet 5: (Optional) Executive Summary
- Critical findings summary
- Key risk metrics
- Trend indicators

## Data Sources

### EPSS (Exploit Prediction Scoring System)
- **Source**: https://epss.empiricalsecurity.com/epss_scores-current.csv.gz
- **Method**: Curl download + gunzip decompression
- **Coverage**: ~200,000+ CVEs with daily updates
- **Fallback**: Uses scan data EPSS if external fetch fails

### CISA KEV (Known Exploited Vulnerabilities)
- **Source**: CISA official feeds (3 sources with fallback)
- **Coverage**: Active exploitations in the wild
- **Updates**: Real-time when available

## Architecture

```
setup.sh (main setup script)
  ├── Creates /tmp/krikpi_sandbox/ directory structure
  ├── Creates launcher scripts (setup.sh, run.sh)
  ├── Sets up Python virtual environment
  └── Installs dependencies

run.sh (generator launcher)
  ├── Verifies venv and scripts
  ├── Validates input files
  └── Executes kri_kpi_generator_v3.2.py

kri_kpi_generator_v3.2.py (main generator)
  ├── Fetches EPSS from CSV.GZ
  ├── Fetches CISA KEV data
  ├── Loads asset inventory
  ├── Loads vulnerability findings
  ├── Calculates 15 metrics
  └── Generates Excel report
```

## Troubleshooting

### Setup Issues

**Error: Python not found**
```bash
# Install Python
sudo pacman -S python  # CachyOS/Arch
sudo apt-get install python3  # Ubuntu/Debian
```

**Error: curl or gunzip not found**
```bash
# Install required tools
sudo pacman -S curl gzip  # CachyOS/Arch
sudo apt-get install curl gzip  # Ubuntu/Debian
```

### Runtime Issues

**EPSS fetch fails (network blocked)**
- Generator falls back to scan data EPSS automatically
- Metrics still generated with local EPSS values
- Check logs: `cat /tmp/krikpi_sandbox/logs/run_*.log`

**KEV fetch fails**
- All 3 fallback sources attempted
- Generator continues without KEV data
- Exploit-in-the-Wild count will be 0

**Missing input files**
- Ensure filenames match exactly (case-sensitive)
- Check files are in `/tmp/krikpi_sandbox/data/`
- Verify Excel format matches requirements

## Logs

Execution logs saved to: `/tmp/krikpi_sandbox/logs/run_TIMESTAMP.log`

View latest log:
```bash
cat /tmp/krikpi_sandbox/logs/run_*.log | tail -50
```

## Examples

Sample Excel files included in `examples/` folder:
- `Asset_Inventory_EXAMPLE.xlsx` - 50 assets
- `Vulnerability_OPEN_EXAMPLE.xlsx` - 600 open findings
- `Vulnerability_CLOSED_EXAMPLE.xlsx` - 1000 closed findings

Use these as templates for your data.

## License

MIT

## Author

Mark Sherman (@mshermancyber) - Cyber Security Specialist, 19+ years in financial services risk governance

## Contributing

Pull requests welcome. For major changes, open an issue first.

## Support

- Issues: GitHub Issues
- Documentation: See `docs/` folder
- Example data: See `examples/` folder
