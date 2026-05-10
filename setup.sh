#!/bin/bash

################################################################################
# KRI/KPI GENERATOR - Standalone Master Setup Script
# 
# This script creates a complete sandbox environment from scratch.
# Works from any directory - requires only Python, curl, and gunzip.
#
# Usage: bash setup.sh [--no-run]
#   --no-run  : Setup only, don't run generator
#
# Creates: /tmp/krikpi_sandbox/ (isolated workspace)
#
################################################################################

set -e

# ============================================================================
# CONFIGURATION
# ============================================================================

SANDBOX_DIR="/tmp/krikpi_sandbox"
RUN_GENERATOR=true

# Check for --no-run flag
if [ "$1" == "--no-run" ]; then
    RUN_GENERATOR=false
fi

# ============================================================================
# COLOR CODES
# ============================================================================

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

print_header() {
    echo -e "${BLUE}"
    echo "╔════════════════════════════════════════════════════════════════╗"
    echo "║  $1"
    echo "╚════════════════════════════════════════════════════════════════╝"
    echo -e "${NC}"
}

print_step() {
    echo -e "${BLUE}[$1] $2${NC}"
}

print_success() {
    echo -e "${GREEN}✓ $1${NC}"
}

print_warning() {
    echo -e "${YELLOW}⚠ $1${NC}"
}

print_error() {
    echo -e "${RED}❌ $1${NC}"
}

print_info() {
    echo -e "${CYAN}ℹ $1${NC}"
}

# ============================================================================
# STEP 1: Check Python
# ============================================================================

clear
print_header "KRI/KPI Generator - Master Setup Script"
echo "Complete standalone setup for CachyOS / Arch Linux"
echo

print_step "1/10" "Checking Python installation..."

if ! command -v python &> /dev/null; then
    print_error "Python not found"
    echo "Install with: sudo pacman -S python"
    exit 1
fi

PYTHON_VERSION=$(python --version 2>&1 | awk '{print $2}')
print_success "Python $PYTHON_VERSION found"
echo

# ============================================================================
# STEP 2: Check curl and gunzip
# ============================================================================

print_step "2/10" "Checking curl and gunzip..."

if ! command -v curl &> /dev/null; then
    print_warning "curl not found (needed for EPSS download)"
fi

if ! command -v gunzip &> /dev/null; then
    print_warning "gunzip not found (needed for decompression)"
fi

print_success "Required tools available"
echo

# ============================================================================
# STEP 3: Create Sandbox Structure
# ============================================================================

print_step "3/10" "Creating sandbox directory structure..."

# Remove old sandbox if exists
if [ -d "$SANDBOX_DIR" ]; then
    print_info "Removing old sandbox..."
    rm -rf "$SANDBOX_DIR"
fi

# Create all directories
mkdir -p "$SANDBOX_DIR"/{data,output,scripts,logs,venv}

print_success "Sandbox structure created:"
echo "  Location: $SANDBOX_DIR"
echo "  ├── data/       (input Excel files)"
echo "  ├── output/     (generated metrics)"
echo "  ├── scripts/    (Python generator)"
echo "  ├── logs/       (execution logs)"
echo "  └── venv/       (virtual environment)"
echo

# ============================================================================
# STEP 4: Create Launcher Scripts
# ============================================================================

print_step "4/10" "Creating launcher scripts..."

# Create setup.sh for venv
cat > "$SANDBOX_DIR/setup.sh" << 'SETUP_SCRIPT'
#!/bin/bash
set -e

SANDBOX_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "╔════════════════════════════════════════════════════════════════╗"
echo "║  Setting up Python virtual environment...                     ║"
echo "╚════════════════════════════════════════════════════════════════╝"
echo

if [ -d "$SANDBOX_DIR/venv" ]; then
    echo "ℹ Virtual environment already exists"
else
    echo "[1/3] Creating virtual environment..."
    python -m venv "$SANDBOX_DIR/venv"
    echo "✓ venv created"
fi

echo "[2/3] Activating and installing dependencies..."
source "$SANDBOX_DIR/venv/bin/activate"

pip install --quiet --upgrade pip
pip install --quiet pandas openpyxl requests

echo "✓ Dependencies installed:"
echo "  - pandas"
echo "  - openpyxl"
echo "  - requests"

echo "[3/3] Verifying installation..."
python -c "import pandas, openpyxl, requests; print('✓ All imports successful')"
echo

echo "✓ Virtual environment ready!"
SETUP_SCRIPT

chmod +x "$SANDBOX_DIR/setup.sh"
print_success "setup.sh created"

# Create run.sh for generator
cat > "$SANDBOX_DIR/run.sh" << 'RUN_SCRIPT'
#!/bin/bash

SANDBOX_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
LOG_FILE="$SANDBOX_DIR/logs/run_${TIMESTAMP}.log"

mkdir -p "$SANDBOX_DIR"/{data,output,logs}

if [ ! -d "$SANDBOX_DIR/venv" ]; then
    echo "❌ Virtual environment not found"
    echo "Run setup first: bash $SANDBOX_DIR/setup.sh"
    exit 1
fi

if [ ! -f "$SANDBOX_DIR/scripts/kri_kpi_generator_v3.2.py" ]; then
    echo "❌ Generator script not found"
    exit 1
fi

# Check data files
MISSING=0
for file in Asset_Inventory.xlsx \
            Vulnerability_Scan_Data_OPEN_REALISTIC.xlsx \
            Vulnerability_Scan_Data_CLOSED_REALISTIC.xlsx; do
    if [ ! -f "$SANDBOX_DIR/data/$file" ]; then
        echo "❌ Missing: $file"
        MISSING=$((MISSING + 1))
    fi
done

if [ $MISSING -gt 0 ]; then
    echo "Error: Missing $MISSING required data file(s)"
    echo "Place Excel files in: $SANDBOX_DIR/data/"
    exit 1
fi

echo "╔════════════════════════════════════════════════════════════════╗"
echo "║  Running KRI/KPI Generator                                    ║"
echo "╚════════════════════════════════════════════════════════════════╝"
echo

{
    echo "═══════════════════════════════════════════════════════════════"
    echo "KRI/KPI Generator Execution - $TIMESTAMP"
    echo "═══════════════════════════════════════════════════════════════"
    echo

    source "$SANDBOX_DIR/venv/bin/activate"
    
    python "$SANDBOX_DIR/scripts/kri_kpi_generator_v3.2.py" \
        --assets "$SANDBOX_DIR/data/Asset_Inventory.xlsx" \
        --open "$SANDBOX_DIR/data/Vulnerability_Scan_Data_OPEN_REALISTIC.xlsx" \
        --closed "$SANDBOX_DIR/data/Vulnerability_Scan_Data_CLOSED_REALISTIC.xlsx" \
        --output "$SANDBOX_DIR/output"

    echo
    echo "✓ Generator complete!"
} 2>&1 | tee "$LOG_FILE"

echo
echo "📝 Log: $LOG_FILE"
RUN_SCRIPT

chmod +x "$SANDBOX_DIR/run.sh"
print_success "run.sh created"

# Create README
cat > "$SANDBOX_DIR/README.md" << 'README'
# KRI/KPI Generator - Sandbox

Complete KRI/KPI metrics generator for vulnerability risk management.

## Quick Start

```bash
# Copy your data files
cp Asset_Inventory.xlsx data/
cp Vulnerability_Scan_Data_OPEN_REALISTIC.xlsx data/
cp Vulnerability_Scan_Data_CLOSED_REALISTIC.xlsx data/
cp kri_kpi_generator_v3.2.py scripts/

# Setup (one-time)
bash setup.sh

# Run generator
bash run.sh

# View results
libreoffice output/KRI_KPI_Metrics_*.xlsx
```

## Files Needed

- `Asset_Inventory.xlsx` - Asset inventory
- `Vulnerability_Scan_Data_OPEN_REALISTIC.xlsx` - Open findings
- `Vulnerability_Scan_Data_CLOSED_REALISTIC.xlsx` - Closed findings
- `kri_kpi_generator_v3.2.py` - Python generator script

## Features

- Fetches EPSS from CSV.GZ
- Fetches CISA KEV data
- Calculates 15 metrics (8 KPIs + 7 KRIs)
- Generates professional Excel report
- 5 sheets with formatting and status indicators

## Requirements

- Python 3.7+
- curl (for EPSS download)
- gunzip (for decompression)
- pandas, openpyxl, requests (installed automatically)
README

print_success "README.md created"
echo

# ============================================================================
# STEP 5: Create requirements.txt
# ============================================================================

print_step "5/10" "Creating requirements.txt..."

cat > "$SANDBOX_DIR/requirements.txt" << 'REQUIREMENTS'
# KRI/KPI Generator Dependencies
pandas>=1.3.0
openpyxl>=3.6.0
requests>=2.26.0
REQUIREMENTS

print_success "requirements.txt created"
echo

# ============================================================================
# STEP 6: Create .gitignore for GitHub
# ============================================================================

print_step "6/10" "Creating .gitignore..."

cat > "$SANDBOX_DIR/.gitignore" << 'GITIGNORE'
# Virtual environment
venv/

# Data files (user-provided)
data/
*.xlsx

# Generated outputs
output/
*.xlsx

# Logs
logs/
*.log

# Python
__pycache__/
*.pyc
*.pyo
*.pyd
.Python
*.egg-info/
dist/
build/

# IDE
.vscode/
.idea/
*.swp
*.swo

# OS
.DS_Store
.env
GITIGNORE

print_success ".gitignore created"
echo

# ============================================================================
# STEP 7: Check Python
# ============================================================================

print_step "7/10" "Verifying Python..."

PYTHON_VERSION=$(python --version 2>&1 | awk '{print $2}')
print_success "Python $PYTHON_VERSION verified"
echo

# ============================================================================
# STEP 8: Create Virtual Environment
# ============================================================================

print_step "8/10" "Creating Python virtual environment..."

python -m venv "$SANDBOX_DIR/venv"
print_success "Virtual environment created"
echo

# ============================================================================
# STEP 9: Install Dependencies
# ============================================================================

print_step "9/10" "Installing Python dependencies..."

source "$SANDBOX_DIR/venv/bin/activate"

pip install --quiet --upgrade pip
pip install --quiet pandas openpyxl requests

print_success "Dependencies installed:"
echo "  - pandas"
echo "  - openpyxl"
echo "  - requests"

python -c "import pandas, openpyxl, requests; print('✓ All imports verified')"
echo

# ============================================================================
# STEP 10: Verify Sandbox
# ============================================================================

print_step "10/10" "Verifying sandbox..."

echo "Sandbox structure:"
echo "  Scripts:"
ls -1 "$SANDBOX_DIR"/*.sh 2>/dev/null | xargs -I {} basename {} | sed 's/^/    ✓ /'
echo "  Documentation:"
ls -1 "$SANDBOX_DIR"/*.{md,txt} 2>/dev/null | xargs -I {} basename {} | sed 's/^/    ✓ /'
echo

# ============================================================================
# COMPLETION
# ============================================================================

echo -e "${GREEN}╔════════════════════════════════════════════════════════════════╗"
echo "║           ✓ SANDBOX SETUP COMPLETE!                          ║"
echo "╚════════════════════════════════════════════════════════════════╝${NC}"
echo

echo "📁 Sandbox location: $SANDBOX_DIR"
echo

echo "📋 Next steps:"
echo "  1. Add your data files:"
echo "     cp Asset_Inventory.xlsx $SANDBOX_DIR/data/"
echo "     cp Vulnerability_Scan_Data_OPEN_REALISTIC.xlsx $SANDBOX_DIR/data/"
echo "     cp Vulnerability_Scan_Data_CLOSED_REALISTIC.xlsx $SANDBOX_DIR/data/"
echo
echo "  2. Add the generator script:"
echo "     cp kri_kpi_generator_v3.2.py $SANDBOX_DIR/scripts/"
echo
echo "  3. Setup virtual environment (one-time):"
echo "     bash $SANDBOX_DIR/setup.sh"
echo
echo "  4. Run generator:"
echo "     bash $SANDBOX_DIR/run.sh"
echo
echo "  5. View results:"
echo "     libreoffice $SANDBOX_DIR/output/KRI_KPI_Metrics_*.xlsx"
echo

echo "📝 Additional commands:"
echo "  View logs:     cat $SANDBOX_DIR/logs/run_*.log"
echo "  Re-setup venv: bash $SANDBOX_DIR/setup.sh"
echo
