#!/bin/bash
# ─────────────────────────────────────────────────────────────────
# NA Bench Forecast Dashboard — launcher
# ─────────────────────────────────────────────────────────────────
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(dirname "$SCRIPT_DIR")"
VENV="$REPO_ROOT/.venv"

echo ""
echo "  📊  NA Bench Forecast Dashboard"
echo "  ─────────────────────────────────────────────"

# Activate virtual environment if present
if [ -f "$VENV/bin/activate" ]; then
    echo "  ✓  Activating virtual environment..."
    source "$VENV/bin/activate"
fi

# Install / verify dependencies
echo "  ✓  Checking dependencies..."
pip install --quiet streamlit plotly pandas openpyxl

echo "  ✓  Launching dashboard on http://localhost:8503"
echo ""

cd "$SCRIPT_DIR"
streamlit run bench_forecast_dashboard.py \
    --server.port 8503 \
    --server.headless false \
    --browser.gatherUsageStats false
