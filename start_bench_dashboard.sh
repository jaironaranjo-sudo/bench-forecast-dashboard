#!/bin/bash
# ─────────────────────────────────────────────────────────────────
# NA Bench Forecast Dashboard — launcher
# Double-click (or run from terminal) to start the dashboard.
# All dependencies are self-contained inside this folder's .venv —
# completely isolated from every other app in this workspace.
# ─────────────────────────────────────────────────────────────────
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV="$SCRIPT_DIR/.venv"

echo ""
echo "  📊  NA Bench Forecast Dashboard"
echo "  ─────────────────────────────────────────────"

# Create the local venv if it doesn't exist yet
if [ ! -f "$VENV/bin/activate" ]; then
    echo "  ✓  Creating isolated virtual environment..."
    python3 -m venv "$VENV"
fi

echo "  ✓  Activating isolated virtual environment..."
source "$VENV/bin/activate"

echo "  ✓  Checking dependencies..."
pip install --quiet -r "$SCRIPT_DIR/requirements.txt"

echo "  ✓  Launching dashboard on http://localhost:8503"
echo ""

cd "$SCRIPT_DIR"
streamlit run bench_forecast_dashboard.py \
    --server.port 8503 \
    --server.headless false \
    --browser.gatherUsageStats false
