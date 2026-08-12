#!/bin/bash
# ─────────────────────────────────────────────────────────────────
# NA Bench Forecast — Push Data Update
# Double-click this file to commit & push the latest xlsx to GitHub.
# Make sure the NA Bench Forecast.xlsx file is CLOSED in Excel first.
# ─────────────────────────────────────────────────────────────────

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "  📤  NA Bench Forecast — Push Data Update"
echo "  ─────────────────────────────────────────────"
echo ""

# Check for Excel lock file (file is still open)
if [ -f "~\$NA Bench Forecast.xlsx" ]; then
    echo "  ⚠️  ERROR: The file appears to be open in Excel."
    echo "       Please close it and double-click this script again."
    echo ""
    read -p "  Press Enter to close..."
    exit 1
fi

# Check the file exists
if [ ! -f "NA Bench Forecast.xlsx" ]; then
    echo "  ⚠️  ERROR: NA Bench Forecast.xlsx not found in this folder."
    echo ""
    read -p "  Press Enter to close..."
    exit 1
fi

# Stage the xlsx (and overrides if any)
git add "NA Bench Forecast.xlsx"
git add "bench_overrides.json" 2>/dev/null || true   # optional, may not exist

# Check if there's actually anything staged
if git diff --cached --quiet; then
    echo "  ✓  No changes detected — nothing to push."
    echo ""
    read -p "  Press Enter to close..."
    exit 0
fi

# Commit with a timestamp
TIMESTAMP=$(date "+%Y-%m-%d %H:%M")
git commit -m "data: update NA Bench Forecast — $TIMESTAMP"

echo ""
echo "  ✓  Committed. Pushing to GitHub..."
echo ""

git push origin main

echo ""
echo "  ✅  Done! Changes are live on Streamlit Cloud."
echo ""
read -p "  Press Enter to close..."
