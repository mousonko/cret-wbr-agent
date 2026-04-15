#!/bin/bash
# Quick setup and run script for the WBR Bridge Agent
set -e

echo "=== EU LM WBR CRET Bridge Agent ==="

# Install dependencies
if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "Installing dependencies..."
    pip3 install -r requirements.txt -q
fi

# Check for Excel file
if [ -z "$1" ]; then
    echo ""
    echo "Usage: ./run.sh <excel_file> [sheet_name]"
    echo ""
    echo "Examples:"
    echo "  ./run.sh 'DS Bridging Scan Compliance.xlsx' Wk-14"
    echo "  ./run.sh compliance.xlsx                          # auto-detects latest week sheet"
    echo ""
    echo "Options (via environment variables):"
    echo "  SLACK_WEBHOOK_URL=https://hooks.slack.com/...  ./run.sh file.xlsx Wk-14"
    echo "  SLACK_BOT_TOKEN=xoxb-...                       ./run.sh file.xlsx Wk-14"
    exit 1
fi

EXCEL="$1"
SHEET="${2:-}"
OUTDIR="./output"

mkdir -p "$OUTDIR"

# Run main analysis
SHEET_ARG=""
if [ -n "$SHEET" ]; then
    SHEET_ARG="--sheet $SHEET"
fi

echo ""
echo "Running analysis..."
python3 main.py "$EXCEL" $SHEET_ARG --offline --output-dir "$OUTDIR"

# Run WoW analysis if multiple week sheets exist
echo ""
echo "Running week-over-week analysis..."
python3 wow_analysis.py 2>/dev/null || echo "WoW analysis skipped (needs multi-week data)"

# Generate final report with charts
echo ""
echo "Generating final report with charts..."
python3 generate_final_report.py 2>/dev/null || python3 generate_enhanced_report.py 2>/dev/null || echo "Enhanced report generation skipped"

# Post to Slack if configured
if [ -n "$SLACK_WEBHOOK_URL" ] || [ -n "$SLACK_BOT_TOKEN" ]; then
    echo ""
    echo "Posting to Slack..."
    python3 slack_notify.py --output-dir "$OUTDIR"
fi

echo ""
echo "=== Done! ==="
echo "Reports in: $OUTDIR/"
ls -lh "$OUTDIR"/*.docx "$OUTDIR"/*.pptx 2>/dev/null
