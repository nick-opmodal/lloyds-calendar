#!/bin/bash
# Weekly Lloyds Deadline Calendar regeneration
# Run by Hermes cron: Mondays at 6am
#
# 1. Downloads the latest Lloyd's XLSX, generates ICS feeds
# 2. Commits and pushes to GitHub Pages (calendar.submissionmarkets.news)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_PYTHON="$HOME/.hermes/hermes-agent/venv/bin/python"

# Full feed (all deadlines)
"$VENV_PYTHON" "$SCRIPT_DIR/generate_lloyds_calendar.py" -o "$SCRIPT_DIR/lloyds-deadlines.ics" 2>&1

# Curated feed (commentary only)
"$VENV_PYTHON" "$SCRIPT_DIR/generate_lloyds_calendar.py" --curated-only -o "$SCRIPT_DIR/lloyds-deadlines-curated.ics" 2>&1

FULL_COUNT=$(grep -c 'BEGIN:VEVENT' "$SCRIPT_DIR/lloyds-deadlines.ics")
CURATED_COUNT=$(grep -c 'BEGIN:VEVENT' "$SCRIPT_DIR/lloyds-deadlines-curated.ics")

echo ""
echo "Calendar regenerated: $(date '+%Y-%m-%d %H:%M')"
echo "Full feed:    $FULL_COUNT events ($(stat -f%z "$SCRIPT_DIR/lloyds-deadlines.ics") bytes)"
echo "Curated feed: $CURATED_COUNT events ($(stat -f%z "$SCRIPT_DIR/lloyds-deadlines-curated.ics") bytes)"

# Push to GitHub Pages
cd "$SCRIPT_DIR"
git add lloyds-deadlines.ics lloyds-deadlines-curated.ics
if git diff --cached --quiet; then
    echo "No changes to push"
else
    git commit -m "Weekly regeneration: $(date '+%Y-%m-%d') — $FULL_COUNT events ($CURATED_COUNT curated)"
    git push
    echo "Pushed to GitHub Pages → https://calendar.submissionmarkets.news/"
fi
