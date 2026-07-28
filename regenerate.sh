#!/bin/bash
# Weekly Lloyds Deadline Calendar regeneration
# Run by Hermes cron: Mondays at 6am

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_PYTHON="$HOME/.hermes/hermes-agent/venv/bin/python"

# Full feed (all deadlines)
"$VENV_PYTHON" "$SCRIPT_DIR/generate_lloyds_calendar.py" -o "$SCRIPT_DIR/lloyds-deadlines.ics" 2>&1

# Curated feed (commentary only)
"$VENV_PYTHON" "$SCRIPT_DIR/generate_lloyds_calendar.py" --curated-only -o "$SCRIPT_DIR/lloyds-deadlines-curated.ics" 2>&1

echo ""
echo "Calendar regenerated: $(date '+%Y-%m-%d %H:%M')"
echo "Full feed:    $(grep -c 'BEGIN:VEVENT' "$SCRIPT_DIR/lloyds-deadlines.ics") events ($(stat -f%z "$SCRIPT_DIR/lloyds-deadlines.ics") bytes)"
echo "Curated feed: $(grep -c 'BEGIN:VEVENT' "$SCRIPT_DIR/lloyds-deadlines-curated.ics") events ($(stat -f%z "$SCRIPT_DIR/lloyds-deadlines-curated.ics") bytes)"
