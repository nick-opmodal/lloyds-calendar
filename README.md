# Lloyd's Deadlines — Calendar Feed for Submission Markets

Subscribable ICS calendar of Lloyd's Business Timetable deadlines with
opmodal editorial commentary. Two outputs:

- **Full feed** (`lloyds-deadlines.ics`): every deadline in the rolling
  8-week window. Subscribers get the complete timetable in their calendar.
- **Curated feed** (`lloyds-deadlines-curated.ics`): only deadlines with
  opmodal commentary attached. ~17 events at any time, high-signal.

## How it works

```
Lloyd's XLSX export ──→ generate_lloyds_calendar.py ──→ .ics file
(lloyds.com, weekly)     (download + parse + annotate)   (Outlook/Apple Calendar)
                                     ↑
                              commentary.yaml
                         (editorial rules: regex → "why it matters")
```

## Files

| File | Purpose |
|------|---------|
| `generate_lloyds_calendar.py` | Downloads Lloyd's timetable XLSX, parses deadlines, attaches commentary, emits RFC 5545 ICS |
| `commentary.yaml` | Editorial layer — regex rules matching deadlines to opmodal commentary and article links |
| `regenerate.sh` | Wrapper script for cron/CI — runs generator and reports stats |
| `newsletter-fortnight-ahead.html` | Beehiiv-ready HTML block for the "Fortnight Ahead" section in Submission Markets |
| `lloyds-deadlines.ics` | Generated full feed (71 events as of 2026-07-28) |
| `lloyds-deadlines-curated.ics` | Generated curated feed (only events with commentary, ~17 events) |

## Adding commentary

Edit `commentary.yaml`. Each rule:

```yaml
- match: "QMB"              # case-insensitive regex
  comment: >-               # "Why it matters" text
    The Quarterly Monitoring Return is the performance data Lloyd's uses
    to track plan versus actual.
  month: "2026-09"          # optional — only apply in this month
```

First matching rule wins. Put specific rules above general ones.

## Weekly regeneration

Monday 6am via Hermes cron (job `e3713d36cff8`). Runs `regenerate.sh` 
which calls the generator against the live Lloyd's XLSX.

Manual run:
```bash
cd ~/Projects/ElCapo/skunkworks/lloyds-calendar
~/.hermes/hermes-agent/venv/bin/python generate_lloyds_calendar.py
```

## Hosting

The `.ics` file needs a public URL for subscribers. Options:
- GitHub Pages (free, zero infra)
- VM nginx at `calendar.opmodal.com`
- Wix file hosting (if supported)

Update the `href` in `newsletter-fortnight-ahead.html` when the URL is chosen.

## Subscriber instructions

**Outlook:** File → Open & Export → Import/Export → Import an iCalendar (.ics)
  → Open as New. Outlook will refresh automatically.

**Apple Calendar:** File → New Calendar Subscription → paste the URL.
  Auto-refresh interval is set by the `X-PUBLISHED-TTL:PT12H` header.

**Google Calendar:** Settings → Add Calendar → From URL → paste the `.ics` URL.
  Google polls roughly every 8 hours.

## Notes

- Commentary is Opmodal's own; deadline data remains Lloyd's. The calendar
  description includes the disclaimer: "Always verify against the official timetable."
- UIDs are stable (date + title hash) — regeneration updates events in place,
  no duplicates in subscribers' calendars.
- Check Lloyd's terms of use regarding automated XLSX retrieval.
