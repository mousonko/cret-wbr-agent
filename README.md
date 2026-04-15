# EU LM WBR — CRET Scan Compliance Bridge Agent

AI-powered agent that analyzes the DS Bridging Scan Compliance Excel, identifies top 5 root causes, generates WBR-ready reports with charts, and tracks week-over-week trends.

## Quick Start

```bash
# 1. Clone the repo
git clone <repo-url>
cd wbr-bridge-agent

# 2. Install dependencies (one time)
pip3 install -r requirements.txt

# 3. Download the Excel from SharePoint
# https://amazon-my.sharepoint.com/...DS%20Bridging%20Scan%20Compliance.xlsx

# 4. Run
./run.sh "DS Bridging Scan Compliance.xlsx" Wk-14
```

## Output

The agent generates in `./output/`:

| File | Description |
|------|-------------|
| `WBR_Bridge_Analysis_FINAL.docx` | Word report with all charts, WoW trends, root causes |
| `WBR_Bridge_Analysis_FINAL.pptx` | 9-slide PowerPoint deck ready for WBR |
| `wbr_analysis.json` | Raw structured analysis (JSON) |
| `wow_analysis.json` | Week-over-week trending data |
| `chart_*.png` | Individual chart images |

## Report Contents

- **Executive Summary** — data-driven narrative
- **Key Metrics** — flagged sites, bridge completion, deteriorating/improving/new/resolved
- **7 Charts** — compliance by site, root cause pareto, miss type breakdown, WoW trending, deteriorating sites, country breakdown, flagged sites per week
- **Top 5 Root Causes** — with impact, actions, owner, timeline
- **WBR Talking Points** — ready to use in the meeting
- **Appendix** — sites without bridges (accountability tracker)

## Slack Integration

Post results to Slack automatically:

```bash
# With webhook
SLACK_WEBHOOK_URL=https://hooks.slack.com/services/... ./run.sh file.xlsx Wk-14

# With bot token (supports file uploads)
SLACK_BOT_TOKEN=xoxb-... ./run.sh file.xlsx Wk-14
```

## Architecture

```
run.sh                    → One-command entry point
main.py                   → CLI orchestrator
excel_parser.py           → Parses Excel, auto-detects headers
analyzer.py               → Bedrock Claude analysis + offline fallback
charts.py                 → 4 WBR-style charts (matplotlib)
wow_analysis.py           → Week-over-week trending + 3 additional charts
report_generator.py       → Basic Word + PowerPoint
report_generator_v2.py    → Enhanced reports with embedded charts
generate_final_report.py  → Final comprehensive report with all charts + WoW
slack_notify.py           → Slack webhook/bot integration
```

## Requirements

- Python 3.10+
- Dependencies: `pip3 install -r requirements.txt`
- Optional: AWS credentials for Bedrock Claude (richer analysis)
- Optional: Slack webhook/bot token for auto-posting
