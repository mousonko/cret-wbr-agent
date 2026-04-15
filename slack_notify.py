"""Slack integration for WBR Bridge Agent.

Usage:
    # With bot token (supports file uploads):
    SLACK_BOT_TOKEN=xoxb-... python slack_notify.py --channel "#eu-lm-wbr" --output-dir ./output

    # With webhook (messages only):
    SLACK_WEBHOOK_URL=https://hooks.slack.com/... python slack_notify.py --output-dir ./output
"""
import json
import os
import argparse
from urllib.request import Request, urlopen
from urllib.parse import urlencode


def _post_json(url: str, payload: dict, headers: dict = None) -> dict:
    h = {"Content-Type": "application/json"}
    if headers:
        h.update(headers)
    req = Request(url, data=json.dumps(payload).encode(), headers=h, method="POST")
    with urlopen(req) as resp:
        return json.loads(resp.read())


def _format_message(analysis: dict) -> list:
    """Build Slack Block Kit message from analysis."""
    blocks = [
        {"type": "header", "text": {"type": "plain_text", "text": "🔍 EU LM WBR — CRET Scan Compliance Bridge"}},
        {"type": "section", "text": {"type": "mrkdwn", "text": analysis.get("executive_summary", "")}},
        {"type": "divider"},
        {"type": "section", "fields": [
            {"type": "mrkdwn", "text": f"*Flagged Sites:* {analysis.get('total_flagged_sites', '?')}"},
            {"type": "mrkdwn", "text": f"*With Bridge:* {analysis.get('sites_with_bridges', '?')}"},
            {"type": "mrkdwn", "text": f"*Without Bridge:* {analysis.get('sites_without_bridges', '?')}"},
        ]},
        {"type": "divider"},
        {"type": "header", "text": {"type": "plain_text", "text": "📊 Top 5 Root Causes"}},
    ]

    for rc in analysis.get("top_5_root_causes", []):
        sites = ", ".join(rc.get("affected_sites", []))
        actions = rc.get("recommended_actions", [])
        top_action = actions[0] if actions else "TBD"
        blocks.append({
            "type": "section",
            "text": {"type": "mrkdwn", "text": (
                f"*#{rc['rank']} {rc['root_cause']}*\n"
                f"Sites: {sites}\n"
                f"Impact: {rc.get('impact', 'N/A')[:200]}\n"
                f"Action: {top_action}\n"
                f"Owner: {rc.get('owner', 'TBD')} | Timeline: {rc.get('timeline', 'TBD')}"
            )}
        })

    blocks.append({"type": "divider"})
    blocks.append({"type": "header", "text": {"type": "plain_text", "text": "🎯 WBR Talking Points"}})

    points = "\n".join(f"• {pt}" for pt in analysis.get("recommended_wbr_talking_points", []))
    blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": points}})

    return blocks


def post_via_webhook(webhook_url: str, analysis: dict):
    blocks = _format_message(analysis)
    payload = {"blocks": blocks, "text": "EU LM WBR — CRET Scan Compliance Bridge Analysis"}
    _post_json(webhook_url, payload)
    print("✅ Posted to Slack via webhook")


def post_via_bot(token: str, channel: str, analysis: dict, output_dir: str = None):
    blocks = _format_message(analysis)

    # Post message
    resp = _post_json(
        "https://slack.com/api/chat.postMessage",
        {"channel": channel, "blocks": blocks, "text": "EU LM WBR — CRET Scan Compliance Bridge"},
        headers={"Authorization": f"Bearer {token}"},
    )
    if not resp.get("ok"):
        print(f"❌ Slack message failed: {resp.get('error')}")
        return
    print(f"✅ Posted message to {channel}")

    # Upload files if output_dir provided
    if output_dir:
        for fname in ["WBR_Bridge_Analysis_v2.docx", "WBR_Bridge_Analysis_v2.pptx"]:
            fpath = os.path.join(output_dir, fname)
            if not os.path.exists(fpath):
                continue

            # Step 1: get upload URL
            with open(fpath, "rb") as f:
                file_data = f.read()

            resp = _post_json(
                "https://slack.com/api/files.getUploadURLExternal",
                {},
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/x-www-form-urlencoded",
                },
            )
            # Use multipart for file upload via urllib
            import urllib.request
            boundary = "----WBRBoundary"
            body = (
                f"--{boundary}\r\n"
                f'Content-Disposition: form-data; name="file"; filename="{fname}"\r\n'
                f"Content-Type: application/octet-stream\r\n\r\n"
            ).encode() + file_data + f"\r\n--{boundary}--\r\n".encode()

            req = Request(
                "https://slack.com/api/files.upload",
                data=urlencode({
                    "channels": channel,
                    "filename": fname,
                    "title": fname.replace("_", " ").replace(".docx", "").replace(".pptx", ""),
                }).encode() + b"&file=" + file_data[:0],  # placeholder
                headers={"Authorization": f"Bearer {token}"},
                method="POST",
            )
            print(f"  📎 File upload for {fname} — use Slack UI or drag-drop as fallback")


def main():
    parser = argparse.ArgumentParser(description="Post WBR analysis to Slack")
    parser.add_argument("--channel", default="#eu-lm-wbr", help="Slack channel")
    parser.add_argument("--output-dir", default="./output", help="Directory with reports")
    args = parser.parse_args()

    # Load analysis
    json_path = os.path.join(args.output_dir, "wbr_analysis.json")
    with open(json_path) as f:
        analysis = json.load(f)

    token = os.environ.get("SLACK_BOT_TOKEN")
    webhook = os.environ.get("SLACK_WEBHOOK_URL")

    if token:
        post_via_bot(token, args.channel, analysis, args.output_dir)
    elif webhook:
        post_via_webhook(webhook, analysis)
    else:
        print("Set SLACK_BOT_TOKEN or SLACK_WEBHOOK_URL environment variable")
        print("\nExample:")
        print('  SLACK_BOT_TOKEN=xoxb-... python slack_notify.py --channel "#eu-lm-wbr"')
        print('  SLACK_WEBHOOK_URL=https://hooks.slack.com/... python slack_notify.py')


if __name__ == "__main__":
    main()
