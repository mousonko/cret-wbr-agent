"""Use Amazon Bedrock (Claude) to analyze bridges and extract root causes."""
import json
import boto3
from excel_parser import SiteEntry


def build_prompt(entries: list[SiteEntry]) -> str:
    sites_text = ""
    for e in entries:
        sites_text += f"""
---
Site: {e.ds} (MP: {e.mp})
Week: {e.week}
Trending Scan Compliance (T4W): {e.trending_scan_compliance}
WK-14 Scan Compliance: {e.wk14_scan_compliance}
Deep-dive: {e.deep_dive}
Pickup to Stow (<2 days): {e.pickup_to_stow}
Pickup to Depart (<7 days): {e.pickup_to_depart}
DD on RTS: {e.dd_on_rts}
Bridge: {e.bridge}
Expected Improvement Week: {e.improvement_week}
POC: {e.poc}
"""

    return f"""You are an expert Amazon Last Mile operations analyst preparing a WBR (Weekly Business Review) bridge analysis for EU LM CRET Scan Compliance.

Below are the flagged sites and their bridges for this week:

{sites_text}

Provide the following analysis in JSON format:
{{
  "executive_summary": "2-3 sentence high-level summary of the week's CRET scan compliance performance across EU LM",
  "total_flagged_sites": <number>,
  "top_5_root_causes": [
    {{
      "rank": 1,
      "root_cause": "Clear description of the root cause",
      "affected_sites": ["site1", "site2"],
      "impact": "Quantified impact where possible",
      "recommended_actions": ["action1", "action2"],
      "owner": "Suggested owner or team",
      "timeline": "Expected resolution timeline"
    }}
  ],
  "site_summaries": [
    {{
      "site": "DS code",
      "mp": "MP",
      "compliance": "WK-14 compliance %",
      "bridge_summary": "1-sentence summary of the bridge",
      "severity": "HIGH/MEDIUM/LOW"
    }}
  ],
  "patterns_and_trends": "Notable patterns across sites",
  "recommended_wbr_talking_points": ["point1", "point2", "point3"]
}}

Be specific, data-driven, and actionable. Group similar root causes together. Prioritize by impact (number of sites affected × compliance gap)."""


def analyze_with_bedrock(
    entries: list[SiteEntry],
    region: str = "eu-west-1",
    model_id: str = "anthropic.claude-3-5-sonnet-20241022-v2:0",
) -> dict:
    """Call Bedrock to analyze the bridges and return structured analysis."""
    if not entries:
        return {"error": "No entries to analyze"}

    client = boto3.client("bedrock-runtime", region_name=region)
    prompt = build_prompt(entries)

    response = client.invoke_model(
        modelId=model_id,
        contentType="application/json",
        accept="application/json",
        body=json.dumps({
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": 4096,
            "messages": [{"role": "user", "content": prompt}],
        }),
    )

    result = json.loads(response["body"].read())
    text = result["content"][0]["text"]

    # Extract JSON from response
    start = text.find("{")
    end = text.rfind("}") + 1
    if start >= 0 and end > start:
        return json.loads(text[start:end])
    return {"raw_response": text}


def analyze_offline(entries: list[SiteEntry]) -> dict:
    """Fallback: rule-based analysis without LLM."""
    from collections import Counter

    bridges = [e.bridge for e in entries if e.bridge]
    # Simple keyword-based root cause clustering
    keywords = {
        "staffing": ["staff", "headcount", "hiring", "capacity", "labor", "HC"],
        "process_compliance": ["process", "training", "SOP", "compliance", "scan"],
        "system_issues": ["system", "tool", "app", "device", "scanner", "IT"],
        "volume_spike": ["volume", "peak", "surge", "demand", "spike"],
        "returns_backlog": ["backlog", "RTS", "return", "pending", "aging"],
    }

    cause_counts = Counter()
    cause_sites = {k: [] for k in keywords}

    for e in entries:
        bridge_lower = (e.bridge + " " + e.deep_dive).lower()
        for cause, kws in keywords.items():
            if any(kw.lower() in bridge_lower for kw in kws):
                cause_counts[cause] += 1
                cause_sites[cause].append(e.ds)

    top_causes = []
    for rank, (cause, count) in enumerate(cause_counts.most_common(5), 1):
        top_causes.append({
            "rank": rank,
            "root_cause": cause.replace("_", " ").title(),
            "affected_sites": cause_sites[cause],
            "impact": f"{count} sites affected",
            "recommended_actions": ["Review site-level bridges for specific actions"],
            "owner": "TBD",
            "timeline": "TBD",
        })

    site_summaries = []
    for e in entries:
        site_summaries.append({
            "site": e.ds,
            "mp": e.mp,
            "compliance": e.wk14_scan_compliance,
            "bridge_summary": e.bridge[:100] if e.bridge else "No bridge provided",
            "severity": "HIGH",
        })

    return {
        "executive_summary": f"{len(entries)} sites flagged for CRET scan compliance this week.",
        "total_flagged_sites": len(entries),
        "top_5_root_causes": top_causes,
        "site_summaries": site_summaries,
        "patterns_and_trends": "See individual bridges for details.",
        "recommended_wbr_talking_points": [
            f"{len(entries)} sites flagged",
            f"Top root cause: {top_causes[0]['root_cause'] if top_causes else 'N/A'}",
        ],
    }
