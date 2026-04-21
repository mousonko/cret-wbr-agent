"""Use Amazon Bedrock (Claude) to analyze bridges and extract root causes."""
import json
import boto3
from excel_parser import SiteEntry


def build_prompt(entries: list) -> str:
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
    entries: list,
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


def analyze_offline(entries: list) -> dict:
    """Smart offline analysis — clusters bridges by theme and generates narratives."""
    from collections import Counter, defaultdict

    bridged = [e for e in entries if e.bridge and e.bridge.strip()]
    no_bridge = [e for e in entries if not e.bridge or not e.bridge.strip()]
    no_bridge_mps = Counter(e.mp for e in no_bridge)
    top_no_bridge_mp = no_bridge_mps.most_common(1)[0][0] if no_bridge_mps else "N/A"

    # Smarter keyword patterns with priorities
    patterns = {
        "bank_holiday_truck": {
            "keywords": ["bank holiday", "BH", "holiday", "truck", "B2FC", "no truck", "adhoc", "ad-hoc", "pallet"],
            "title": "Bank Holiday / No B2FC Truck — Items dwelling >7 days",
            "action_template": "sites to raise ad-hoc truck request prior BH to avoid delay. Pre-BH truck scheduling SOP needed. Owner: {owner}. ETA: TBD.",
        },
        "backlog": {
            "keywords": ["backlog", "BL", "collection", "CAP", "catch up", "clear floor", "insufficient"],
            "title": "Backlog — Insufficient collection capacity",
            "action_template": "Escalate to transport and align with Relo sites if adhoc sourcing fails. Owner: {owner}. ETA: TBD.",
        },
        "sop_3pl": {
            "keywords": ["3PL", "3pl", "SOP", "depart scan missing", "missing depart", "wrong scan", "SA ", "training"],
            "title": "SOP adherence / 3PL depart scan compliance gap",
            "action_template": "Recommend weekly 3PL audit cadence. SAM to ensure SOP compliance by providing refresher training. Owner: SAM. ETA: TBD.",
        },
        "tech_system": {
            "keywords": ["tech", "system", "app", "label", "P180", "infinity", "device", "crash", "bug", "IT"],
            "title": "Tech / System issues",
            "action_template": "Escalated and being monitored. SAM to confirm resolution. Owner: SAM/Tech. ETA: TBD.",
        },
        "volume": {
            "keywords": ["volume", "spike", "surge", "peak", "high volume", "PPH"],
            "title": "Volume spike — Capacity not adjusted",
            "action_template": "Review staffing model vs volume forecast. Owner: Ops. ETA: TBD.",
        },
        "process_jedi": {
            "keywords": ["JEDI", "virtual depart", "wrong", "stow", "FIFO", "misunderstanding"],
            "title": "Process errors — JEDI / Virtual depart / FIFO",
            "action_template": "Reinforce process training at affected stations. Owner: Ops. ETA: TBD.",
        },
    }

    cause_sites = defaultdict(list)
    cause_entries = defaultdict(list)
    matched_sites = set()

    for e in entries:
        text = e.bridge.lower() if e.bridge else ""
        if not text:
            continue
        for cause_key, pattern in patterns.items():
            if any(kw.lower() in text for kw in pattern["keywords"]):
                if e.ds not in [s.ds for s in cause_entries[cause_key]]:
                    cause_sites[cause_key].append(e.ds)
                    cause_entries[cause_key].append(e)
                    matched_sites.add(e.ds)

    # Build narrative root causes sorted by site count
    sorted_causes = sorted(cause_sites.items(), key=lambda x: -len(x[1]))
    narrative_causes = []
    structured_causes = []

    for rank, (cause_key, sites) in enumerate(sorted_causes[:5], 1):
        pattern = patterns[cause_key]
        entries_for_cause = cause_entries[cause_key]
        mps = set(e.mp for e in entries_for_cause)
        owner = "/".join(sorted(mps)) + " MSL"

        # Build site detail string with compliance
        site_details = []
        for e in entries_for_cause:
            try:
                pct = f"{float(e.wk14_scan_compliance)*100:.0f}%"
            except (ValueError, TypeError):
                pct = e.wk14_scan_compliance
            site_details.append(f"{e.ds} ({pct})")

        sites_str = ", ".join(site_details)
        action = pattern["action_template"].format(owner=owner)

        narrative_causes.append({
            "rank": rank,
            "title": f"{pattern['title']}:",
            "narrative": f"{sites_str} — {len(sites)} sites affected. {action}",
        })

        structured_causes.append({
            "rank": rank,
            "root_cause": pattern["title"],
            "affected_sites": sites,
            "impact": f"{len(sites)} sites affected",
            "recommended_actions": [action],
            "owner": owner,
            "timeline": "TBD",
        })

    # Site summaries
    site_summaries = []
    for e in entries:
        try:
            pct = float(e.wk14_scan_compliance)
            pct_str = f"{pct*100:.0f}%"
            severity = "HIGH" if pct < 0.5 else "MEDIUM" if pct < 0.8 else "LOW"
        except (ValueError, TypeError):
            pct_str = e.wk14_scan_compliance
            severity = "HIGH"
        site_summaries.append({
            "site": e.ds,
            "mp": e.mp,
            "compliance": pct_str,
            "bridge_summary": e.bridge[:120] if e.bridge else "No bridge provided",
            "severity": severity,
        })

    site_summaries.sort(key=lambda s: s["compliance"])

    return {
        "executive_summary": (
            f"{len(entries)} sites flagged for CRET scan compliance this week across "
            f"{', '.join(f'{mp} ({c})' for mp, c in Counter(e.mp for e in entries).most_common())}. "
            f"Only {len(bridged)} of {len(entries)} sites provided bridges — "
            f"{len(no_bridge)} sites (primarily {top_no_bridge_mp}) have no bridge or POC assigned."
        ),
        "total_flagged_sites": len(entries),
        "sites_with_bridges": len(bridged),
        "sites_without_bridges": len(no_bridge),
        "bridge_summary": (
            f"{len(entries)} sites flagged this week — {len(no_bridge)} without bridges, "
            f"primarily {top_no_bridge_mp}. Action: enforce bridge completion by next WBR. "
            f"Owner: {top_no_bridge_mp} MSL."
        ),
        "top_root_causes_narrative": narrative_causes,
        "top_5_root_causes": structured_causes,
        "site_summaries": site_summaries,
        "recommended_wbr_talking_points": [
            f"{len(entries)} sites flagged — {len(no_bridge)} without bridges, primarily {top_no_bridge_mp}. Enforce bridge completion.",
        ] + [f"RC{rc['rank']} {rc['title']} {rc['narrative'][:100]}..." for rc in narrative_causes[:3]],
    }
