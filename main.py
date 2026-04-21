#!/usr/bin/env python3
"""EU LM WBR Bridge Analysis Agent — CLI entry point.

Usage:
    python main.py <excel_file> [--sheet SHEET] [--offline] [--region REGION] [--output-dir DIR]

Examples:
    python main.py "DS Bridging Scan Compliance.xlsx" --sheet WK14
    python main.py compliance.xlsx --offline
    python main.py compliance.xlsx --region eu-west-1
"""
import argparse
import json
import os
import sys

from excel_parser import parse_excel
from analyzer import analyze_with_bedrock, analyze_offline
from report_generator_v3 import generate_narrative_word, generate_narrative_pptx


def main():
    parser = argparse.ArgumentParser(description="EU LM WBR CRET Scan Compliance Bridge Analyzer")
    parser.add_argument("excel_file", help="Path to the DS Bridging Scan Compliance Excel file")
    parser.add_argument("--sheet", help="Sheet name (e.g., WK14). Defaults to last sheet.", default=None)
    parser.add_argument("--offline", action="store_true", help="Use rule-based analysis (no Bedrock)")
    parser.add_argument("--region", default="eu-west-1", help="AWS region for Bedrock (default: eu-west-1)")
    parser.add_argument("--output-dir", default=".", help="Output directory for reports")
    args = parser.parse_args()

    if not os.path.exists(args.excel_file):
        print(f"Error: File not found: {args.excel_file}")
        sys.exit(1)

    # 1. Parse Excel
    print("=" * 60)
    print("Step 1: Parsing Excel file...")
    entries = parse_excel(args.excel_file, args.sheet)
    if not entries:
        print("No flagged sites found. Exiting.")
        sys.exit(0)

    # Filter entries that have a bridge
    bridged = [e for e in entries if e.bridge and e.bridge.strip()]
    print(f"  {len(entries)} total entries, {len(bridged)} with bridges")

    # 2. Analyze
    print("=" * 60)
    if args.offline:
        print("Step 2: Running offline (rule-based) analysis...")
        analysis = analyze_offline(entries)
    else:
        print("Step 2: Analyzing with Bedrock Claude...")
        try:
            analysis = analyze_with_bedrock(entries, region=args.region)
        except Exception as e:
            print(f"  Bedrock failed: {e}")
            print("  Falling back to offline analysis...")
            analysis = analyze_offline(entries)

    # Save raw JSON
    os.makedirs(args.output_dir, exist_ok=True)
    json_path = os.path.join(args.output_dir, "wbr_analysis.json")
    with open(json_path, "w") as f:
        json.dump(analysis, f, indent=2)
    print(f"  JSON saved: {json_path}")

    # 3. Generate reports
    print("=" * 60)
    print("Step 3: Generating reports...")

    # Use empty wow dict for web app (no multi-sheet analysis)
    wow = {}

    word_path = os.path.join(args.output_dir, "WBR_Bridge_Analysis.docx")
    generate_narrative_word(analysis, wow, args.output_dir)
    print(f"  Word report: {os.path.join(args.output_dir, 'WBR_Bridge_Analysis_FINAL.docx')}")

    generate_narrative_pptx(analysis, wow, args.output_dir)
    print(f"  PowerPoint: {os.path.join(args.output_dir, 'WBR_Bridge_Analysis_FINAL.pptx')}")

    # 4. Print summary to console
    print("=" * 60)
    print("ANALYSIS SUMMARY")
    print("=" * 60)
    print(f"\n{analysis.get('executive_summary', '')}\n")
    print("TOP ROOT CAUSES:")
    for rc in analysis.get("top_root_causes_narrative", analysis.get("top_5_root_causes", [])):
        if "title" in rc:
            print(f"  RC{rc['rank']} {rc['title']} {rc['narrative']}")
        else:
            print(f"  #{rc['rank']} {rc['root_cause']}")
            print(f"     Sites: {', '.join(rc.get('affected_sites', []))}")
            print(f"     Actions: {'; '.join(rc.get('recommended_actions', []))}")
    print()
    print("WBR TALKING POINTS:")
    for pt in analysis.get("recommended_wbr_talking_points", []):
        print(f"  • {pt}")
    print("=" * 60)


if __name__ == "__main__":
    main()
