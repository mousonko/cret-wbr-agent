"""Week-over-week trending analysis across WK-12, WK-13, WK-14."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import os
import json
from collections import defaultdict
from excel_parser import parse_excel

FILEPATH = "/home/mousonko/DS Bridging Scan Compliance-WK14.xlsx"
OUTPUT = "/home/mousonko/.workspace/wbr-bridge-agent/output"
WEEKS = ["Wk-12", "Wk-13", "Wk-14"]


def load_all_weeks():
    """Load data from all 3 weekly sheets, return {site: {week: compliance}}."""
    site_data = defaultdict(lambda: {"mp": "", "weeks": {}, "bridges": {}})

    for sheet in WEEKS:
        entries = parse_excel(FILEPATH, sheet)
        for e in entries:
            try:
                val = float(e.wk14_scan_compliance)
            except (ValueError, TypeError):
                continue
            site_data[e.ds]["mp"] = e.mp
            site_data[e.ds]["weeks"][sheet] = val
            if e.bridge:
                site_data[e.ds]["bridges"][sheet] = e.bridge

    return dict(site_data)


def classify_trend(data: dict) -> str:
    """Classify site as improving, deteriorating, new, or stable."""
    weeks_present = [w for w in WEEKS if w in data["weeks"]]
    if len(weeks_present) < 2:
        return "new"
    vals = [data["weeks"][w] for w in weeks_present]
    if vals[-1] > vals[-2] + 0.05:
        return "improving"
    elif vals[-1] < vals[-2] - 0.05:
        return "deteriorating"
    return "stable"


def chart_wow_trending(site_data: dict, output_dir: str) -> str:
    """Line chart showing compliance trend for repeat offenders."""
    # Sites present in 2+ weeks
    repeat_sites = {s: d for s, d in site_data.items()
                    if sum(1 for w in WEEKS if w in d["weeks"]) >= 2}

    fig, ax = plt.subplots(figsize=(12, 7))
    week_labels = ["WK-12", "WK-13", "WK-14"]

    for site, data in sorted(repeat_sites.items()):
        vals = [data["weeks"].get(w) for w in WEEKS]
        xs = [i for i, v in enumerate(vals) if v is not None]
        ys = [v for v in vals if v is not None]

        trend = classify_trend(data)
        style = {"deteriorating": ("--", "o", 2.5),
                 "improving": ("-", "^", 2),
                 "stable": ("-", "s", 1.5),
                 "new": (":", "d", 1.5)}
        ls, marker, lw = style.get(trend, ("-", "o", 1.5))

        color = "#d32f2f" if ys[-1] < 0.5 else "#f57c00" if ys[-1] < 0.8 else "#388e3c"
        ax.plot(xs, ys, marker=marker, linestyle=ls, linewidth=lw,
                label=f"{site} ({data['mp']})", color=color, markersize=6)

    ax.axhline(y=0.90, color="#388e3c", linestyle="--", alpha=0.5, label="Target 90%")
    ax.set_xticks(range(len(week_labels)))
    ax.set_xticklabels(week_labels, fontsize=12)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(1.0))
    ax.set_ylabel("Scan Compliance %")
    ax.set_title("Week-over-Week Compliance Trend — Repeat Flagged Sites", fontweight="bold", fontsize=13)
    ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=8, ncol=1)
    ax.set_ylim(-0.05, 1.1)
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    path = os.path.join(output_dir, "chart_wow_trending.png")
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def chart_wow_site_count(site_data: dict, output_dir: str) -> str:
    """Bar chart: number of flagged sites per week + by country."""
    from collections import Counter
    week_counts = {}
    week_country = {}
    for sheet in WEEKS:
        entries = parse_excel(FILEPATH, sheet)
        week_counts[sheet] = len(entries)
        cc = Counter(e.mp for e in entries)
        week_country[sheet] = cc

    countries = sorted(set(c for cc in week_country.values() for c in cc))
    week_labels = ["WK-12", "WK-13", "WK-14"]

    fig, ax = plt.subplots(figsize=(8, 5))
    bottom = [0] * 3
    colors = {"DE": "#1565c0", "UK": "#d32f2f", "ES": "#f57c00", "FR": "#388e3c", "IT": "#7b1fa2"}

    for country in countries:
        vals = [week_country[w].get(country, 0) for w in WEEKS]
        ax.bar(week_labels, vals, bottom=bottom, label=country,
               color=colors.get(country, "#999"), edgecolor="white")
        bottom = [b + v for b, v in zip(bottom, vals)]

    for i, total in enumerate([week_counts[w] for w in WEEKS]):
        ax.text(i, total + 0.3, str(total), ha="center", fontweight="bold", fontsize=12)

    ax.set_ylabel("# Flagged Sites")
    ax.set_title("Flagged Sites per Week by Country", fontweight="bold", fontsize=13)
    ax.legend()
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout()
    path = os.path.join(output_dir, "chart_wow_site_count.png")
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def chart_deteriorating_sites(site_data: dict, output_dir: str) -> str:
    """Grouped bar chart for sites that got worse WK-13 → WK-14."""
    deteriorating = []
    for site, data in site_data.items():
        wk13 = data["weeks"].get("Wk-13")
        wk14 = data["weeks"].get("Wk-14")
        if wk13 is not None and wk14 is not None and wk14 < wk13 - 0.02:
            deteriorating.append((site, data["mp"], wk13, wk14))

    deteriorating.sort(key=lambda x: x[3] - x[2])  # biggest drop first

    if not deteriorating:
        return ""

    fig, ax = plt.subplots(figsize=(10, max(4, len(deteriorating) * 0.5)))
    names = [f"{s} ({mp})" for s, mp, _, _ in deteriorating]
    wk13_vals = [w13 for _, _, w13, _ in deteriorating]
    wk14_vals = [w14 for _, _, _, w14 in deteriorating]

    y = range(len(names))
    ax.barh([i + 0.15 for i in y], wk13_vals, height=0.3, label="WK-13", color="#1565c0", alpha=0.7)
    ax.barh([i - 0.15 for i in y], wk14_vals, height=0.3, label="WK-14", color="#d32f2f")

    # Add delta labels
    for i, (_, _, w13, w14) in enumerate(deteriorating):
        delta = w14 - w13
        ax.text(max(w13, w14) + 0.02, i, f"{delta:+.0%}", va="center", fontsize=9,
                color="#d32f2f", fontweight="bold")

    ax.set_yticks(list(y))
    ax.set_yticklabels(names)
    ax.xaxis.set_major_formatter(mticker.PercentFormatter(1.0))
    ax.axvline(x=0.90, color="#388e3c", linestyle="--", alpha=0.5)
    ax.set_title("Deteriorating Sites: WK-13 → WK-14", fontweight="bold", fontsize=13)
    ax.legend()
    ax.set_xlabel("Scan Compliance")
    plt.tight_layout()
    path = os.path.join(output_dir, "chart_deteriorating.png")
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def build_wow_analysis(site_data: dict) -> dict:
    """Build structured WoW analysis."""
    deteriorating = []
    improving = []
    new_flags = []
    resolved = []

    # Sites in WK-13 but not WK-14 = resolved
    wk13_sites = {s for s, d in site_data.items() if "Wk-13" in d["weeks"]}
    wk14_sites = {s for s, d in site_data.items() if "Wk-14" in d["weeks"]}
    resolved_sites = wk13_sites - wk14_sites
    new_sites = wk14_sites - wk13_sites

    for site, data in site_data.items():
        trend = classify_trend(data)
        wk14 = data["weeks"].get("Wk-14")
        wk13 = data["weeks"].get("Wk-13")
        entry = {
            "site": site, "mp": data["mp"],
            "wk14": f"{wk14:.0%}" if wk14 is not None else "N/A",
            "wk13": f"{wk13:.0%}" if wk13 is not None else "N/A",
            "delta": f"{wk14 - wk13:+.0%}" if wk14 is not None and wk13 is not None else "N/A",
            "bridge_wk14": data["bridges"].get("Wk-14", ""),
            "bridge_wk13": data["bridges"].get("Wk-13", ""),
        }
        if site in new_sites:
            new_flags.append(entry)
        elif trend == "deteriorating":
            deteriorating.append(entry)
        elif trend == "improving":
            improving.append(entry)

    for site in resolved_sites:
        data = site_data[site]
        resolved.append({
            "site": site, "mp": data["mp"],
            "wk13": f"{data['weeks'].get('Wk-13', 0):.0%}",
        })

    return {
        "summary": (
            f"WK-14 has {len(wk14_sites)} flagged sites vs {len(wk13_sites)} in WK-13. "
            f"{len(new_sites)} new flags, {len(resolved_sites)} resolved, "
            f"{len(deteriorating)} deteriorating, {len(improving)} improving."
        ),
        "deteriorating": sorted(deteriorating, key=lambda x: x["delta"]),
        "improving": sorted(improving, key=lambda x: x["delta"], reverse=True),
        "new_flags": new_flags,
        "resolved": resolved,
    }


if __name__ == "__main__":
    print("Loading all weeks...")
    site_data = load_all_weeks()
    print(f"Total unique sites across all weeks: {len(site_data)}")

    print("\nGenerating WoW charts...")
    chart_wow_trending(site_data, OUTPUT)
    chart_wow_site_count(site_data, OUTPUT)
    chart_deteriorating_sites(site_data, OUTPUT)

    print("\nBuilding WoW analysis...")
    wow = build_wow_analysis(site_data)
    print(f"\n{wow['summary']}")

    print(f"\n📉 DETERIORATING ({len(wow['deteriorating'])}):")
    for s in wow["deteriorating"]:
        print(f"  {s['site']} ({s['mp']}): {s['wk13']} → {s['wk14']} ({s['delta']})")
        if s["bridge_wk14"]:
            print(f"    Bridge: {s['bridge_wk14'][:100]}")

    print(f"\n📈 IMPROVING ({len(wow['improving'])}):")
    for s in wow["improving"]:
        print(f"  {s['site']} ({s['mp']}): {s['wk13']} → {s['wk14']} ({s['delta']})")

    print(f"\n🆕 NEW FLAGS ({len(wow['new_flags'])}):")
    for s in wow["new_flags"]:
        print(f"  {s['site']} ({s['mp']}): {s['wk14']}")

    print(f"\n✅ RESOLVED ({len(wow['resolved'])}):")
    for s in wow["resolved"]:
        print(f"  {s['site']} ({s['mp']}): was {s['wk13']} in WK-13")

    # Save
    with open(os.path.join(OUTPUT, "wow_analysis.json"), "w") as f:
        json.dump(wow, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUTPUT}/wow_analysis.json")
