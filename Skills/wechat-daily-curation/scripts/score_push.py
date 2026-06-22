#!/usr/bin/env python3
"""
score_push.py — WeChat Curation Output Scoring System
======================================================

Implements the reward function from references/eval_criteria.md.

Usage:
    python3 scripts/score_push.py < push.json          # score a single push
    python3 scripts/score_push.py --candidates pool.json  # best-of-N selection

Input JSON format (stdin or file):
    [
      {
        "title": "Article title",
        "source": "Source name",
        "link": "https://...",
        "date": "2026-06-16",
        "why": "why this matters"
      },
      ...
    ]

Output: human-readable report (default) or JSON (--json flag).
"""

import json
import sys
from pathlib import Path
from datetime import datetime, timezone

# ── Config ───────────────────────────────────────────────────────────────────
K = 3                     # top-k for deep tech (use 2 for jobs/papers)
FLOOR = 0.45              # weak-item threshold
LAMBDA = 0.5              # padding penalty strength

WEIGHTS = {
    "relevance":     0.30,
    "actionability": 0.25,
    "depth":         0.20,
    "authority":     0.15,
    "freshness":     0.10,
}

TRUSTED_SOURCES = [
    "搜狐", "36氪", "CB Insights", "Google DeepMind", "Anthropic",
    "小米", "华为", "快手", "腾讯云", "IBM", "新华网", "环球",
    "LangChain", "GitHub", "arXiv",
]


# ── Gate ─────────────────────────────────────────────────────────────────────

def check_gate(items):
    """Gate = 0 if any check fails. Deterministic, programmatic."""
    failures = []

    # Non-empty
    if not items or len(items) < 1:
        failures.append("NON_EMPTY: zero items")

    # Required fields
    required = {"title", "source", "link", "why", "date"}
    for i, item in enumerate(items):
        missing = required - set(item.keys())
        if missing:
            failures.append(f"REQUIRED_FIELDS: item {i} missing {missing}")

    # Freshness window (≤7d for deep tech)
    for i, item in enumerate(items):
        try:
            days_raw = item.get("freshness_days")
            if days_raw is None and item.get("date"):
                d = datetime.strptime(item["date"], "%Y-%m-%d")
                days_raw = (datetime.now(timezone.utc) - d.replace(tzinfo=timezone.utc)).days
            if days_raw is not None and days_raw > 7:
                failures.append(f"FRESHNESS: item {i} '{item.get('title','')[:30]}' is {days_raw}d old (max 7d)")
        except (ValueError, TypeError):
            failures.append(f"FRESHNESS: item {i} unparseable date '{item.get('date')}'")

    # No duplicates (by URL, case-insensitive)
    seen = set()
    for i, item in enumerate(items):
        url = item.get("link", "").lower().strip()
        if url in seen:
            failures.append(f"DUPLICATE_URL: item {i} '{item.get('title','')[:30]}' shares URL with another item")
        seen.add(url)

    return {
        "pass": len(failures) == 0,
        "failures": failures,
    }


# ── Quality Scoring ──────────────────────────────────────────────────────────

def score_item(item):
    """Score a single item against the 5 criteria rubric.

    In production this should call an LLM reward model per eval_criteria.md §3.
    This heuristic reference implementation exists so the aggregation pipeline
    can be tested end-to-end without LLM API costs.
    """
    t = item.get("title", "")
    w = item.get("why", "")
    src = item.get("source", "")
    days = item.get("freshness_days", 7)

    text = (t + " " + w).lower()

    # 1. Relevance (w=0.30)
    relevance = 0.5
    if any(kw in text for kw in ["ai agent", "agent", "智能体", "多智能体"]):
        relevance = 0.85
    if any(kw in w for kw in ["AI PM", "产品经理", "PM", "求职", "面试"]):
        relevance = min(1.0, relevance + 0.10)

    # 2. Actionability (w=0.25)
    actionability = 0.4
    if any(kw in w for kw in ["参考", "值得关注", "能力边界", "案例", "借鉴"]):
        actionability = 0.6
    if any(kw in w for kw in ["面试", "差异化", "PM必须", "能力要求", "求职方向"]):
        actionability = 0.8

    # 3. Depth / Insightfulness (w=0.20)
    depth_kws = sum(1 for kw in ["实测", "架构", "搭载", "基于", "采用",
                                   "技术", "数据", "成本", "token", "参数",
                                   "上下文", "开源", "协议"] if kw in text)
    depth = min(0.90, 0.30 + depth_kws * 0.08)

    # 4. Authority (w=0.15)
    authority = 0.3
    for ts in TRUSTED_SOURCES:
        if ts.lower() in src.lower():
            if ts in ("CB Insights",):
                authority = 0.90
            elif ts in ("Anthropic", "Google DeepMind"):
                authority = 0.85
            elif ts in ("36氪", "新华网", "环球", "IBM"):
                authority = 0.80
            elif ts in ("搜狐",):
                authority = 0.55
            else:
                authority = 0.70
            break

    # 5. Freshness (w=0.10) — graded bonus above the gate
    if days <= 1:
        freshness = 0.95
    elif days <= 3:
        freshness = 0.80
    elif days <= 5:
        freshness = 0.60
    elif days <= 7:
        freshness = 0.40
    else:
        freshness = 0.10

    item_score = (
        WEIGHTS["relevance"] * relevance +
        WEIGHTS["actionability"] * actionability +
        WEIGHTS["depth"] * depth +
        WEIGHTS["authority"] * authority +
        WEIGHTS["freshness"] * freshness
    )

    return item_score, {
        "relevance":     round(relevance, 3),
        "actionability": round(actionability, 3),
        "depth":         round(depth, 3),
        "authority":     round(authority, 3),
        "freshness":     round(freshness, 3),
    }


# ── Aggregation ──────────────────────────────────────────────────────────────

def aggregate(scores, k=K, floor=FLOOR, lam=LAMBDA):
    """Quality(push) = mean(top_k) - lambda * sum(max(0, floor - score))."""
    top_k = sorted(scores, reverse=True)[:k]
    penalty = sum(max(0, floor - s) for s in scores)
    quality = sum(top_k) / k - lam * penalty
    return max(0.0, min(1.0, quality))


# ── Main ─────────────────────────────────────────────────────────────────────

def evaluate(items):
    gate = check_gate(items)
    scored = []
    for item in items:
        score, breakdown = score_item(item)
        scored.append({
            "title": item.get("title", ""),
            "score": round(score, 3),
            "breakdown": breakdown,
        })
    item_scores = [s["score"] for s in scored]
    quality = aggregate(item_scores)
    r_push = quality if gate["pass"] else 0.0

    if r_push > 0.70:
        rating = "★★★★ Excellent"
    elif r_push > 0.55:
        rating = "★★★ Good"
    elif r_push > 0.40:
        rating = "★★ Adequate"
    else:
        rating = "★ Needs rework"

    # Human-readable report
    print("=" * 65)
    print("CURATION EVALUATION REPORT")
    print("=" * 65)

    print(f"\nGATE:")
    print(f"  {'PASS' if gate['pass'] else 'FAIL'} "
          f"({'R_push = quality × {quality:.3f}' if gate['pass'] else 'R_push = 0 (zeroed by gate)'})")
    for f in gate["failures"]:
        print(f"    ✗ {f}")

    print(f"\nITEMS (sorted by score):")
    print(f"{'#':>3} {'Title':<48} {'Score':>6}")
    print("-" * 60)
    for i, s in enumerate(sorted(scored, key=lambda x: x["score"], reverse=True)):
        print(f"{i+1:>3} {s['title'][:48]:<48} {s['score']:>6.3f}")
        b = s["breakdown"]
        print(f"     R={b['relevance']:.2f} A={b['actionability']:.2f} "
              f"D={b['depth']:.2f} Au={b['authority']:.2f} F={b['freshness']:.2f}")

    print(f"\nAGGREGATION (top-{K}, floor={FLOOR}, lambda={LAMBDA}):")
    top_k = sorted(item_scores, reverse=True)[:K]
    print(f"  top-{K}:     {[round(s,3) for s in top_k]}")
    print(f"  mean:       {sum(top_k)/K:.4f}")
    penalty = sum(max(0, FLOOR - s) for s in item_scores)
    print(f"  penalty:    {LAMBDA} x {penalty:.3f} = {LAMBDA*penalty:.3f}")
    print(f"  Quality:    {quality:.4f}")
    print(f"  R_push:     {r_push:.4f}")
    print(f"  Rating:     {rating}")

    return {
        "gate": gate,
        "items": scored,
        "quality": round(quality, 4),
        "r_push": round(r_push, 4),
        "rating": rating,
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Score WeChat curation pushes")
    parser.add_argument("file", nargs="?", help="JSON file with push items (default: stdin)")
    parser.add_argument("--json", action="store_true", help="Output raw JSON only")
    args = parser.parse_args()

    if args.file:
        with open(args.file) as f:
            items = json.load(f)
    else:
        items = json.load(sys.stdin)

    result = evaluate(items)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
