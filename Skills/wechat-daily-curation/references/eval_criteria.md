# WeChat Curation — Output Scoring System (Reward Function)

Purpose: define a single scalar reward `R_push` for the output of each curation
run, used both to **select** the best push at inference (best-of-N) and to
**improve** the cron prompts over time (prompt hill-climbing / RL-style loop).

Scope: this scores **final output quality only**. It does not evaluate the
agent's trajectory, tool usage, or judge reliability. The only judges are an
**LLM reward model** (dense, every run) and **human pairwise labels** (sparse,
ground-truth anchor).

Applies to all four jobs (科技早报 / 深度技术 / 职位 / research papers), with
per-job freshness windows and weights noted where they differ.

---

## 1. Reward structure

```
R_push = Gate(push) × Quality(push)

  Gate(push)    ∈ {0, 1}    hard constraints, programmatic, multiplicative
  Quality(push) ∈ [0, 1]    weighted item quality, dense, LLM-scored
```

Multiplicative by design: a well-written push containing a dead link or stale
content scores **0**, never "high minus a bit." The gate is the kill switch and
the first line of defense against reward hacking.

Reward unit: **per item**, aggregated to a per-push scalar (Section 4). Per-item
scoring gives denser signal and supports selection-quality analysis.

---

## 2. Gate — hard constraints (programmatic, 0/1)

`Gate(push) = 0` if **any** item or the push as a whole fails. No LLM involved.

| Check | Rule | Source |
| --- | --- | --- |
| Link resolves | Every link returns HTTP 200 | — |
| Freshness window | news ≤24h · deep tech ≤7d · jobs = currently open · papers ≤7d | SKILL line 25 |
| No duplicates | No two items share a story/URL cluster | SKILL line 26 |
| Required fields | Each item has source + link + date/recency + "why it matters" | SKILL line 23 |
| Delivered | Push actually reached WeChat (no silent failure) | — |
| Non-empty | At least 1 qualifying item (volume floor, not target — see §4) | — |

Gates are deterministic and free. They guarantee the LLM reward model never
scores structurally broken output.

---

## 3. Quality — per-item scoring (LLM reward model)

Each item is scored on 5 criteria, each `cᵢ ∈ [0, 1]`, by the LLM reward model
using the rubric below. Weights produce a weighted sum.

```
item_score = Σ wᵢ · cᵢ
```

| Criterion | Weight `wᵢ` | `cᵢ = 1.0` (strong) | `cᵢ = 0.0` (weak) |
| --- | --- | --- | --- |
| **Relevance** | 0.30 | Directly on AI Agents / AI-PM work / macro-finance / user's job targets | Off-topic or only tangentially related |
| **Actionability** | 0.25 | User can act: what to learn, apply to, watch, prepare | Pure FYI, no implication drawn |
| **Depth / Insightfulness** | 0.20 | Adds mechanism, data, analysis, practical insight — *and* connects the item to other highly relevant concepts to surface non-obvious, unconventional, or inspiring insight | Headline restatement, no substance, no connections drawn |
| **Authority** | 0.15 | Primary source / reputable media / official job page / credible author | Aggregator rumor, unknown blog, SEO spam |
| **Freshness** | 0.10 | Very recent, original (graded bonus above the gate) | Recycled / borderline within window |

Notes:
- Weights are **starting values**. They should be refit from human pairwise
  labels (Section 5) so the weighted sum reproduces human rankings.
- Freshness is mostly enforced by the gate; this weight is a small graded bonus
  for "exceptionally fresh / original," not a re-test of the window.
- Per-job tweak: for the **jobs** cron, raise Actionability (deadline + how to
  apply) and Authority (official careers page vs aggregator); see SKILL §"Job
  Curation" and the deadline-freshness warning.

### Rubric anti-patterns (force `cᵢ` low)
- Relevance inflated by **sensational framing** → anchor to substance, not hook.
- Depth proxied by **length** → reward mechanism/data, not word count.
- Authority faked by **confident tone** → require a verifiable primary source.

---

## 4. Aggregation — items → push

Do **not** use a plain mean (rewards padding neutrally). Use top-k peaks minus a
weak-item penalty:

```
Quality(push) = mean( top_k(item_scores) ) − λ · Σ max(0, floor − item_score)

  top_k    : k best items   (suggest k=3 for news/tech, k=2 for jobs/papers)
  floor    : weak-item threshold (suggest 0.45)
  λ         : padding penalty strength (suggest 0.5), tune on human labels
```

Rationale:
- **top-k** rewards the strongest items — curation is about peaks, not average.
- **penalty term** makes every weak included item *cost* reward, directly
  encoding SKILL's "quality over volume / send fewer strong items" (lines 22, 26)
  as negative reward rather than neutral.
- The non-empty gate is a **floor, not a target**: hitting a count never raises
  reward; only item quality does.

Clamp `Quality(push)` to `[0, 1]` after the penalty.

---

## 5. Human signal — pairwise preferences (ground truth)

Humans do **not** score every push. They provide sparse, high-reliability labels
that anchor and refit the LLM reward model.

- **Format: pairwise preference**, not absolute scores. "Push A ≻ Push B" (or
  "item A ≻ item B"). Humans are consistent at ranking, noisy at scoring.
- **Sampling**: label a small set per week (e.g. 10–20 pairs), drawn from real
  runs and best-of-N variants.
- **Use of labels (pick one):**
  - **(a) Refit weights (preferred):** fit `wᵢ` via Bradley-Terry / logistic
    regression so the weighted sum agrees with human orderings → data-driven
    weights instead of hand-set ones.
  - **(b) Rubric spot-fix:** use disagreements to patch the LLM rubric prompt.
- **Calibration check (minimal but required):** periodically measure preference
  agreement (e.g. Kendall's τ or % pairs agreed) between the LLM reward model and
  human labels. This is not "judge reliability for its own sake" — it is the only
  guard that stops the optimization loop from exploiting a reward that has
  drifted from the user's actual taste. Keep it light; do not skip it entirely.

---

## 6. Reward-hacking watchlist

The reward will be gamed by whatever the policy can exploit. Designed-in defenses:

| Exploit | Defense |
| --- | --- |
| Padding with filler items | top-k aggregation + weak-item penalty (§4) |
| Sensational / clickbait framing | relevance rubric anchored to substance + human prefs punish it |
| Recency-only chasing | freshness capped at w=0.10 and mostly gated |
| Length = depth | depth rubric rewards mechanism/data, not word count; mobile-readability cap |
| Source monoculture (same easy RSS daily) | optional diversity bonus across source/topic (§7) |

---

## 7. Optional: diversity bonus

To discourage source/topic monoculture, add a small additive term before clamp:

```
Quality(push) += β · diversity(push)      β small (≈0.05)
  diversity = normalized count of distinct sources AND distinct topics
```

Off by default; enable only if monoculture appears in practice.

---

## 8. How the reward is used (the optimization loop)

The policy is the cron prompt + selection instructions (`references/*_prompt.txt`);
there are no trainable weights, so "RL" = hill-climbing the prompt on this reward.

```
1. Run cron → produce push
2. Gate (programmatic) → 0/1
3. LLM reward model scores items → Quality(push) → R_push
4. Weekly: collect human pairwise labels → refit wᵢ / fix rubric → recalibrate
5. Improve policy to raise R:
     - best-of-N at inference: generate several candidate pushes, ship the
       highest-R one, log the rest as preference data
     - prompt optimization: vary selection criteria / source priority / item
       caps; keep variants that raise mean R on a fixed eval set
```

- **Fixed eval set:** snapshot ~15–20 past runs (with candidate pools) and
  re-score after every prompt edit to catch regressions.
- **Track trends, not single runs:** curation reward is noisy day-to-day; watch
  the rolling distribution (mean R, gate pass-rate) per week.

---

## 9. Implementation target

`scripts/score_push.py`:
- input: a push (items with fields) + optional candidate pool
- `Gate()` — programmatic checks (§2), returns 0/1 with reasons
- `score_items()` — LLM call returning `cᵢ` per item against the §3 rubric
- `aggregate()` — top-k minus penalty (§4) → `Quality`
- output: `R_push`, per-item breakdown, gate failures
- `refit_weights(pairwise_labels)` — Bradley-Terry fit of `wᵢ` (§5)

Logging prerequisite for selection-quality analysis: capture the **candidate
pool** (what the agent considered), not only the final push.
