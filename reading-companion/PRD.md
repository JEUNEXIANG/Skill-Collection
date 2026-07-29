# Reading Companion — PRD

*Status: draft · Owner: ZHENXIANG · Last updated: 2026-06-25*

---

## 1. Problem & Vision

Readers in a non-native or technical domain hit words/phrases they don't understand. The current loop — copy → dictionary → guess which sense applies → switch back — interrupts reading.

**Reading Companion:** select any word or phrase → get the meaning as it applies in this article, in your native language, from your chosen LLM. One tap saves it to Obsidian or a local `.md` file.

No dictionary fallback by design: the value is the contextual meaning, not a list of senses to sift through.

### 1.1 Target scenarios

Two dimensions frame where word/phrase help is needed. The two reading objects differ not just in subject but in **what "help" means**:

| Reading object | What the reader needs | Emphasis |
|---|---|---|
| **Professional / academic** (law, engineering, CS, …) | **Explanation** in context | Correctness + understanding the process, structure, or mechanism the term *denotes* — not just a synonym for it |
| **Foreign-language literature** | **Translation** in context | The in-context, literary sense — how it reads *here* |

By **platform**, the need splits into offline PDF readers vs. online website reading.

**v1 focus:**
- **Platform:** online **website reading** (offline PDF readers are out of scope).
- **Reading objects:** foreign-language **literature** *and* **professional content read casually** — a curious reader deepening understanding, not a practitioner acting on it.

### 1.2 Positioning & non-goals
- **Casual / learning context, not high-accountability use.** It is **not** built for life- or rights-critical interpretation — e.g. the dosage term on a **medicine label** where a mistake risks health, or the **legal interpretation of a word in a case**. Those demand authoritative, accountable sources, which this tool deliberately does not claim to be.

---

## 2. Two Product Forms (same reading experience, two save paths)

**The reading/explanation experience is ALWAYS a browser extension.** Only a browser extension can draw a "tap a word" UI onto a live webpage. Code installed into Obsidian runs only inside the Obsidian window and cannot augment pages you read in Chrome. So both forms read the article on the live website via the extension — they differ only in **how saving into Obsidian works**.

| | **Form B — Standalone extension** (v1) | **Form A — Extension + companion Obsidian plugin** (richer save) |
|---|---|---|
| Reads article | Browser extension, on the live web | Browser extension, on the live web (identical) |
| Obsidian-side install | None | Our own small Obsidian plugin |
| How it saves long-term | `obsidian://` URI link, or local `.md` download | Extension → `http://localhost:<port>` → plugin writes vault file directly |
| Save capability | One-way, limited (create/append a note; URL-length limits; can steal focus) | Two-way, rich (structured JSON, choose folder/template, append, dedup, "saved ✓" reply) |
| Requirement | Nothing on Obsidian side | Obsidian running + our plugin installed at save time |

**How Form A's localhost save works:** our Obsidian plugin opens a tiny web server on the user's own machine (e.g. `http://localhost:27123`) that never touches the internet. The extension POSTs the entry to that address; the plugin — which has full vault file access — writes the `.md` file and replies "saved." The user installs **our** plugin rather than a third-party one.

**v1 build target: Form B, Chrome only.** Form A is post-v1 (§10).

---

## 3. v1 Scope

In scope:
1. Per-site consent to read the article.
2. Select word/phrase → inline popup with contextual explanation in the user's native language.
3. Save a lookup → adds it to the article's side-panel list (short-term, manual, 7-day expiry).
4. Side-panel dictionary (native Chrome side panel) — the saved list for the current article.
5. Export the side-panel list → Obsidian via `obsidian://` or local `.md` download.
6. Settings: provider + API key + model, native language, vault name, save target.

Out of scope: Safari, Form A, spaced-repetition, offline/PDF, dictionary fallback.

---

## 4. Extension File Structure

```
reading-companion/
├── manifest.json        permissions, entry points (MV3)
├── background.js        LLM API calls (service worker)
├── content.js           runs in the article page: captures selection, shows popup
├── sidepanel.html/.js   session dictionary, Save buttons
├── settings.html/.js    provider, key, model, language, vault
├── prompt-template.txt  system-prompt template (the "agent.md")
└── icons/
```

---

## 5. Data Flow


```
User selects "X"
        │
content.js: capture X + full article text + URL + title
        │
background.js: fill prompt-template, call LLM → spinner
        ▼
popup: meaning_in_context → explanation
        │
written to short-term memory (browser storage)
        │
[ Save ] ──► obsidian:// URI  /  local .md download
```

- **Lazy:** LLM called only on user action, never pre-translated.
- **Full article as context:** covers the word's meaning and its role in the article's argument.
- **Cache per word/phrase per article:** re-tapping is free.
- **Spinner-then-result:** full response rendered at once.

---

## 6. Explanation Core

### 6.1 Provider adapter
User picks provider; swapping is config, not code. Supports: Claude, OpenAI, Gemini, any OpenAI-compatible endpoint. API key stored locally, never in the repo.

### 6.2 Prompt template
Slots: `{selection}`, `{article_text}`, `{article_title}`, `{native_language}`.

Instruction: explain `{selection}` as used in this specific article — its meaning and how it relates to the surrounding content — in `{native_language}`; treat phrases as a unit, not word-by-word.

**Auto-adapts** to reading type: technical terms get the mechanism/process they denote; literary phrases get their in-context interpretation. Fallback: explicit Explain/Translate toggle in the popup if auto-detection proves unreliable.

### 6.3 Structured output

```json
{
  "selection": "…",
  "pronunciation": "… (optional)",
  "meaning_in_context": "…",
  "general_meaning": "… (optional)",
  "explanation": "1–3 sentences",
  "example": "… (optional)"
}
```

Popup renders: **meaning_in_context → explanation**. Article URL is captured by the extension and appended silently at save time.

---

## 7. Memory

**Short-term (the side-panel list):** populated by an explicit **Save** on a
lookup — not auto-logged. One bucket per article URL in browser storage, shown in
the native Chrome side panel, each item auto-expiring 7 days after it was saved.
This is the user's curated collection for the article. Each item: word/phrase,
explanation fields, surrounding context, article URL, savedAt.

**Long-term:** a separate **Export** action promotes the side-panel list to a
**single per-article note** (keyed by URL/title) in Obsidian (via `obsidian://`)
or a local `.md`. (Export is a later increment.)

> Note: recording is manual (Save), so short-term and long-term are two explicit
> steps — *curate in the panel*, then *export the panel* — rather than an
> auto-log plus a save.

---

## 8. Settings

- Provider + API key + model
- Native language (defaults from browser locale, changeable)
- Default save target (Obsidian vault name / local download)
- Per-site consent (remembered)
- Daily token cap (optional, post-v1)

---

## 9. Privacy & Consent

- Per-site consent required before reading any article.
- Consent copy states: full article text + selected phrase are sent to the user's chosen LLM provider.
- API key and all settings are local to the browser.

---

## 10. Future

- Form A companion Obsidian plugin (localhost, direct vault writes).
- Safari Web Extension.
- Spaced-repetition review over long-term memory.
- **Optional web-search / tool enrichment (opt-in, off by default).** Only for
  terms the article does not itself define (obscure entities, post–training-cutoff
  references). Not in v1: the article-as-context design already covers the core
  "meaning here" need, and live search risks reintroducing the "sift external
  definitions" problem v1 deliberately removes. Also complicates the
  provider-agnostic adapter (tool-calling support varies by provider).
