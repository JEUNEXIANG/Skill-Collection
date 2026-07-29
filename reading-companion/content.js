// Runs inside every article page. On a text selection it shows a small "Explain"
// chip; clicking it captures the selection + full article text, asks the
// background worker to explain it, and renders the result in a popup.

let chip = null;
let popup = null;

function removeChip() {
  if (chip) {
    chip.remove();
    chip = null;
  }
}

function removePopup() {
  if (popup) {
    popup.remove();
    popup = null;
  }
}

// Best-effort article extraction for the skeleton. Readability comes in a later
// increment; for now prefer the main article container, fall back to body.
function extractArticleText() {
  const root =
    document.querySelector("article") ||
    document.querySelector("main") ||
    document.body;
  return (root.innerText || "").trim();
}

function pagePos(rect) {
  return {
    top: window.scrollY + rect.bottom + 6,
    left: window.scrollX + rect.left,
  };
}

document.addEventListener("mouseup", (e) => {
  if (e.target.closest && e.target.closest(".rc-ui")) return; // ignore our own UI
  const sel = window.getSelection();
  const text = sel ? sel.toString().trim() : "";
  removeChip();
  if (!text || text.length > 200) return; // skip empty / very long selections
  const range = sel.getRangeAt(0);
  const rect = range.getBoundingClientRect();
  const context = getContextSnippet(range);
  const sentence = getSentence(range, text);
  showChip(rect, text, context, sentence);
});

function blockOf(range) {
  const node = range.startContainer;
  const el = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
  return (
    (el &&
      el.closest(
        "p, li, blockquote, td, h1, h2, h3, h4, h5, h6, section, article, div"
      )) ||
    el
  );
}

// The block the selection sits in — stored with a saved item so the long-term
// note can show where the word appeared.
function getContextSnippet(range) {
  const block = blockOf(range);
  return (block?.innerText || "").replace(/\s+/g, " ").trim().slice(0, 240);
}

// The single sentence containing the selection — sent to the model as the anchor.
function getSentence(range, selectionText) {
  const block = blockOf(range);
  const text = (block?.innerText || "").replace(/\s+/g, " ").trim();
  if (!text || !selectionText) return text.slice(0, 300);
  const idx = text.indexOf(selectionText);
  if (idx === -1) return text.slice(0, 300);
  const ender = /[.!?。！？]/;
  let start = idx;
  while (start > 0 && !ender.test(text[start - 1])) start--;
  let end = idx + selectionText.length;
  while (end < text.length && !ender.test(text[end])) end++;
  if (end < text.length) end++; // include the closing punctuation
  return text.slice(start, end).trim();
}

document.addEventListener("mousedown", (e) => {
  if (e.target.closest && e.target.closest(".rc-ui")) return;
  removePopup();
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    removeChip();
    removePopup();
  }
});

function showChip(rect, text, context, sentence) {
  const { top, left } = pagePos(rect);
  chip = document.createElement("div");
  chip.className = "rc-ui rc-chip";
  chip.textContent = "📖 Explain";
  chip.style.top = `${top}px`;
  chip.style.left = `${left}px`;
  // mousedown (not click) so we act before the browser clears the selection
  chip.addEventListener("mousedown", (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    lookup(text, rect, context, sentence);
  });
  document.body.appendChild(chip);
}

function showPopup(rect, contentNode) {
  removePopup();
  const { top, left } = pagePos(rect);
  popup = document.createElement("div");
  popup.className = "rc-ui rc-popup";
  popup.style.top = `${top}px`;
  popup.style.left = `${left}px`;
  popup.appendChild(contentNode);
  document.body.appendChild(popup);
}

async function lookup(text, rect, context, sentence) {
  removeChip();
  showPopup(rect, spinnerNode());
  try {
    const resp = await chrome.runtime.sendMessage({
      type: "explain",
      payload: {
        selection: text,
        contextSentence: sentence,
        articleText: extractArticleText(),
        title: document.title,
        url: location.href,
      },
    });
    if (!resp || !resp.ok) throw new Error(resp?.error || "Unknown error");
    const entry = {
      ...resp.result,
      context,
      url: location.href,
      title: document.title,
    };
    showPopup(rect, resultNode(entry));
  } catch (err) {
    showPopup(rect, errorNode(String(err.message || err)));
  }
}

// Save the looked-up entry into this article's side-panel bucket.
async function saveEntry(entry, buttonEl) {
  const key = "rc:lookups:" + entry.url;
  const store = await chrome.storage.local.get(key);
  const bucket = store[key] || { url: entry.url, title: entry.title, items: [] };
  bucket.title = entry.title;
  bucket.items.push({
    selection: entry.selection || "",
    meaning_in_context: entry.meaning_in_context || "",
    substance: entry.substance || "",
    distinction:
      entry.distinction && entry.distinction.vs ? entry.distinction : null,
    diagram: entry.diagram || "",
    related_terms: Array.isArray(entry.related_terms) ? entry.related_terms : [],
    context: entry.context || "",
    savedAt: Date.now(),
  });
  await chrome.storage.local.set({ [key]: bucket });
  if (buttonEl) {
    buttonEl.textContent = "Saved ✓";
    buttonEl.disabled = true;
  }
}

// ---- render helpers (use textContent to keep model output out of the DOM as HTML) ----

function spinnerNode() {
  const wrap = document.createElement("div");
  wrap.className = "rc-row";
  const spin = document.createElement("div");
  spin.className = "rc-spinner";
  const label = document.createElement("span");
  label.textContent = "Explaining…";
  wrap.append(spin, label);
  return wrap;
}

function field(label, value, cls) {
  const el = document.createElement("div");
  el.className = `rc-field ${cls || ""}`.trim();
  if (label) {
    const k = document.createElement("span");
    k.className = "rc-label";
    k.textContent = label;
    el.appendChild(k);
  }
  const v = document.createElement("span");
  v.className = "rc-value";
  v.textContent = value;
  el.appendChild(v);
  return el;
}

function resultNode(r) {
  const wrap = document.createElement("div");

  const head = document.createElement("div");
  head.className = "rc-head";
  const sel = document.createElement("span");
  sel.className = "rc-selection";
  sel.textContent = r.selection || "";
  head.appendChild(sel);
  wrap.appendChild(head);

  // Concise by default: the intuitive meaning is always shown.
  if (r.meaning_in_context)
    wrap.appendChild(field("", r.meaning_in_context, "rc-meaning"));

  // Richer detail, hidden until "More" is clicked.
  const details = document.createElement("div");
  details.className = "rc-details";
  details.style.display = "none";
  if (r.substance)
    details.appendChild(field("Substance", r.substance, "rc-secondary"));
  if (r.distinction && r.distinction.vs) {
    const d = document.createElement("div");
    d.className = "rc-field rc-distinction";
    const k = document.createElement("span");
    k.className = "rc-label";
    k.textContent = "vs " + r.distinction.vs;
    d.appendChild(k);
    if (r.distinction.difference) {
      const diff = document.createElement("div");
      diff.className = "rc-value";
      diff.textContent = r.distinction.difference;
      d.appendChild(diff);
    }
    if (r.distinction.scenario) {
      const sc = document.createElement("div");
      sc.className = "rc-distinction-scenario";
      sc.textContent = r.distinction.scenario;
      d.appendChild(sc);
    }
    details.appendChild(d);
  }
  if (r.diagram) {
    const dwrap = document.createElement("div");
    dwrap.className = "rc-field rc-secondary";
    const k = document.createElement("span");
    k.className = "rc-label";
    k.textContent = "Diagram";
    const pre = document.createElement("pre");
    pre.className = "rc-diagram";
    pre.textContent = r.diagram; // textContent: keep model output out of the DOM as HTML
    dwrap.append(k, pre);
    details.appendChild(dwrap);
  }
  if (Array.isArray(r.related_terms) && r.related_terms.length) {
    const rt = document.createElement("div");
    rt.className = "rc-field rc-related";
    const k = document.createElement("span");
    k.className = "rc-label";
    k.textContent = "Anchors";
    rt.appendChild(k);
    r.related_terms
      .filter((t) => t && t.term)
      .forEach((t) => {
        const row = document.createElement("div");
        row.className = "rc-related-row";
        const term = document.createElement("span");
        term.className = "rc-related-term";
        term.textContent = t.term;
        const rel = document.createElement("span");
        rel.className = "rc-related-rel";
        rel.textContent = t.relationship ? " — " + t.relationship : "";
        row.append(term, rel);
        rt.appendChild(row);
      });
    details.appendChild(rt);
  }

  const actions = document.createElement("div");
  actions.className = "rc-actions";

  if (details.children.length > 0) {
    const moreBtn = document.createElement("button");
    moreBtn.className = "rc-more";
    moreBtn.textContent = "More ▾";
    moreBtn.addEventListener("click", () => {
      const open = details.style.display !== "none";
      details.style.display = open ? "none" : "block";
      moreBtn.textContent = open ? "More ▾" : "Less ▴";
    });
    actions.appendChild(moreBtn);
  }

  const saveBtn = document.createElement("button");
  saveBtn.className = "rc-save";
  saveBtn.textContent = "Save";
  saveBtn.addEventListener("click", () => saveEntry(r, saveBtn));
  actions.appendChild(saveBtn);

  wrap.appendChild(details);
  wrap.appendChild(actions);

  return wrap;
}

function errorNode(message) {
  const el = document.createElement("div");
  el.className = "rc-error";
  el.textContent = message;
  return el;
}
