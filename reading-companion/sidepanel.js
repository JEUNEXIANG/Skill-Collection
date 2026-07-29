// Side-panel dictionary: shows the saved words for the article in the active
// tab. Items expire 7 days after they were saved.

const SEVEN_DAYS = 7 * 24 * 60 * 60 * 1000;

const els = {
  title: document.getElementById("article-title"),
  count: document.getElementById("count"),
  empty: document.getElementById("empty"),
  list: document.getElementById("list"),
};

let currentKey = null;

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

// Drop expired items; persist the cleanup. Returns the live items.
async function pruneBucket(key, bucket) {
  const now = Date.now();
  const live = (bucket.items || []).filter((it) => now - it.savedAt < SEVEN_DAYS);
  if (live.length !== (bucket.items || []).length) {
    if (live.length === 0) {
      await chrome.storage.local.remove(key);
    } else {
      await chrome.storage.local.set({ [key]: { ...bucket, items: live } });
    }
  }
  return live;
}

async function render() {
  const tab = await getActiveTab();
  const url = tab?.url || "";
  const key = url ? "rc:lookups:" + url : null;
  currentKey = key;

  if (!key) {
    els.title.textContent = "—";
    els.count.textContent = "";
    els.list.replaceChildren();
    els.empty.style.display = "block";
    return;
  }

  const store = await chrome.storage.local.get(key);
  const bucket = store[key];
  const items = bucket ? await pruneBucket(key, bucket) : [];

  els.title.textContent = bucket?.title || tab.title || url;
  els.count.textContent = items.length
    ? `${items.length} word${items.length === 1 ? "" : "s"}`
    : "";
  els.empty.style.display = items.length ? "none" : "block";

  els.list.replaceChildren(
    ...items
      .slice()
      .reverse() // newest first
      .map((it) => itemRow(key, it))
  );
}

function line(cls, text) {
  const el = document.createElement("div");
  el.className = cls;
  el.textContent = text;
  return el;
}

function itemRow(key, it) {
  const li = document.createElement("li");

  const remove = document.createElement("button");
  remove.className = "remove";
  remove.textContent = "×";
  remove.title = "Remove";
  remove.addEventListener("click", () => removeItem(key, it.savedAt));
  li.appendChild(remove);

  const head = document.createElement("div");
  const sel = document.createElement("span");
  sel.className = "sel";
  sel.textContent = it.selection;
  head.appendChild(sel);
  if (it.pronunciation) {
    const pron = document.createElement("span");
    pron.className = "pron";
    pron.textContent = it.pronunciation;
    head.appendChild(pron);
  }
  li.appendChild(head);

  if (it.meaning_in_context) li.appendChild(line("meaning", it.meaning_in_context));
  if (it.context) li.appendChild(line("context", it.context));

  return li;
}

async function removeItem(key, savedAt) {
  const store = await chrome.storage.local.get(key);
  const bucket = store[key];
  if (!bucket) return;
  const items = (bucket.items || []).filter((it) => it.savedAt !== savedAt);
  if (items.length === 0) {
    await chrome.storage.local.remove(key);
  } else {
    await chrome.storage.local.set({ [key]: { ...bucket, items } });
  }
  // render() runs via the storage.onChanged listener below.
}

// Refresh when the saved data changes...
chrome.storage.onChanged.addListener((changes) => {
  if (!currentKey || currentKey in changes) render();
});

// ...and when the user switches to a different tab / page.
chrome.tabs.onActivated.addListener(render);
chrome.tabs.onUpdated.addListener((_id, info, tab) => {
  if (tab.active && (info.status === "complete" || info.url)) render();
});

render();