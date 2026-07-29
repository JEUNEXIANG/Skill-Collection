// Service worker: receives lookup requests from the content script, calls the
// user's chosen LLM provider, and returns the structured explanation.
//
// v0.1 (walking skeleton): single provider — DeepSeek (OpenAI-compatible API).
// The provider adapter is shaped so OpenAI / OpenRouter / local models can be
// added later by extending PROVIDERS.

const PROVIDERS = {
  deepseek: {
    url: "https://api.deepseek.com/chat/completions",
    defaultModel: "deepseek-chat",
  },
};

const MAX_ARTICLE_CHARS = 12000; // keep token cost bounded for the skeleton

async function loadPromptTemplate() {
  const res = await fetch(chrome.runtime.getURL("prompt-template.txt"));
  return res.text();
}

function fillTemplate(tpl, vars) {
  return tpl
    .replaceAll("{native_language}", vars.native_language)
    .replaceAll("{article_title}", vars.article_title)
    .replaceAll("{article_text}", vars.article_text)
    .replaceAll("{context_sentence}", vars.context_sentence)
    .replaceAll("{selection}", vars.selection);
}

async function explain({ selection, contextSentence, articleText, title }) {
  const { apiKey, nativeLanguage, model } = await chrome.storage.local.get([
    "apiKey",
    "nativeLanguage",
    "model",
  ]);

  if (!apiKey) {
    throw new Error(
      "No API key set. Open the extension's options page and add your DeepSeek API key."
    );
  }

  const tpl = await loadPromptTemplate();
  const prompt = fillTemplate(tpl, {
    native_language: nativeLanguage || "English",
    article_title: title || "",
    article_text: (articleText || "").slice(0, MAX_ARTICLE_CHARS),
    context_sentence: contextSentence || "",
    selection,
  });

  const resp = await fetch(PROVIDERS.deepseek.url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: model || PROVIDERS.deepseek.defaultModel,
      messages: [{ role: "user", content: prompt }],
      response_format: { type: "json_object" },
      temperature: 0.3,
    }),
  });

  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Provider error ${resp.status}: ${body.slice(0, 300)}`);
  }

  const data = await resp.json();
  const content = data.choices?.[0]?.message?.content ?? "{}";
  try {
    return JSON.parse(content);
  } catch {
    throw new Error("Provider returned invalid JSON.");
  }
}

// Clicking the toolbar icon opens the side-panel dictionary.
chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel
    .setPanelBehavior({ openPanelOnActionClick: true })
    .catch(() => {});
});

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg?.type === "explain") {
    explain(msg.payload)
      .then((result) => sendResponse({ ok: true, result }))
      .catch((err) => sendResponse({ ok: false, error: String(err.message || err) }));
    return true; // keep the message channel open for the async response
  }
});
