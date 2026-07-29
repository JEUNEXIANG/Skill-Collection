const els = {
  apiKey: document.getElementById("apiKey"),
  model: document.getElementById("model"),
  nativeLanguage: document.getElementById("nativeLanguage"),
  save: document.getElementById("save"),
  status: document.getElementById("status"),
};

// Default native language from the browser locale on first run.
function defaultLanguage() {
  try {
    return (
      new Intl.DisplayNames([navigator.language], { type: "language" }).of(
        navigator.language.split("-")[0]
      ) || "English"
    );
  } catch {
    return "English";
  }
}

async function load() {
  const { apiKey, model, nativeLanguage } = await chrome.storage.local.get([
    "apiKey",
    "model",
    "nativeLanguage",
  ]);
  els.apiKey.value = apiKey || "";
  els.model.value = model || "deepseek-chat";
  els.nativeLanguage.value = nativeLanguage || defaultLanguage();
}

els.save.addEventListener("click", async () => {
  await chrome.storage.local.set({
    apiKey: els.apiKey.value.trim(),
    model: els.model.value.trim() || "deepseek-chat",
    nativeLanguage: els.nativeLanguage.value.trim() || "English",
  });
  els.status.textContent = "Saved ✓";
  setTimeout(() => (els.status.textContent = ""), 1500);
});

load();
