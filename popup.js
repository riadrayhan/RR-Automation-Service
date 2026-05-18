const $ = (id) => document.getElementById(id);

async function refreshStatus() {
  const { status } = await chrome.storage.local.get("status");
  if (!status) return;
  const el = $("status");
  el.textContent = status.text;
  el.style.color =
    status.kind === "error" ? "#dc3545" :
    status.kind === "ok" ? "#28a745" : "#555";
  const pbar = $("pbar");
  const fill = $("pbar-fill");
  pbar.classList.remove("ok", "error");
  if (status.kind === "ok") pbar.classList.add("ok");
  if (status.kind === "error") pbar.classList.add("error");
  fill.style.width = (status.progress || 0) + "%";
}

// Refresh on open + every 700ms while popup is open
refreshStatus();
const timer = setInterval(refreshStatus, 700);
window.addEventListener("unload", () => clearInterval(timer));

$("start").addEventListener("click", async () => {
  $("status").textContent = "Starting…";
  const resp = await chrome.runtime.sendMessage({ type: "START_AUTOMATION" });
  if (!resp?.ok) {
    $("status").textContent = "Error: " + (resp?.error || "unknown");
    $("status").style.color = "#dc3545";
  }
});

$("view").addEventListener("click", async () => {
  await chrome.tabs.create({ url: chrome.runtime.getURL("viewer.html") });
});

$("details").addEventListener("click", async () => {
  const id = chrome.runtime.id;
  await chrome.tabs.create({ url: "chrome://extensions/?id=" + id });
});

$("options").addEventListener("click", (e) => {
  e.preventDefault();
  chrome.runtime.openOptionsPage();
});
