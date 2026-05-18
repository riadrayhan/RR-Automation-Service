const $ = (id) => document.getElementById(id);

(async () => {
  const { creds } = await chrome.storage.local.get("creds");
  if (creds) {
    $("u").value = creds.username || "";
    $("p").value = creds.password || "";
  }
})();

$("save").addEventListener("click", async () => {
  const username = $("u").value.trim();
  const password = $("p").value;
  if (!username || !password) {
    $("msg").textContent = "Both fields are required.";
    $("msg").style.color = "#a00";
    return;
  }
  await chrome.storage.local.set({ creds: { username, password } });
  $("msg").textContent = "Saved.";
  $("msg").style.color = "#28a745";
});
