// viewer.js — fully automatic Word-style viewer.
// Reads xlsx bytes from chrome.storage.local (set by the background after
// the Excel button is auto-clicked), renders ALL sheets, ALL rows.
// No manual controls. Re-renders the moment new bytes arrive.

(function () {
  const subtitle = document.getElementById("subtitle");
  const body = document.getElementById("body");
  const progressWrap = document.getElementById("progress-wrap");
  const progressText = document.getElementById("progress-text");
  const progressPct = document.getElementById("progress-pct");
  const pbar = document.getElementById("pbar");
  const pbarFill = document.getElementById("pbar-fill");

  if (typeof XLSX === "undefined") {
    subtitle.innerHTML = '<span class="empty">SheetJS (libs/xlsx.full.min.js) is missing.</span>';
    return;
  }

  async function tryRender() {
    const { lastXlsxB64, lastXlsxTs } = await chrome.storage.local.get(["lastXlsxB64", "lastXlsxTs"]);
    if (!lastXlsxB64) return false;
    const buf = base64ToArrayBuffer(lastXlsxB64);
    if (!looksLikeXlsx(buf)) {
      subtitle.innerHTML = '<span class="empty">Stored data is not a valid .xlsx (zip signature missing).</span>';
      return false;
    }
    let wb;
    try { wb = XLSX.read(buf, { type: "array" }); }
    catch (e) {
      subtitle.innerHTML = '<span class="empty">Parse error: ' + escapeHtml(e.message) + '</span>';
      return false;
    }
    renderWorkbook(wb, lastXlsxTs);
    // Render succeeded — force progress to 100% so it doesn't get stuck at 75%.
    renderProgress({ text: "Done. Data displayed.", kind: "ok", progress: 100 });
    return true;
  }

  function renderWorkbook(wb, ts) {
    const sheetCount = wb.SheetNames.length;
    let totalRecords = 0;
    let html = "";

    wb.SheetNames.forEach((name, idx) => {
      const ws = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
      if (!rows.length) return;
      const dataRows = rows.length - 1;
      totalRecords += Math.max(dataRows, 0);
      html += '<div class="doc-h1">' + escapeHtml(name);
      if (sheetCount > 1) html += ' <span style="color:#888;font-weight:400;font-size:11pt">(sheet ' + (idx + 1) + '/' + sheetCount + ')</span>';
      html += '</div>';
      html += renderTable(rows);
    });

    subtitle.textContent =
      totalRecords + " record(s) across " + sheetCount + " sheet(s) \u00b7 " +
      new Date(ts || Date.now()).toLocaleString();
    body.innerHTML = html || '<div class="empty">Workbook is empty.</div>';
  }

  function renderTable(rows) {
    const header = rows[0].map((c) => String(c == null ? "" : c));
    let html = '<table class="data"><thead><tr><th class="serial">#</th>';
    header.forEach((h) => { html += '<th>' + escapeHtml(h) + '</th>'; });
    html += '</tr></thead><tbody>';
    for (let i = 1; i < rows.length; i++) {
      html += '<tr><td class="serial">' + i + '</td>';
      for (let j = 0; j < header.length; j++) {
        html += '<td>' + escapeHtml(String(rows[i][j] == null ? "" : rows[i][j])) + '</td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table>';
    return html;
  }

  function renderProgress(status) {
    if (!status) { progressWrap.classList.add("hidden"); return; }
    progressWrap.classList.remove("hidden");
    progressText.textContent = status.text || "";
    const pct = Math.max(0, Math.min(100, Number(status.progress) || 0));
    progressPct.textContent = pct + "%";
    pbarFill.style.width = pct + "%";
    pbar.classList.remove("ok", "error");
    if (status.kind === "ok") pbar.classList.add("ok");
    if (status.kind === "error") pbar.classList.add("error");
    if (pct >= 100 && status.kind === "ok") {
      setTimeout(() => progressWrap.classList.add("hidden"), 1500);
    }
  }

  // Initial paint
  (async () => {
    const rendered = await tryRender();
    if (!rendered) subtitle.textContent = "Waiting for automation to capture the data\u2026";
    const { status } = await chrome.storage.local.get(["status"]);
    renderProgress(status);
  })();

  // Auto-update whenever storage changes
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    if (changes.lastXlsxB64 && changes.lastXlsxB64.newValue) tryRender();
    if (changes.status) renderProgress(changes.status.newValue);
  });

  // Helpers
  function looksLikeXlsx(buf) {
    if (!buf || buf.byteLength < 4) return false;
    const v = new Uint8Array(buf, 0, 4);
    return v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04;
  }
  function base64ToArrayBuffer(b64) {
    const bin = atob(b64);
    const len = bin.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  }
  function escapeHtml(s) {
    return s.replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }
})();
