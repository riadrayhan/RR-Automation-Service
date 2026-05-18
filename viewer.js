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

  // Each viewer tab is locked to a specific flow via URL hash (e.g. viewer.html#todayservicing).
  const flowKey = (location.hash || "").replace(/^#/, "") || null;
  const xlsxKey = flowKey ? "xlsx_" + flowKey : "lastXlsxB64";
  const tsKey = flowKey ? "xlsxTs_" + flowKey : "lastXlsxTs";
  const titleKey = flowKey ? "flowTitle_" + flowKey : "flowTitle";
  const formatKey = flowKey ? "xlsxFormat_" + flowKey : "lastXlsxFormat";

  async function tryRender() {
    const data = await chrome.storage.local.get([xlsxKey, tsKey, formatKey]);
    const b64 = data[xlsxKey];
    const ts = data[tsKey];
    let format = data[formatKey] || "";
    if (!b64) return false;
    const buf = base64ToArrayBuffer(b64);

    // Auto-detect if format wasn't stored or is wrong.
    if (!format) {
      const v = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
      const isZip = v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04;
      format = isZip ? "xlsx" : "csv";
    }

    let wb;
    let parseErr = null;
    try {
      if (format === "csv") {
        let text = new TextDecoder("utf-8").decode(new Uint8Array(buf));
        if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
        // First try SheetJS auto-detect (handles ,  ;  tab  | delimiters).
        wb = XLSX.read(text, { type: "string", raw: true });
        // If SheetJS gave us only one column or empty, do a manual fallback.
        const firstSheet = wb.SheetNames[0] && wb.Sheets[wb.SheetNames[0]];
        const sampleRows = firstSheet ? XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" }) : [];
        const looksSingleCol = sampleRows.length > 0 && sampleRows.every((r) => r.length <= 1);
        if (!sampleRows.length || looksSingleCol) {
          wb = parseCsvManual(text);
        }
      } else {
        wb = XLSX.read(buf, { type: "array" });
      }
    } catch (e) {
      parseErr = e;
    }

    if (parseErr || !wb || !wb.SheetNames.length) {
      // Show diagnostic so user can see what was actually captured.
      const previewText = (() => {
        try {
          let t = new TextDecoder("utf-8").decode(new Uint8Array(buf, 0, Math.min(800, buf.byteLength)));
          if (t.charCodeAt(0) === 0xFEFF) t = t.slice(1);
          return t;
        } catch (_) { return "(non-text bytes)"; }
      })();
      subtitle.innerHTML = '<span class="empty">Parse failed (' + escapeHtml((parseErr && parseErr.message) || "empty workbook") +
        '). Format=' + escapeHtml(format) + ', bytes=' + buf.byteLength + '.</span>';
      body.innerHTML = '<div class="doc-h1">Captured raw content (first 800 chars)</div>' +
        '<pre style="white-space:pre-wrap;word-break:break-all;background:#f7f7f7;padding:12px;border:1px solid #ddd;border-radius:4px;font-family:Consolas,monospace;font-size:11pt">' +
        escapeHtml(previewText) + '</pre>';
      return false;
    }

    renderWorkbook(wb, ts);
    renderProgress({ text: "Done. Data displayed.", kind: "ok", progress: 100 });
    return true;
  }

  // Manual CSV parser fallback — handles quoted fields, escaped quotes, common delimiters.
  function parseCsvManual(text) {
    // Detect delimiter from the first line.
    const firstNl = text.indexOf("\n");
    const firstLine = firstNl >= 0 ? text.slice(0, firstNl) : text;
    const counts = {
      ",": (firstLine.match(/,/g) || []).length,
      ";": (firstLine.match(/;/g) || []).length,
      "\t": (firstLine.match(/\t/g) || []).length,
      "|": (firstLine.match(/\|/g) || []).length
    };
    let delim = ",";
    let best = -1;
    for (const k in counts) if (counts[k] > best) { best = counts[k]; delim = k; }

    const rows = [];
    let cur = [];
    let field = "";
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else inQuotes = false;
        } else {
          field += c;
        }
      } else {
        if (c === '"') inQuotes = true;
        else if (c === delim) { cur.push(field); field = ""; }
        else if (c === "\n") { cur.push(field); field = ""; rows.push(cur); cur = []; }
        else if (c === "\r") { /* skip */ }
        else field += c;
      }
    }
    if (field.length || cur.length) { cur.push(field); rows.push(cur); }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "CSV");
    return wb;
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
    const data = await chrome.storage.local.get([titleKey]);
    const flowTitle = data[titleKey];
    if (flowTitle) {
      const t = document.getElementById("doc-title");
      if (t) t.textContent = flowTitle;
      document.title = flowTitle + " — RR Automation Service";
    }
    const rendered = await tryRender();
    if (!rendered) subtitle.textContent = "Waiting for automation to capture the data\u2026";
    const { status } = await chrome.storage.local.get(["status"]);
    renderProgress(status);
  })();

  // Auto-update whenever storage changes
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    if (changes[xlsxKey] && changes[xlsxKey].newValue) tryRender();
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
