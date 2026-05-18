// background.js — service worker (MV3)
// Orchestrates the automation flow and captures the xlsx download.

// Load credentials from config.js (acts as a .env). Keep config.js out of VCS.
try { importScripts("config.js"); } catch (e) { console.warn("config.js not loaded:", e); }

const TARGET_URL = "https://sahajmobile.com/customers/todayservicing";

// Supported automation flows. The popup picks one via {flow: "<key>"}.
const FLOWS = {
  todayservicing: {
    url: "https://sahajmobile.com/customers/todayservicing",
    buttonHints: ["Excel"],
    title: "Today Servicing",
    waitForData: true
  },
  online_payment_report: {
    url: "https://sahajmobile.com/reports/online_payment_report",
    buttonHints: ["Download All"],
    title: "Online Payment Report",
    waitForData: false
  }
};

// On install / startup, mirror config.js into chrome.storage.local so the rest of the code
// can keep using a single source of truth.
async function syncConfigToStorage() {
  const cfg = self.SAHAJ_CONFIG;
  if (cfg && cfg.username && cfg.password) {
    await chrome.storage.local.set({ creds: { username: cfg.username, password: cfg.password } });
  }
}
chrome.runtime.onInstalled.addListener(syncConfigToStorage);
chrome.runtime.onStartup.addListener(syncConfigToStorage);
// Also sync immediately on every SW wake-up.
syncConfigToStorage().catch(() => {});

// Clicking the extension icon starts automation immediately (no popup).
chrome.action.onClicked.addListener(async () => {
  await clearLogs();
  await appendLog("Icon clicked — starting run");
  // Detach any previous debugger session and clear stale state so re-clicks always work.
  if (runState && runState.tabId) {
    try { await chrome.debugger.detach({ tabId: runState.tabId }); } catch (_) {}
  }
  runState = null;
  const creds = await getCreds();
  if (!creds) {
    await setStatus("Credentials missing. Edit config.js.", "error", 0);
    return;
  }
  await startAutomation(creds);
});

// State for an in-flight automation run
let runState = null;
// Maps chrome.downloads id -> the flowKey it was created for. Lets us reject
// stale onChanged events from a previous flow's download once we've chained.
const downloadIdToFlow = new Map();

function setBadge(text, color = "#00c0ef") {
  chrome.action.setBadgeBackgroundColor({ color });
  chrome.action.setBadgeText({ text });
}

async function setStatus(text, kind = "info", progress) {
  console.log("[sahaj-auto][status]", kind, text, "progress:", progress);
  const cur = (await chrome.storage.local.get("status")).status || {};
  const next = {
    text,
    kind,
    ts: Date.now(),
    progress: typeof progress === "number" ? progress : (cur.progress || 0)
  };
  await chrome.storage.local.set({ status: next });
  await appendLog(text, kind);
}

async function appendLog(text, kind = "info") {
  const { logs } = await chrome.storage.local.get("logs");
  const arr = Array.isArray(logs) ? logs : [];
  arr.push({ t: Date.now(), text, kind });
  // keep last 200 entries
  while (arr.length > 200) arr.shift();
  await chrome.storage.local.set({ logs: arr });
}

async function clearLogs() {
  await chrome.storage.local.set({ logs: [] });
}

async function getCreds() {
  const { creds } = await chrome.storage.local.get("creds");
  if (creds && creds.username && creds.password) return creds;
  // Fallback to in-memory config (in case storage hasn't synced yet)
  const cfg = self.SAHAJ_CONFIG;
  if (cfg && cfg.username && cfg.password) return cfg;
  return null;
}

// Message router
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  (async () => {
    try {
      if (msg.type === "START_AUTOMATION") {
        const creds = await getCreds();
        if (!creds || !creds.username || !creds.password) {
          sendResponse({ ok: false, error: "Credentials not set. Open Options first." });
          return;
        }
        const flowKey = msg.flow || "online_payment_report";
        // After online_payment_report completes, automatically run todayservicing.
        const chainNext = flowKey === "online_payment_report" ? "todayservicing" : null;
        await startAutomation(creds, flowKey, chainNext);
        sendResponse({ ok: true });
      } else if (msg.type === "CONTENT_READY") {
        // content script announcing page load
        if (runState) handlePageReady(msg.url, sender.tab?.id).catch(console.error);
        sendResponse({ ok: true });
      } else if (msg.type === "EXCEL_BUTTON_URL") {
        // content script discovered the export URL/method
        if (runState) {
          runState.exportInfo = msg.info;
        }
        sendResponse({ ok: true });
      } else if (msg.type === "GET_LAST_ROWS") {
        const { lastRows } = await chrome.storage.local.get("lastRows");
        sendResponse({ ok: true, rows: lastRows || [] });
      } else if (msg.type === "LOG_DEBUG") {
        await appendLog(msg.text || "(empty)", msg.kind || "info");
        sendResponse({ ok: true });
      } else if (msg.type === "XLSX_BYTES") {
        // Bytes captured by the content/page hook (preferred path). Dedupe by size+ts.
        try {
          // Reject messages whose sender tab doesn't match the active flow's tab —
          // otherwise stale bytes from the previous flow's tab can overwrite the
          // new flow's viewer.
          if (!runState || (sender && sender.tab && sender.tab.id !== runState.tabId)) {
            await appendLog("Ignored XLSX_BYTES from non-active tab " +
              (sender && sender.tab && sender.tab.id), "info");
            sendResponse({ ok: true, stale: true });
            return;
          }
          if (runState.xlsxDelivered) {
            await appendLog("Ignored duplicate XLSX_BYTES (size=" + msg.size + ")", "info");
            sendResponse({ ok: true, duplicate: true });
            return;
          }
          const { lastXlsxSize, lastXlsxTs } = await chrome.storage.local.get(["lastXlsxSize", "lastXlsxTs"]);
          if (lastXlsxSize === msg.size && lastXlsxTs && Date.now() - lastXlsxTs < 15000) {
            await appendLog("Ignored duplicate XLSX_BYTES (size+recent)", "info");
            sendResponse({ ok: true, duplicate: true });
            return;
          }
          runState.xlsxDelivered = true;
          const buf = base64ToArrayBuffer(msg.b64);
          const fmt = detectReportFormat(buf) || "xlsx";
          // Do NOT null runState after deliverXlsx — chaining may have already
          // set up the next flow's runState; clearing here would clobber it.
          await deliverXlsx(msg.b64, msg.size, "page-hook", fmt);
        } catch (e) {
          await setStatus("Failed to store xlsx: " + e.message, "error");
        }
        sendResponse({ ok: true });
      }
    } catch (e) {
      console.error(e);
      sendResponse({ ok: false, error: String(e?.message || e) });
    }
  })();
  return true; // async response
});

async function startAutomation(creds, flowKey, chainNext) {
  flowKey = flowKey || "todayservicing";
  const flow = FLOWS[flowKey] || FLOWS.todayservicing;
  setBadge("…");
  // Clear stale legacy capture so dedup doesn't reject new bytes for this flow.
  try { await chrome.storage.local.remove(["lastXlsxB64", "lastXlsxSize", "lastXlsxTs", "lastXlsxFormat"]); } catch (_) {}
  // Clear stale per-flow capture so the viewer doesn't render the previous run's bytes.
  try {
    await chrome.storage.local.remove([
      `xlsx_${flowKey}`, `xlsxSize_${flowKey}`, `xlsxTs_${flowKey}`, `xlsxFormat_${flowKey}`
    ]);
  } catch (_) {}
  await chrome.storage.local.set({ flowTitle: flow.title });
  await setStatus("Opening " + flow.title + "…", "info", 5);
  const tab = await chrome.tabs.create({ url: flow.url, active: true });
  runState = {
    tabId: tab.id,
    creds,
    flow,
    flowKey,
    chainNext: chainNext || null,
    stage: "navigating",
    excelClicked: false,
    loginAttempted: false,
    downloadId: null,
    debuggerAttached: false,
    pendingResponses: new Map()
  };
  await attachDebugger(tab.id);
}

// ----- Chrome DevTools Protocol path (most reliable response capture) -----
const DEBUGGEE_VERSION = "1.3";

async function attachDebugger(tabId) {
  try {
    await new Promise((resolve, reject) => {
      chrome.debugger.attach({ tabId }, DEBUGGEE_VERSION, () => {
        const err = chrome.runtime.lastError;
        if (err) reject(new Error(err.message)); else resolve();
      });
    });
    await sendCmd(tabId, "Network.enable", {});
    // Intercept responses so we can read bodies for downloads (Fetch domain).
    await sendCmd(tabId, "Fetch.enable", {
      patterns: [{ urlPattern: "*", requestStage: "Response" }]
    });
    if (runState) runState.debuggerAttached = true;
    console.log("[sahaj-auto] debugger attached to tab", tabId);
  } catch (e) {
    console.warn("[sahaj-auto] could not attach debugger:", e && e.message);
    await setStatus("Could not attach debugger: " + e.message, "error");
  }
}

function sendCmd(tabId, method, params) {
  return new Promise((resolve, reject) => {
    chrome.debugger.sendCommand({ tabId }, method, params, (res) => {
      const err = chrome.runtime.lastError;
      if (err) reject(new Error(err.message)); else resolve(res);
    });
  });
}

async function detachDebuggerSafe(tabId) {
  return new Promise((resolve) => {
    chrome.debugger.detach({ tabId }, () => {
      // ignore lastError
      resolve();
    });
  });
}

chrome.debugger.onEvent.addListener(async (source, method, params) => {
  // CRITICAL: Fetch.requestPaused MUST always be continued, otherwise the page hangs.
  // Handle that case before any early-return checks.
  if (method === "Fetch.requestPaused") {
    const reqId = params && params.requestId;
    // If this event isn't for our active flow's tab, just continue it and bail.
    if (!runState || source.tabId !== runState.tabId) {
      try { await sendDebugCmd(source.tabId, "Fetch.continueRequest", { requestId: reqId }); } catch (_) {}
      return;
    }
  } else {
    // Non-Fetch events: keep the old guard.
    if (!runState || source.tabId !== runState.tabId) return;
  }

  try {
    if (method === "Fetch.requestPaused") {
      const reqId = params.requestId;
      const url = (params.request && params.request.url) || "";
      const status = params.responseStatusCode || 0;
      const headers = params.responseHeaders || [];
      const hdr = (n) => {
        const f = headers.find((h) => h.name && h.name.toLowerCase() === n);
        return f ? f.value : "";
      };
      const ct = hdr("content-type");
      const cd = hdr("content-disposition");
      // Broad URL/CT match — the actual content is verified by detectReportFormat,
      // which rejects HTML/CSS/XML/JS payloads even if their URL looked download-y.
      const looksXlsx =
        /\.xlsx?(\?|#|$)/i.test(url) ||
        /\.csv(\?|#|$)/i.test(url) ||
        /export|download/i.test(url) ||
        /spreadsheet|excel|officedocument|text\/csv|application\/csv|application\/octet-stream/i.test(ct) ||
        /\.(xlsx?|csv)/i.test(cd) ||
        /attachment/i.test(cd);

      if (!looksXlsx) {
        // Pass through without inspection
        try { await sendDebugCmd(source.tabId, "Fetch.continueRequest", { requestId: reqId }); } catch (_) {}
        return;
      }

      console.log("[sahaj-auto][fetch] candidate:", url, "status:", status, "ct:", ct, "cd:", cd);
      let captured = false;
      let buf = null;
      let alreadyFulfilled = false;
      // (1) Try Fetch.getResponseBody (works for most text/JSON responses).
      try {
        const body = await sendCmd(source.tabId, "Fetch.getResponseBody", { requestId: reqId });
        if (body && typeof body.body === "string") {
          if (body.base64Encoded) buf = base64ToArrayBuffer(body.body);
          else buf = new TextEncoder().encode(body.body).buffer;
        }
      } catch (e) {
        console.warn("[sahaj-auto][fetch] getResponseBody failed:", e && e.message);
      }
      // (2) Fallback for binary downloads where getResponseBody returns empty —
      // stream the body via Fetch.takeResponseBodyAsStream + IO.read.
      if (!buf || buf.byteLength === 0) {
        try {
          const stream = await sendCmd(source.tabId, "Fetch.takeResponseBodyAsStream", { requestId: reqId });
          if (stream && stream.stream) {
            const handle = stream.stream;
            const chunks = [];
            let total = 0;
            while (true) {
              const r = await sendCmd(source.tabId, "IO.read", { handle, size: 1 << 16 });
              if (!r || !r.data) break;
              const part = r.base64Encoded
                ? new Uint8Array(base64ToArrayBuffer(r.data))
                : new TextEncoder().encode(r.data);
              chunks.push(part);
              total += part.byteLength;
              if (r.eof) break;
            }
            try { await sendCmd(source.tabId, "IO.close", { handle }); } catch (_) {}
            if (total > 0) {
              const merged = new Uint8Array(total);
              let off = 0;
              for (const c of chunks) { merged.set(c, off); off += c.byteLength; }
              buf = merged.buffer;
              console.log("[sahaj-auto][fetch] streamed body via IO.read, bytes=", total);
              // NOTE: taking the stream consumes the response. We must fulfillRequest
              // to deliver something back to the page; otherwise the browser may stall.
              try {
                const b64Body = arrayBufferToBase64(buf);
                await sendCmd(source.tabId, "Fetch.fulfillRequest", {
                  requestId: reqId,
                  responseCode: status || 200,
                  responseHeaders: headers,
                  body: b64Body
                });
                alreadyFulfilled = true;
              } catch (e) {
                console.warn("[sahaj-auto][fetch] fulfillRequest after stream failed:", e && e.message);
              }
            }
          }
        } catch (e) {
          console.warn("[sahaj-auto][fetch] takeResponseBodyAsStream failed:", e && e.message);
        }
      }

      if (buf && buf.byteLength) {
        const fmt = detectReportFormat(buf);
        if (fmt && !runState.xlsxDelivered) {
          runState.xlsxDelivered = true;
          captured = true;
          const b64 = arrayBufferToBase64(buf);
          await deliverXlsx(b64, buf.byteLength, "cdp-fetch", fmt);
        } else {
          console.log("[sahaj-auto][fetch] body not a recognized report, byteLen:", buf.byteLength);
        }
      } else {
        console.log("[sahaj-auto][fetch] no body could be read");
      }

      // Continue only if we didn't already fulfill via the streaming path.
      if (!alreadyFulfilled) {
        try { await sendDebugCmd(source.tabId, "Fetch.continueRequest", { requestId: reqId }); } catch (_) {}
      }

      if (captured) {
        // Wind down: detach debugger from the captured tab after a brief delay.
        setTimeout(() => detachDebuggerSafe(source.tabId), 500);
      }
    }
  } catch (e) {
    console.error("[sahaj-auto][cdp] handler error:", e);
  }
});

// sendCmd that never throws, used purely for fire-and-forget continueRequest.
function sendDebugCmd(tabId, method, params) {
  return new Promise((resolve) => {
    chrome.debugger.sendCommand({ tabId }, method, params, () => {
      void chrome.runtime.lastError;
      resolve();
    });
  });
}

chrome.debugger.onDetach.addListener((source, reason) => {
  console.log("[sahaj-auto] debugger detached:", source, "reason:", reason);
  if (runState && source.tabId === runState.tabId) runState.debuggerAttached = false;
});

function base64ToArrayBuffer(b64) {
  const bin = atob(b64);
  const len = bin.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

// Detects "xlsx" (ZIP magic) or "csv" (printable text). Returns null otherwise.
// Explicitly rejects HTML, XML, CSS, JS — these are common false positives because
// they're also printable text but they are NOT the report.
function detectReportFormat(buf) {
  if (!buf || buf.byteLength < 2) return null;
  const v = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
  if (v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04) return "xlsx";
  // Look at up to 1024 bytes for content sniffing.
  const sniffLen = Math.min(1024, buf.byteLength);
  const sample = new Uint8Array(buf, 0, sniffLen);
  // Reject obvious non-CSV text payloads (HTML, XML, CSS @font-face, JS).
  let head = "";
  for (let i = 0; i < Math.min(256, sample.length); i++) head += String.fromCharCode(sample[i]);
  const trimmed = head.replace(/^\uFEFF/, "").trimStart().toLowerCase();
  if (
    trimmed.startsWith("<") ||
    trimmed.startsWith("<!doctype") ||
    trimmed.startsWith("<html") ||
    trimmed.startsWith("<?xml") ||
    trimmed.startsWith("@font-face") ||
    trimmed.startsWith("@import") ||
    trimmed.startsWith("@charset") ||
    trimmed.startsWith("/*") ||
    /^\s*(function|var|const|let|import|export)\s/.test(trimmed)
  ) return null;
  let printable = 0;
  for (let i = 0; i < sample.length; i++) {
    const c = sample[i];
    if (c === 9 || c === 10 || c === 13 || (c >= 32 && c <= 126) || c >= 160) printable++;
  }
  if (printable / sample.length < 0.92) return null;
  // Reasonable CSV has at least one comma or newline in the first 1KB.
  if (!/[,\n;\t|]/.test(head)) return null;
  return "csv";
}

async function focusOrOpenViewer(flowKey) {
  // Always open a fresh viewer tab on successful capture; the flow key is on the URL hash.
  const url = chrome.runtime.getURL("viewer.html") + (flowKey ? "#" + flowKey : "");
  const t = await chrome.tabs.create({ url, active: true });
  try { await chrome.windows.update(t.windowId, { focused: true }); } catch (_) {}
  return t.id;
}

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

/**
 * Deliver captured xlsx bytes to a fresh viewer tab. Each flow gets its own
 * storage key (xlsx_<flowKey>) so multiple flows can be displayed independently.
 */
async function deliverXlsx(b64, size, sourceLabel, format) {
  format = format || "xlsx";
  const flowKey = runState && runState.flow ? runState.flowKey : "todayservicing";
  const flowTitle = runState && runState.flow ? runState.flow.title : "Result";
  await setStatus("Download captured (" + size + " bytes, " + format + "). Opening viewer…", "info", 30);
  await focusOrOpenViewer(flowKey);
  await sleep(400);
  await setStatus("Preparing data…", "info", 50);
  await sleep(250);
  await setStatus("Parsing " + format.toUpperCase() + "…", "info", 70);
  await sleep(200);
  const payload = {
    [`xlsx_${flowKey}`]: b64,
    [`xlsxSize_${flowKey}`]: size,
    [`xlsxTs_${flowKey}`]: Date.now(),
    [`xlsxFormat_${flowKey}`]: format,
    [`flowTitle_${flowKey}`]: flowTitle,
    lastXlsxB64: b64,
    lastXlsxSize: size,
    lastXlsxTs: Date.now(),
    lastXlsxFormat: format,
    flowTitle
  };
  await chrome.storage.local.set(payload);
  await setStatus("Rendering data…", "info", 90);
  await sleep(150);
  setBadge("OK", "#28a745");
  await setStatus(flowTitle + " ready.", "ok", 100);

  // Chain: after this flow succeeds, automatically run the next one — but only
  // AFTER the previous flow has fully wound down (debugger detached, any pending
  // download.onChanged settled). Strict serial: flow 1 100% done → flow 2 starts.
  if (runState && runState.chainNext) {
    const nextFlowKey = runState.chainNext;
    const nextFlow = FLOWS[nextFlowKey];
    const prevTabId = runState.tabId;
    const prevDownloadId = runState.downloadId;

    // Detach the previous tab's debugger NOW so its Fetch.requestPaused events
    // can no longer block anything.
    try { await detachDebuggerSafe(prevTabId); } catch (_) {}

    // Wait up to 4s for the previous browser download (if any) to reach
    // "complete" state so its stale onChanged fires & is rejected by the
    // downloadIdToFlow check while we're still on the OLD flow.
    if (prevDownloadId != null) {
      const deadline = Date.now() + 4000;
      while (Date.now() < deadline) {
        try {
          const items = await chrome.downloads.search({ id: prevDownloadId });
          const st = items && items[0] && items[0].state;
          if (st === "complete" || st === "interrupted") break;
        } catch (_) { break; }
        await sleep(200);
      }
      // Allow the onChanged grace-period (1.2s) to elapse so the stale-rejection
      // path runs to completion before we mutate runState.
      await sleep(1400);
    } else {
      await sleep(500);
    }

    runState = null;
    await appendLog("Previous flow fully wound down — chaining to: " + nextFlowKey);
    const creds = await getCreds();
    if (creds && nextFlow) await startAutomation(creds, nextFlowKey, null);
  }
}

// Drive the flow off tab navigation events — more reliable than relying on content script timing.
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (!runState || tabId !== runState.tabId) return;
  if (changeInfo.status !== "complete" || !tab.url) return;
  handlePageReady(tab.url, tabId).catch((e) => console.error("handlePageReady:", e));
});

async function handlePageReady(url, tabId) {
  if (!runState || tabId !== runState.tabId) return;
  console.log("[sahaj-auto] page ready:", url, "stage:", runState.stage);

  // On target page → click flow button (once).
  const flowUrl = runState.flow ? runState.flow.url : TARGET_URL;
  if (url.startsWith(flowUrl)) {
    if (runState.excelClicked) return;
    runState.excelClicked = true;
    runState.stage = "clicking-excel";
    await setStatus("On target page. Waiting for data to load…", "info", 55);
    setTimeout(async () => {
      try {
        const res = await chrome.tabs.sendMessage(tabId, {
          type: "CLICK_EXCEL",
          buttonHints: runState.flow ? runState.flow.buttonHints : ["Excel"],
          waitForData: runState.flow ? runState.flow.waitForData !== false : true
        });
        if (res?.ok) await setStatus("Button clicked. Waiting for download…", "info", 75);
        else await setStatus("Export button NOT found on page.", "error");
      } catch (e) {
        await setStatus("Could not message content script: " + e.message, "error");
      }
    }, 1200);
    return;
  }

  // Otherwise we're somewhere else (login page, dashboard, etc.).
  if (!runState.loginAttempted) {
    runState.loginAttempted = true;
    runState.stage = "logging-in";
    await setStatus("On " + new URL(url).pathname + " — attempting login…", "info", 20);
    try {
      const res = await chrome.tabs.sendMessage(tabId, { type: "DO_LOGIN", creds: runState.creds });
      if (res?.ok) await setStatus("Login submitted. Waiting for redirect…", "info", 35);
      else await setStatus("Login form not found on page.", "error");
    } catch (e) {
      await setStatus("Could not message content script: " + e.message, "error");
    }
    return;
  }

  // Login already attempted but we still aren't on target. Force-navigate.
  if (runState.stage !== "forcing-target") {
    runState.stage = "forcing-target";
    await setStatus("Logged in. Navigating to target page…", "info", 50);
    chrome.tabs.update(tabId, { url: runState.flow ? runState.flow.url : TARGET_URL });
  }
}

// Capture the xlsx download triggered by the Excel button
chrome.downloads.onCreated.addListener(async (item) => {
  await appendLog("download.onCreated: url=" + item.url + " mime=" + item.mime + " fn=" + item.filename);
  if (!runState) { await appendLog("(no runState — ignoring)", "info"); return; }
  const isReport = (item.filename && /\.(xlsx?|csv)$/i.test(item.filename)) ||
                   (item.mime && /spreadsheet|excel|csv|text\/plain/i.test(item.mime)) ||
                   (item.url && /\.(xlsx?|csv)(\?|$)/i.test(item.url)) ||
                   (item.url && /excel|export|download/i.test(item.url));
  if (!isReport) { await appendLog("download did not match report heuristics", "info"); return; }

  // If bytes already captured (via CDP/hook/refetch), just record the download
  // for tracking — but DO NOT cancel: the user wants the actual file on disk.
  const { lastXlsxB64 } = await chrome.storage.local.get("lastXlsxB64");
  if (lastXlsxB64 || runState.xlsxDelivered) {
    runState.downloadId = item.id;
    downloadIdToFlow.set(item.id, runState.flowKey);
    await appendLog("Browser download running alongside captured bytes (id=" + item.id + ") — letting it save");
    return;
  }

  runState.downloadId = item.id;
  runState.downloadUrl = item.url;
  downloadIdToFlow.set(item.id, runState.flowKey);
  await appendLog("xlsx download detected (id=" + item.id + ", flow=" + runState.flowKey + ")");

  // Refetch URL from page context (has session cookies).
  if (!runState.xlsxDelivered) {
    try {
      await chrome.scripting.executeScript({
        target: { tabId: runState.tabId },
        world: "MAIN",
        func: async (url) => {
          try {
            const r = await fetch(url, { credentials: "include" });
            if (!r.ok) { console.warn("[sahaj-refetch] !ok", r.status); return; }
            const buf = await r.arrayBuffer();
            const v = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
            const isZip = v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04;
            // CSV check: ≥92% printable in first 512 bytes.
            let isCsv = false;
            if (!isZip) {
              const sample = new Uint8Array(buf, 0, Math.min(512, buf.byteLength));
              let p = 0;
              for (let i = 0; i < sample.length; i++) {
                const c = sample[i];
                if (c === 9 || c === 10 || c === 13 || (c >= 32 && c <= 126) || c >= 160) p++;
              }
              isCsv = sample.length > 0 && (p / sample.length) >= 0.92;
            }
            if (!isZip && !isCsv) {
              console.warn("[sahaj-refetch] not xlsx/csv");
              return;
            }
            const bytes = new Uint8Array(buf);
            let bin = "";
            const chunk = 0x8000;
            for (let i = 0; i < bytes.length; i += chunk) {
              bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
            }
            const b64 = btoa(bin);
            window.postMessage({ __sahaj: true, type: "XLSX_BYTES", b64, size: buf.byteLength, url }, "*");
          } catch (e) { console.warn("[sahaj-refetch] failed", e); }
        },
        args: [item.url]
      });
      await appendLog("page-context refetch dispatched");
    } catch (e) {
      await appendLog("page-context refetch dispatch failed: " + (e && e.message), "error");
    }
  }
});

chrome.downloads.onChanged.addListener(async (delta) => {
  if (delta.state && delta.state.current) {
    await appendLog("download.onChanged state=" + delta.state.current + " id=" + delta.id);
  }
  if (!(delta.state && delta.state.current === "complete")) {
    if (delta.error) {
      setBadge("ERR", "#dc3545");
      await setStatus("Download error: " + (delta.error.current || "unknown"), "error");
    }
    return;
  }

  // Identify which flow this download belongs to. Reject anything that isn't
  // the active flow's download — otherwise a late completion from the previous
  // flow would overwrite the current flow's viewer with stale bytes.
  const ownerFlow = downloadIdToFlow.get(delta.id);
  if (!ownerFlow) {
    await appendLog("onChanged: download id=" + delta.id + " has no tracked flow — ignoring", "info");
    return;
  }
  if (!runState || runState.flowKey !== ownerFlow) {
    await appendLog("onChanged: stale download id=" + delta.id + " (flow=" + ownerFlow +
      ") doesn't match active flow=" + (runState && runState.flowKey) + " — ignoring", "info");
    downloadIdToFlow.delete(delta.id);
    return;
  }

  // Give CDP / page-hook / refetch paths a brief grace period to deliver bytes
  // before we fall back to reading from disk.
  await sleep(1200);
  // Re-check active flow after the sleep (chaining may have advanced).
  if (!runState || runState.flowKey !== ownerFlow) {
    await appendLog("onChanged: flow advanced during grace period — ignoring id=" + delta.id, "info");
    downloadIdToFlow.delete(delta.id);
    return;
  }
  if (runState.xlsxDelivered) {
    await appendLog("onChanged: bytes already delivered for flow=" + ownerFlow + " — disk read skipped");
    downloadIdToFlow.delete(delta.id);
    return;
  }
  try {
    await setStatus("Download complete. Reading file from disk…", "info", 85);
    const buf = await readDownloadedFile(delta.id);
    if (!buf) {
      await appendLog(
        "Disk read failed. If no viewer opened, enable 'Allow access to file URLs' on this extension.",
        "info"
      );
      downloadIdToFlow.delete(delta.id);
      return;
    }
    // Final guard before writing: confirm we're still on the same flow.
    if (!runState || runState.flowKey !== ownerFlow) {
      await appendLog("onChanged: flow advanced during disk read — dropping bytes for id=" + delta.id, "info");
      downloadIdToFlow.delete(delta.id);
      return;
    }
    const b64 = arrayBufferToBase64(buf);
    await appendLog("Disk read OK — " + buf.byteLength + " bytes");
    const fmt = detectReportFormat(buf) || "xlsx";
    runState.xlsxDelivered = true;
    await deliverXlsx(b64, buf.byteLength, "disk", fmt);
  } catch (e) {
    console.error(e);
    setBadge("ERR", "#dc3545");
    await setStatus(e.message, "error");
  } finally {
    downloadIdToFlow.delete(delta.id);
  }
});

async function readDownloadedFile(downloadId) {
  // Look up the saved filename on disk and fetch via file:// (requires
  // "Allow access to file URLs" toggle on the extension card).
  const items = await chrome.downloads.search({ id: downloadId });
  if (!items || !items.length || !items[0].filename) {
    console.warn("[sahaj-auto] downloads.search returned no filename");
    return null;
  }
  const raw = items[0].filename;
  console.log("[sahaj-auto] downloaded file path:", raw);

  // Build a file:// URL. Properly encode each path segment.
  const norm = raw.replace(/\\/g, "/");
  const parts = norm.split("/").map((seg) => encodeURIComponent(seg));
  // Re-join; Windows drive letter "C:" becomes "C%3A" after encodeURIComponent, fix it.
  if (/^[A-Za-z]%3A$/i.test(parts[0])) {
    parts[0] = parts[0].replace(/%3A/i, ":");
  }
  const joined = parts.join("/");
  const fileUrl = norm.startsWith("/") ? "file://" + joined : "file:///" + joined;
  console.log("[sahaj-auto] fetching", fileUrl);

  try {
    const resp = await fetch(fileUrl);
    if (!resp.ok) {
      console.warn("[sahaj-auto] file fetch !ok:", resp.status);
      return null;
    }
    const buf = await resp.arrayBuffer();
    if (!detectReportFormat(buf)) {
      console.warn("[sahaj-auto] file did not look like xlsx or csv");
      return null;
    }
    return buf;
  } catch (e) {
    console.warn("[sahaj-auto] file:// fetch threw:", e && e.message);
    return null;
  }
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}
