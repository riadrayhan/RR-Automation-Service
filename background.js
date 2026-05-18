// background.js — service worker (MV3)
// Orchestrates the automation flow and captures the xlsx download.

// Load credentials from config.js (acts as a .env). Keep config.js out of VCS.
try { importScripts("config.js"); } catch (e) { console.warn("config.js not loaded:", e); }

const TARGET_URL = "https://sahajmobile.com/customers/todayservicing";

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
        await startAutomation(creds);
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
          if (runState && runState.xlsxDelivered) {
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
          if (runState) runState.xlsxDelivered = true;
          await deliverXlsx(msg.b64, msg.size, "page-hook");
          runState = null;
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

async function startAutomation(creds) {
  setBadge("…");
  // Clear any stale capture from a previous run so dedup doesn't reject new bytes.
  try { await chrome.storage.local.remove(["lastXlsxB64", "lastXlsxSize", "lastXlsxTs"]); } catch (_) {}
  await setStatus("Opening target page…", "info", 5);
  // Go straight to the target page. If not authenticated, the site will redirect to its login page.
  const tab = await chrome.tabs.create({ url: TARGET_URL, active: true });
  runState = {
    tabId: tab.id,
    creds,
    stage: "navigating",
    excelClicked: false,
    loginAttempted: false,
    downloadId: null,
    debuggerAttached: false,
    pendingResponses: new Map() // requestId -> { url, mime }
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
  if (!runState || source.tabId !== runState.tabId) return;
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
      const looksXlsx =
        /\.xlsx?(\?|#|$)/i.test(url) ||
        /excel|export/i.test(url) ||
        /spreadsheet|excel|officedocument|application\/octet-stream/i.test(ct) ||
        /\.xlsx?/i.test(cd) ||
        /attachment/i.test(cd);

      if (!looksXlsx) {
        // Pass through without inspection
        try { await sendCmd(runState.tabId, "Fetch.continueRequest", { requestId: reqId }); } catch (_) {}
        return;
      }

      console.log("[sahaj-auto][fetch] candidate:", url, "status:", status, "ct:", ct, "cd:", cd);
      let captured = false;
      try {
        const body = await sendCmd(runState.tabId, "Fetch.getResponseBody", { requestId: reqId });
        if (body && typeof body.body === "string") {
          let buf;
          if (body.base64Encoded) buf = base64ToArrayBuffer(body.body);
          else buf = new TextEncoder().encode(body.body).buffer;

          const v = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
          const isZip = v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04;
          if (isZip && !runState.xlsxDelivered) {
            runState.xlsxDelivered = true;
            captured = true;
            const b64 = arrayBufferToBase64(buf);
            await deliverXlsx(b64, buf.byteLength, "cdp-fetch");
          } else {
            console.log("[sahaj-auto][fetch] body not xlsx-zip, byteLen:", buf.byteLength);
          }
        } else {
          console.log("[sahaj-auto][fetch] empty body returned");
        }
      } catch (e) {
        console.warn("[sahaj-auto][fetch] getResponseBody failed:", e && e.message);
      }
      try { await sendCmd(runState.tabId, "Fetch.continueRequest", { requestId: reqId }); } catch (_) {}

      if (captured) {
        // Wind down: detach debugger after a brief delay
        const tid = runState.tabId;
        setTimeout(() => detachDebuggerSafe(tid), 500);
        runState = null;
      }
    }
  } catch (e) {
    console.error("[sahaj-auto][cdp] handler error:", e);
  }
});

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

async function focusOrOpenViewer() {
  // Always open a fresh viewer tab on successful capture.
  const url = chrome.runtime.getURL("viewer.html");
  const t = await chrome.tabs.create({ url, active: true });
  try { await chrome.windows.update(t.windowId, { focused: true }); } catch (_) {}
  return t.id;
}

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

/**
 * Deliver captured xlsx bytes to a fresh viewer tab with a visible staged
 * progress animation: open tab → 40% → 60% → store bytes (triggers render) → 100%.
 */
async function deliverXlsx(b64, size, sourceLabel) {
  // 1. Announce capture and open a fresh viewer tab immediately.
  await setStatus("Download captured (" + size + " bytes). Opening viewer…", "info", 30);
  await focusOrOpenViewer();
  // 2. Give the viewer ~400ms to mount and subscribe to storage events.
  await sleep(400);
  await setStatus("Preparing data…", "info", 50);
  await sleep(250);
  await setStatus("Parsing workbook…", "info", 70);
  await sleep(200);
  // 3. Store the bytes — viewer's storage.onChanged listener will render now.
  await chrome.storage.local.set({ lastXlsxB64: b64, lastXlsxSize: size, lastXlsxTs: Date.now() });
  await setStatus("Rendering data…", "info", 90);
  await sleep(150);
  setBadge("OK", "#28a745");
  await setStatus("Done. Data displayed in viewer.", "ok", 100);
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

  // On target page → click Excel (once).
  if (url.startsWith(TARGET_URL)) {
    if (runState.excelClicked) return;
    runState.excelClicked = true;
    runState.stage = "clicking-excel";
    await setStatus("On target page. Waiting for data to load…", "info", 55);
    // small delay to let the page render fully
    setTimeout(async () => {
      try {
        const res = await chrome.tabs.sendMessage(tabId, { type: "CLICK_EXCEL" });
        if (res?.ok) await setStatus("Excel button clicked. Waiting for download…", "info", 75);
        else await setStatus("Excel button NOT found on page.", "error");
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
    chrome.tabs.update(tabId, { url: TARGET_URL });
  }
}

// Capture the xlsx download triggered by the Excel button
chrome.downloads.onCreated.addListener(async (item) => {
  await appendLog("download.onCreated: url=" + item.url + " mime=" + item.mime + " fn=" + item.filename);
  if (!runState) { await appendLog("(no runState — ignoring)", "info"); return; }
  const isXlsx = (item.filename && /\.xlsx?$/i.test(item.filename)) ||
                 (item.mime && /spreadsheet|excel/i.test(item.mime)) ||
                 (item.url && /\.xlsx?(\?|$)/i.test(item.url)) ||
                 (item.url && /excel|export/i.test(item.url));
  if (!isXlsx) { await appendLog("download did not match xlsx heuristics", "info"); return; }

  // If we already have the bytes in storage, this is a duplicate browser download — cancel & erase.
  const { lastXlsxB64 } = await chrome.storage.local.get("lastXlsxB64");
  if (lastXlsxB64 || runState.xlsxDelivered) {
    await appendLog("Duplicate xlsx download (id=" + item.id + ") — cancelling browser save", "info");
    try { await chrome.downloads.cancel(item.id); } catch (_) {}
    try { await chrome.downloads.erase({ id: item.id }); } catch (_) {}
    return;
  }

  runState.downloadId = item.id;
  runState.downloadUrl = item.url;
  await appendLog("xlsx download detected (id=" + item.id + ")");

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
            const v = new Uint8Array(buf, 0, 4);
            if (!(v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04)) {
              console.warn("[sahaj-refetch] not zip");
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
  // Allow this to run even if runState was cleared (so the disk read is a true fallback).
  if (delta.state && delta.state.current === "complete") {
    // If bytes already delivered (via hook/CDP/refetch), skip.
    const cur = await chrome.storage.local.get("lastXlsxB64");
    if (cur.lastXlsxB64) {
      await appendLog("Storage already has xlsx — disk read skipped");
      if (runState) runState = null;
      return;
    }
    try {
      await setStatus("Download complete. Reading file from disk…", "info", 85);
      const buf = await readDownloadedFile(delta.id);
      if (!buf) {
        // Auto-open the extension's details page so the user can flip the toggle.
        const extId = chrome.runtime.id;
        try { await chrome.tabs.create({ url: "chrome://extensions/?id=" + extId, active: true }); } catch (_) {}
        throw new Error(
          "Cannot read downloaded file. The Chrome page just opened — scroll to " +
          "'Allow access to file URLs' and turn it ON, then click the extension icon again."
        );
      }
      const b64 = arrayBufferToBase64(buf);
      await appendLog("Disk read OK — " + buf.byteLength + " bytes");
      await deliverXlsx(b64, buf.byteLength, "disk");
    } catch (e) {
      console.error(e);
      setBadge("ERR", "#dc3545");
      await setStatus(e.message, "error");
    } finally {
      runState = null;
    }
  } else if (delta.error) {
    setBadge("ERR", "#dc3545");
    await setStatus("Download error: " + (delta.error.current || "unknown"), "error");
    runState = null;
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
    const v = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
    if (!(v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04)) {
      console.warn("[sahaj-auto] file did not start with zip magic");
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
