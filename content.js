// content.js — runs on every sahajmobile.com page
(function () {
  console.log("[sahaj-auto] content script loaded:", location.href);
  // Notify background that the page is ready
  try {
    chrome.runtime.sendMessage({ type: "CONTENT_READY", url: location.href });
  } catch (_) {}

  // Listen for xlsx bytes captured by the page-hook (MAIN-world fetch/XHR override).
  window.addEventListener("message", (ev) => {
    const d = ev.data;
    if (!d || !d.__sahaj || d.type !== "XLSX_BYTES") return;
    console.log("[sahaj-auto] received bytes from page-hook:", d.size, "from", d.url);
    chrome.runtime.sendMessage({ type: "XLSX_BYTES", b64: d.b64, size: d.size, url: d.url });
  });

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    (async () => {
      console.log("[sahaj-auto] msg:", msg?.type);
      if (msg.type === "DO_LOGIN") {
        const ok = await tryLogin(msg.creds);
        sendResponse({ ok });
      } else if (msg.type === "CLICK_EXCEL") {
        const ok = await clickExcel(msg.buttonHints, msg.waitForData !== false);
        sendResponse({ ok });
      }
    })();
    return true;
  });

  function $(sel, root = document) { return root.querySelector(sel); }
  function $$(sel, root = document) { return Array.from(root.querySelectorAll(sel)); }

  function findInput(typeOrNames) {
    // Try by name/id/placeholder containing any of the hints
    const hints = typeOrNames.map((s) => s.toLowerCase());
    const inputs = $$("input");
    for (const inp of inputs) {
      const hay = [
        inp.name, inp.id, inp.placeholder, inp.getAttribute("aria-label"), inp.type
      ].filter(Boolean).join(" ").toLowerCase();
      if (hints.some((h) => hay.includes(h))) return inp;
    }
    return null;
  }

  async function tryLogin(creds) {
    // Wait for login fields to appear (form may render late on some routes)
    const passField = await waitFor(() =>
      findInput(["password", "pass"]) || document.querySelector('input[type="password"]'),
      10000
    );
    if (!passField) {
      console.warn("[sahaj-auto] password field not found");
      return false;
    }
    const userField =
      findInput(["username", "user", "email", "login"]) ||
      document.querySelector('input[type="text"]') ||
      document.querySelector('input:not([type="password"]):not([type="hidden"]):not([type="submit"])');
    if (!userField) {
      console.warn("[sahaj-auto] username field not found");
      return false;
    }

    console.log("[sahaj-auto] filling login form");
    setReactValue(userField, creds.username);
    setReactValue(passField, creds.password);

    // Find submit button
    const form = userField.closest("form") || passField.closest("form");
    const submitBtn =
      (form && form.querySelector('button[type="submit"], input[type="submit"]')) ||
      $$('button, input[type="submit"]').find((b) =>
        /sign in|log\s*in|login|submit/i.test((b.value || b.textContent || "").trim())
      );

    if (submitBtn) {
      console.log("[sahaj-auto] clicking submit");
      submitBtn.click();
    } else if (form) {
      console.log("[sahaj-auto] submitting form()");
      form.submit();
    } else {
      console.warn("[sahaj-auto] no submit button or form found");
      return false;
    }
    return true;
  }

  function setReactValue(el, value) {
    const proto = el instanceof HTMLTextAreaElement
      ? window.HTMLTextAreaElement.prototype
      : window.HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, "value").set;
    setter.call(el, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  async function waitForDataLoaded(timeoutMs = 30000) {
    chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Waiting for data table to populate…" });
    const start = Date.now();
    let lastCount = -1;
    let stableSince = 0;
    while (Date.now() - start < timeoutMs) {
      // Count "real" data rows: <tbody><tr>...</tr></tbody> excluding header/loading placeholders.
      const tbodyRows = $$("table tbody tr").filter((tr) => {
        const t = (tr.innerText || "").trim().toLowerCase();
        if (!t) return false;
        if (t.includes("loading") || t.includes("no data") || t.includes("no record")) return false;
        // Need at least one non-empty <td>
        const tds = tr.querySelectorAll("td");
        if (!tds.length) return false;
        return Array.from(tds).some((td) => (td.innerText || "").trim().length > 0);
      });
      const count = tbodyRows.length;
      // Also check for visible loading/spinner overlay
      const spinner = document.querySelector(
        ".loading, .spinner, .loader, .dataTables_processing[style*='block'], [class*='loading']:not([style*='none'])"
      );
      const spinnerVisible = spinner && spinner.offsetParent !== null &&
        getComputedStyle(spinner).display !== "none" &&
        getComputedStyle(spinner).visibility !== "hidden";

      if (count > 0 && !spinnerVisible) {
        if (count === lastCount) {
          if (Date.now() - stableSince >= 1500) {
            chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Data table ready: " + count + " rows", kind: "ok" });
            return true;
          }
        } else {
          lastCount = count;
          stableSince = Date.now();
        }
      }
      await sleep(500);
    }
    chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Timeout waiting for data; last row count=" + lastCount, kind: "error" });
    return lastCount > 0; // proceed anyway if we ever saw rows
  }

  async function clickExcel(buttonHints, waitForData) {
    if (waitForData !== false) {
      await waitForDataLoaded(30000);
    } else {
      chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Skipping data-table wait (waitForData=false)" });
      // Still give the page a moment to render the button itself.
      await sleep(1500);
    }
    const btn = await waitFor(() => findExcelButton(buttonHints), 20000);
    if (!btn) {
      console.warn("[sahaj-auto] Excel button not found");
      chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Excel button not found in DOM after 20s", kind: "error" });
      return false;
    }
    const desc = btn.outerHTML.slice(0, 400);
    chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Excel element: " + desc });

    // If it's an anchor with href, just fetch that URL directly (most reliable).
    const exportUrl = extractExportUrl(btn);
    if (exportUrl) {
      chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Anchor href detected: " + exportUrl + " — fetching directly" });
      try {
        const resp = await fetch(exportUrl, { credentials: "include" });
        if (!resp.ok) {
          chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Direct fetch !ok: " + resp.status, kind: "error" });
        } else {
          const buf = await resp.arrayBuffer();
          if (looksLikeXlsx(buf)) {
            const b64 = arrayBufferToBase64(buf);
            chrome.runtime.sendMessage({ type: "XLSX_BYTES", b64, size: buf.byteLength, url: exportUrl });
            return true;
          } else {
            chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Direct fetch returned non-xlsx (len=" + buf.byteLength + ") — falling back to click", kind: "info" });
          }
        }
      } catch (e) {
        chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Direct fetch threw: " + (e && e.message), kind: "error" });
      }
    }

    // Single click — dispatchEvent OR native .click(), not both (double-fires the download).
    btn.scrollIntoView({ block: "center" });
    try {
      const opts = { bubbles: true, cancelable: true, view: window, button: 0 };
      btn.dispatchEvent(new MouseEvent("mousedown", opts));
      btn.dispatchEvent(new MouseEvent("mouseup", opts));
      btn.dispatchEvent(new MouseEvent("click", opts));
      chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "Click dispatched (mousedown+mouseup+click)" });
    } catch (e) {
      chrome.runtime.sendMessage({ type: "LOG_DEBUG", text: "dispatchEvent failed: " + e.message + " — using native click", kind: "error" });
      try { btn.click(); } catch (_) {}
    }
    return true;
  }

  function extractExportUrl(el) {
    // <a href="...">
    if (el.tagName === "A" && el.href) return el.href;
    // data-url / data-href / formaction
    const da = el.getAttribute("data-url") || el.getAttribute("data-href") ||
               el.getAttribute("formaction") || el.getAttribute("data-export");
    if (da) return new URL(da, location.href).href;
    // onclick handler containing a URL
    const oc = el.getAttribute("onclick") || "";
    const m = oc.match(/['"]([^'"]*?(?:excel|export|xlsx)[^'"]*)['"]/i) ||
              oc.match(/['"](\/[^'"]+)['"]/);
    if (m) return new URL(m[1], location.href).href;
    return null;
  }

  function looksLikeXlsx(buf) {
    // Accept .xlsx (ZIP) OR .csv (printable text). Rejects HTML/CSS/XML/JS.
    if (!buf || buf.byteLength < 2) return false;
    const v = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
    if (v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04) return true;
    const sniff = new Uint8Array(buf, 0, Math.min(1024, buf.byteLength));
    let head = "";
    for (let i = 0; i < Math.min(256, sniff.length); i++) head += String.fromCharCode(sniff[i]);
    const t = head.replace(/^\uFEFF/, "").trimStart().toLowerCase();
    if (t.startsWith("<") || t.startsWith("@font-face") || t.startsWith("@import") ||
        t.startsWith("@charset") || t.startsWith("/*") ||
        /^\s*(function|var|const|let|import|export)\s/.test(t)) return false;
    let printable = 0;
    for (let i = 0; i < sniff.length; i++) {
      const c = sniff[i];
      if (c === 9 || c === 10 || c === 13 || (c >= 32 && c <= 126) || c >= 160) printable++;
    }
    if (printable / sniff.length < 0.92) return false;
    return /[,\n;\t|]/.test(head);
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

  function findExcelButton(hints) {
    const list = (hints && hints.length) ? hints : ["Excel"];
    const candidates = $$("button, a, input[type='button'], input[type='submit'], span[role='button'], div[role='button']");
    const desc = (el) => [
      el.innerText, el.textContent, el.value, el.getAttribute("title"),
      el.getAttribute("aria-label"), el.getAttribute("data-original-title")
    ].filter(Boolean).join(" ").trim();

    const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    for (const hint of list) {
      const exact = new RegExp("^\\s*" + escRe(hint) + "\\s*$", "i");
      const contains = new RegExp(escRe(hint), "i");
      // 1. Exact text + btn class
      let hit = candidates.find((el) => exact.test(desc(el)) && el.classList.contains("btn"));
      if (hit) return hit;
      // 2. Exact text
      hit = candidates.find((el) => exact.test(desc(el)));
      if (hit) return hit;
      // 3. Contains text + btn class
      hit = candidates.find((el) => contains.test(desc(el)) && el.classList.contains("btn"));
      if (hit) return hit;
      // 4. Any element containing text
      hit = candidates.find((el) => contains.test(desc(el)));
      if (hit) return hit;
    }
    return null;
  }

  function text(el) {
    return (el.innerText || el.textContent || el.value || "").trim();
  }

  function waitFor(fn, timeoutMs = 10000, intervalMs = 200) {
    return new Promise((resolve) => {
      const start = Date.now();
      const tick = () => {
        let v = null;
        try { v = fn(); } catch (_) {}
        if (v) return resolve(v);
        if (Date.now() - start > timeoutMs) return resolve(null);
        setTimeout(tick, intervalMs);
      };
      tick();
    });
  }

  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }
})();
