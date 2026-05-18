// page-hook.js — injected into the PAGE's main world so we can override
// window.fetch and XMLHttpRequest. When the page downloads an xlsx, we
// capture the response body and forward it (base64) to the content script
// via window.postMessage.

(function () {
  if (window.__sahajHookInstalled) return;
  window.__sahajHookInstalled = true;
  const TAG = "[sahaj-hook]";

  function isXlsxUrl(u) {
    return /\.xlsx?(\?|#|$)/i.test(u || "") ||
           /\.csv(\?|#|$)/i.test(u || "") ||
           /export|excel|download/i.test(u || "");
  }
  function isXlsxType(ct) {
    return /spreadsheet|excel|officedocument|csv|application\/octet-stream/i.test(ct || "");
  }

  function toB64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(bin);
  }

  function looksLikeXlsx(buf) {
    if (!buf || buf.byteLength < 2) return false;
    const v = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
    if (v[0] === 0x50 && v[1] === 0x4b && v[2] === 0x03 && v[3] === 0x04) return true;
    // CSV: printable text
    const sample = new Uint8Array(buf, 0, Math.min(512, buf.byteLength));
    let printable = 0;
    for (let i = 0; i < sample.length; i++) {
      const c = sample[i];
      if (c === 9 || c === 10 || c === 13 || (c >= 32 && c <= 126) || c >= 160) printable++;
    }
    return printable / sample.length >= 0.92;
  }

  function postBytes(buf, sourceUrl) {
    if (!looksLikeXlsx(buf)) return;
    try {
      const b64 = toB64(buf);
      window.postMessage({
        __sahaj: true,
        type: "XLSX_BYTES",
        b64,
        size: buf.byteLength,
        url: sourceUrl
      }, "*");
      console.log(TAG, "captured xlsx bytes from", sourceUrl, "size:", buf.byteLength);
    } catch (e) { console.warn(TAG, "post failed:", e); }
  }

  // --- fetch override ---
  const origFetch = window.fetch;
  window.fetch = function (...args) {
    const reqUrl = (typeof args[0] === "string") ? args[0] : (args[0] && args[0].url) || "";
    return origFetch.apply(this, args).then((resp) => {
      try {
        const ct = resp.headers.get("content-type") || "";
        if (isXlsxUrl(reqUrl) || isXlsxType(ct)) {
          const clone = resp.clone();
          clone.arrayBuffer().then((buf) => postBytes(buf, reqUrl)).catch(() => {});
        }
      } catch (_) {}
      return resp;
    });
  };

  // --- XHR override ---
  const XHR = window.XMLHttpRequest;
  const origOpen = XHR.prototype.open;
  const origSend = XHR.prototype.send;
  XHR.prototype.open = function (method, url) {
    this.__sahajUrl = url;
    return origOpen.apply(this, arguments);
  };
  XHR.prototype.send = function () {
    this.addEventListener("loadend", () => {
      try {
        const url = this.__sahajUrl || "";
        const ct = this.getResponseHeader && this.getResponseHeader("content-type");
        if (isXlsxUrl(url) || isXlsxType(ct)) {
          let buf = null;
          if (this.response instanceof ArrayBuffer) buf = this.response;
          else if (this.response instanceof Blob) {
            this.response.arrayBuffer().then((b) => postBytes(b, url)).catch(() => {});
            return;
          } else if (typeof this.responseText === "string" && this.responseText.length) {
            // CSV-style text response — convert to bytes.
            buf = new TextEncoder().encode(this.responseText).buffer;
          }
          if (buf) postBytes(buf, url);
        }
      } catch (_) {}
    });
    return origSend.apply(this, arguments);
  };

  console.log(TAG, "installed on", location.href);
})();
