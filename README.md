# RR Automation Service (Chrome Extension)

Automates: open `https://sahajmobile.com/customers/todayservicing` → log in → click the **Excel** button → capture the downloaded `.xlsx` → display every row as words, serially.

## Files

- `manifest.json` — MV3 manifest
- `background.js` — service worker (orchestration + download capture)
- `content.js` — page automation (login + click Excel)
- `popup.html` / `popup.js` — toolbar UI (Start / View Last)
- `options.html` / `options.js` — credential storage
- `viewer.html` / `viewer.js` — renders downloaded rows as word chips
- `libs/xlsx.full.min.js` — **you must add this** (SheetJS Community Edition)
- `icons/` — optional icons (16/48/128 px)

## One-time setup

1. Add SheetJS. From PowerShell, in the project folder:

   ```powershell
   New-Item -ItemType Directory -Force -Path .\libs | Out-Null
   Invoke-WebRequest -Uri "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js" `
                     -OutFile .\libs\xlsx.full.min.js
   ```

2. (Optional) Add icons at `icons/icon16.png`, `icon48.png`, `icon128.png` — or remove the `icons` block in `manifest.json`.

3. Load the extension:
   - Open `chrome://extensions`
   - Enable **Developer mode**
   - Click **Load unpacked** → select this folder

4. Click the extension icon → **⚙️ Set Credentials** → enter your username & password → Save.

## Run

- Click the extension icon → **Start Automation**.
- A new tab opens to `sahajmobile.com`. The content script fills the login form and submits. After landing on `customers/todayservicing`, it clicks the **Excel** button. The `.xlsx` downloads, and a viewer tab opens showing each row's cells split into word chips, numbered serially.
- Use **View Last Result** to re-open the viewer for the most recent file.

## Security notes

- Credentials live in `chrome.storage.local` (this profile only). Do not use this extension on a shared computer, and do not commit your saved credentials.
- The extension only has access to `https://sahajmobile.com/*`.
- If the login form's field names/IDs differ from the heuristics in `content.js`, update `findInput()` selectors accordingly.

## Troubleshooting

- **Excel button not found** → open the page, right-click the button → Inspect → copy its `id` or a unique class and add it as the first match in `findExcelButton()` in `content.js`.
- **Login form not filled** → inspect the input field names and extend the hint array passed to `findInput()` in `content.js`.
- **Viewer says SheetJS missing** → step 1 above wasn't completed.
