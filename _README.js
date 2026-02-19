/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║                     ORDER SYSTEM - README                       ║
 * ║                                                                  ║
 * ║   Version:  v0.9.16                                              ║
 * ║   Updated:  2026-02-17                                           ║
 * ║                                                                  ║
 * ╚══════════════════════════════════════════════════════════════════╝
 *
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *  HOW TO DEPLOY THE WEB APP
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *
 *  1. In this editor, click  "Deploy"  (top right button)
 *  2. Select  "New deployment"
 *  3. Click the gear icon ⚙️ next to "Select type" → choose  "Web app"
 *  4. Fill in:
 *       • Description:   e.g. "Order System v0.9.16"
 *       • Execute as:    "Me"  (your Google account)
 *       • Who has access: "Anyone"  (so clients can access without login)
 *  5. Click  "Deploy"
 *  6. You will be given a  Web App URL  — COPY THIS URL!
 *
 *
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *  WHERE TO SAVE THE URL
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *
 *  1. Go to the  SETTINGS  sheet in this spreadsheet
 *  2. Find the row with  "WEB_APP_URL"  in column A
 *  3. Paste the deployment URL into  column B
 *
 *  The URL format will look like:
 *    https://script.google.com/macros/s/XXXXXXXXXX/exec
 *
 *  To share a link directly to a specific client, append their ID:
 *    https://script.google.com/macros/s/XXXXXXXXXX/exec?clientId=CLIENT_ID
 *
 *
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *  HOW TO CREATE A QR CODE FOR CLIENTS
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *
 *  Option 1: Google Charts QR API (free, instant)
 *  ─────────────────────────────────────────────────
 *  Paste this into your browser, replacing YOUR_URL with the web app URL:
 *
 *    https://chart.googleapis.com/chart?chs=300x300&cht=qr&chl=YOUR_URL
 *
 *  Example (with a client ID):
 *    https://chart.googleapis.com/chart?chs=300x300&cht=qr&chl=https://script.google.com/macros/s/XXXXX/exec?clientId=ACME
 *
 *  Right-click the QR image → "Save image as" to download it.
 *  Print it on invoices, business cards, or display at the counter.
 *
 *  Option 2: QR Code Generator Websites
 *  ─────────────────────────────────────
 *  • https://www.qr-code-generator.com/
 *  • https://www.qrcode-monkey.com/  (free, customizable with logos)
 *  
 *  Just paste the Web App URL and download the QR code image.
 *
 *  Option 3: Use a Google Sheets formula
 *  ─────────────────────────────────────
 *  In any cell, paste this formula (replace the URL):
 *    =IMAGE("https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl="&ENCODEURL(A1))
 *  Where A1 contains the Web App URL. The QR code will render in the cell.
 *
 *
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *  UPDATING AN EXISTING DEPLOYMENT
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *
 *  When you update the code and want to publish changes:
 *
 *  1. Click  "Deploy"  →  "Manage deployments"
 *  2. Click the  ✏️ pencil icon  next to your active deployment
 *  3. Under "Version", select  "New version"
 *  4. Click  "Deploy"
 *
 *  ⚠️  The URL stays the same — no need to reshare with clients!
 *
 *
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *  PULLING UPDATES (FOR COPIES ONLY)
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *
 *  If this is a COPY of the master spreadsheet:
 *
 *  1. Enable the Apps Script API (one-time):
 *       Go to → https://script.google.com/home/usersettings
 *       Turn ON "Google Apps Script API"
 *
 *  2. In the spreadsheet, go to:
 *       Order System  →  📋 Deployment  →  ⬇️ Pull Updates from Master
 *
 *  3. Confirm the update → code is synced from the master
 *  4. RELOAD the spreadsheet (close & reopen, or Ctrl+Shift+R)
 *
 *  Your data (clients, orders, settings) is NEVER affected by updates.
 *
 *
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *  SYSTEM FILES OVERVIEW
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *
 *  Config.js        → App settings, version number, sheet name constants
 *  Controller.js    → Menu setup, web app entry point (doGet), UI launchers
 *  Operations.js    → Client data, products, header protection, deployment
 *  OrderService.js  → Order processing and submission logic
 *  PDFService.js    → PDF invoice generation and formatting
 *  ProductService.js→ Product catalog management
 *  Setup.js         → Initial template creation helpers
 *  Models.js        → Data models and structures
 *
 *  index.html       → Main web app UI (customer-facing order form)
 *  js.html          → Client-side JavaScript for the order form
 *  css.html         → Stylesheet for the order form
 *
 *  admin_form.html  → Admin dashboard HTML (add products/customers)
 *  admin_logic.html → Admin dashboard JavaScript
 *  admin_css.html   → Admin dashboard styles
 *
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 */
