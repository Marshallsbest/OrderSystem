# Order System — Google Apps Script

> **⚠️ Important:** This is **functional source code**, not an installable plugin or add-on.
> It runs entirely inside a Google Apps Script project bound to a Google Spreadsheet.
> You must copy the files into a Google Apps Script editor and bind them to your own
> spreadsheet. There is no `.exe`, no npm package, and no marketplace listing.

---

## What This Is

The Order System is a **client-facing B2B order management platform** built on Google Apps Script and Google Sheets. It allows sales representatives to share a web link (or QR code) with a retail client, who can then log in, browse a product catalog, and submit an order — all from their phone or computer, with no app download required.

Once an order is submitted, it is saved to the Google Spreadsheet, a PDF order form is automatically generated and saved to Google Drive, and the admin can review, adjust, and re-export the order at any time from within the spreadsheet.

### Core Capabilities

- **Mobile-first order form web app** — clients browse products by category, enter quantities, and submit with one tap
- **Client-gated access** — each client logs in with their unique Client ID; they only see the products and sections they are permitted to view
- **Automatic PDF generation** — each submitted order produces a formatted PDF invoice and/or a completed order form sheet
- **Admin dashboard** (inside the spreadsheet) — add products, add customers, manage categories, generate PDFs for any selected order
- **Order revision tracking** — re-submitted orders are stored as `Rev:N` revisions of the original invoice
- **Category & section management** — products are grouped into colour-coded categories; clients can be restricted to specific sections (A, B, C, D)
- **Fully configurable** — all labels, colours, column mappings, URLs, and section names are controlled via the `SETTINGS` sheet — no code changes required for routine admin tasks

---

## Architecture Overview

```
Google Spreadsheet (data store)
    ├── SETTINGS         — all configuration, colours, column mappings
    ├── CLIENT DATA      — one row per client, with section access flags
    ├── PRODUCTS         — product catalog with variations, pricing, commission
    ├── ORDERS           — one row per submitted order (compact encoded format)
    ├── ORDER_FORM_1/2   — printable order form templates (populated by PDF engine)
    ├── ORDERS_EXPORT    — staging area for export summaries
    ├── DAILY_OPERATIONS — auto-refreshed sales summary dashboard
    └── DASHBOARD        — clickable admin action panel

Google Apps Script (this repository)
    └── Bound to the spreadsheet above

Web App (served by Apps Script)
    └── Accessible via a deployed URL shared with clients
```

---

## File Reference

### Backend — Google Apps Script (`.js` files)

#### `Config.js`
Global constants and shared utility functions used across all other files.
- `SHEET_NAMES` — object mapping logical names to spreadsheet tab names
- `ORDER_FORM_COLORS` — brand colour constants for PDF output
- `superNormalize()` — canonical string normalizer used for fuzzy key matching
- `columnToLetter()` — converts column index to A1-notation letter
- `getSheet()` — safe sheet accessor with case-insensitive fallback
- `include()` — server-side HTML template include helper
- `getOrderFormSheetName()` — reads `FORM_N_SHEET` key from SETTINGS to find the correct order form tab
- `getOrderFormTemplates()` — returns all configured form template mappings for the admin UI

---

#### `Controller.js`
Entry point and UI wiring. This is the file Apps Script calls directly.
- `doGet(e)` — web app entry point; reads URL params (`clientId`, `orderId`), builds the HTML template, and serves the order form page
- `onOpen()` — installs the **Order System** spreadsheet menu with all admin actions
- `showAddProductSidebar()` — opens the product/customer management sidebar
- `showOrderFormDialog()` — detects the active order (if any) and launches the web app in a popup window
- `installDashboardTrigger()` — installs the `onSelectionChange` trigger for the DASHBOARD clickable buttons
- `updateConfigSetting()` — wrapper called from admin UI to update a SETTINGS row
- `showCopyLink()` — displays a shareable copy-link dialog with QR code for distributing the spreadsheet to colleagues

---

#### `Operations.js`
Core data access and administrative operations. The largest file.
- `getAppConfig()` — reads all key/value pairs from SETTINGS; returns a config object including the `ADMIN_KEY` (read from the `ADMIN_LOGIN` named range)
- `updateConfigSetting()` — upserts a key/value row in SETTINGS
- `onSelectionChange(e)` — handles clickable DASHBOARD button actions via selection trigger
- `saveClientInfoUpdate()` — writes client update requests to the `CLIENT_INFO_UPDATES` sheet for admin review
- `getClientById()` — fetches a single client record and calculates section permissions
- `getClientData()` — reads the full CLIENT DATA sheet with dual-header support (section headers in row 1, data headers in row 2)
- `getClientTypes()` — reads the `CLIENT_TYPES` named range for dropdown population
- `getSectionNames()` — reads `SECTION_A–D` named ranges for dynamic section labels
- `addNewClient()` — appends a new client row with proper column mapping and checkbox validation
- `getAppStyles()` — reads `PRIMARY_COLOUR`, `SECONDARY_COLOUR` etc. from named ranges; supports named colours (e.g. "Orange") via `COLOUR_FORMAT_DEFINITIONS`
- `getCategorySettings()` — reads the category table from SETTINGS; returns colour, display order, section gating, and sale status per category
- `getVariationDefaults()` / `getVariationGroups()` — reads variation group definitions from the `VARIATION_GROUPS_AND_VALUES` named range
- `addValueToVariationGroup()` / `createNewVariationGroup()` — mutate the variation groups table

---

#### `OrderService.js`
Order submission and retrieval.
- `processOrder(orderData)` — validates and writes an order row to the ORDERS sheet; handles revision numbering; triggers PDF generation; uses `LockService` to prevent concurrent write conflicts
- `getOrderById(orderId)` — fetches and parses a single order row by invoice number, returning a structured object with items decoded from the compact `[qty|@sku|$price|flag]` encoding
- `getOrdersByClient(clientName)` — returns all orders for a given client name, sorted newest-first

---

#### `Models.js`
Data model factories — pure functions, no I/O.
- `createProductModel(rawData, parentModel, rowIndex)` — builds a fully normalized product object from a raw sheet row; handles parent→child inheritance for name, category, pricing, commission rates, colour, section, and order form assignment; computes `hasCase`, `isAvailable`, and variation headers
- `createOrderModel(rawData)` — builds a WooCommerce-style order object from a raw payload; normalizes line items and billing/shipping structures

---

#### `ProductService.js`
Product catalog management — reading and writing to the PRODUCTS sheet.
- `getProductCatalog()` — reads all products with dynamic header detection; supports variable column order; filters out rows missing SKU or name; reads all variation, pricing, commission, description, image, and colour fields
- `getExistingSkus()` — returns all current SKUs for duplicate checking
- `addProductBatch(newItems)` — appends one or more new product rows; maps fields by column name; auto-assigns REF character sequences and SKUs based on the product name prefix

---

#### `OrderFormPDFService.js`
Generates formatted PDF order forms by writing quantities into the `ORDER_FORM_N` sheet and exporting.
- `generateOrderFormHtmlPdf(params)` — the main PDF entry point; matches ordered items to rows in the order form template by REF code or SKU; writes quantities into the appropriate `Qty` cells; exports the populated sheet to PDF; returns the Drive file URL
- `generateSelectedOrderFormPdf()` — spreadsheet menu trigger; reads the currently selected ORDERS row and calls the PDF generator
- `addProductToOrderFormSheet(sheet, product)` — inserts a product row into the order form template above the "Shipping" sentinel row

---

#### `PDFService.js`
Generates PDF invoices (line-item summary style, distinct from the order form template).
- `createOrderPdf()` / `generateSelectedOrderPdf()` — builds a formatted invoice PDF from order data using Google Docs or Sheets export

---

#### `Setup.js`
Sheet and template maintenance helpers called after initial installation.
- `setupSettingsSheet()` — ensures SETTINGS exists with correct headers and default keys
- `setupOrderDataSheet()` — creates the `ORDER_DATA` staging sheet used for bulk PDF population
- `styleProductHeaders()` / `cleanupProductSheet()` — apply colour formatting to PRODUCTS sheet based on category colours; remove blank/duplicate rows
- `refreshDailyOperationsDashboard()` — recalculates the DAILY_OPERATIONS summary from ORDERS data
- `backupSheetHeaders()` / `compareSheetHeaders()` / `resetSheetHeaders()` — header protection utilities that save and validate column headers to prevent accidental restructuring

---

#### `Installer.js`
Full spreadsheet bootstrapper — run once on a blank spreadsheet.
- `runInstaller()` — orchestrates the creation of all required sheets, named ranges, and default values
- `_createSheet_*()` functions — create each individual sheet with correct headers, sample rows, and column widths
- `_setupNamedRanges()` — maps all `getRangeByName()` calls used throughout the codebase to their corresponding SETTINGS rows
- `_getDefaultSettings()` — returns the complete list of default SETTINGS key/value pairs
- `_protectSystemSheets()` — applies warning-only protection to ORDERS and system sheets
- `_showSetupWizard()` / `_runInstallerFromSidebar()` — sidebar integration for the step-by-step setup UI

---

#### `AnalyticsService.js`
Tracks usage metrics within the spreadsheet for internal reporting.

#### `CheckHeaders.js` / `HEADER_EXTRACTOR.js`
Utilities for validating that sheet column headers match expected structure; used by the header backup/restore system.

#### `Debug.js` / `DebugSheetStats.js` / `TestRunner.js`
Development utilities — log sheet stats, run internal test cases, and debug data mapping issues. Not used in production.

---

### Frontend — HTML Templates

#### `index.html`
The main customer-facing web app shell. Sets up the page structure, injects server-side template variables (`categorySettings`, `appStyles`, `clientId`, `version`), applies the primary theme colour from settings, and includes the CSS and JS partials.

#### `js.html`
All client-side JavaScript for the order form (~1800 lines).
- App initialization, client data fetch, and product rendering
- Product grouping by category with colour-coded headers and sort-order support
- Matrix and compact table rendering modes for different product types
- Quantity input handling, total calculation, and form submission via `google.script.run`
- Admin mode (unlocked with the admin key) enabling order editing, client switching, and summary PDF option
- Order history pre-fill when an `orderId` URL parameter is present

#### `css.html`
Material Design 3 stylesheet for the order form — dark/light surfaces, text fields, buttons, category headers, product tables, sticky headers, and snackbar notifications.

#### `admin_form.html`
HTML markup for the Add Product / Add Customer admin sidebar panel. Supports four modes: Single Entry, Bulk Generator, Add Variation, Edit Product, and Archive.

#### `admin_logic.html`
JavaScript for the admin sidebar — form validation, bulk variation generation, product submission, customer creation, and communication with the server via `google.script.run`.

#### `admin_css.html`
Stylesheet for the admin sidebar panel.

#### `addProductSidebar.html`
Legacy/standalone sidebar for adding products (pre-dates the modular admin panel).

#### `installer_sidebar.html`
Step-by-step setup wizard UI that calls `runInstaller()` from a guided sidebar interface.

---

## How to Deploy

> These are manual steps — there is no automated installer script for GitHub.

1. **Create a new Google Spreadsheet**
2. Open **Extensions → Apps Script**
3. Copy all `.js` files from this repository into the Apps Script editor (one file per `.gs` script file)
4. Copy all `.html` files into the Apps Script editor as HTML files
5. Run **`runInstaller()`** from the editor once to create all required sheets
6. In the Apps Script editor, go to **Deploy → New Deployment → Web App**
   - Execute as: **Me**
   - Who has access: **Anyone** (for client access without login)
7. Copy the deployment URL and paste it into `SETTINGS → WEB_APP_URL`
8. Add your clients to `CLIENT DATA` (row 3 onward)
9. Add your products to the `PRODUCTS` sheet
10. Share the Web App URL (or a QR code) with your clients

---

## Security Notes

- The `ADMIN_LOGIN` named range in SETTINGS controls the admin password. Set this before sharing the spreadsheet.
- Client IDs act as access tokens — keep them unique and non-guessable.
- All data reads/writes are server-side; clients have no direct sheet access.
- Sheet protection (warning-only) is applied to ORDERS and system sheets by the installer.

---

## Requirements

- A Google Account
- Google Sheets + Google Apps Script (both free)
- No external dependencies, npm packages, or third-party services
