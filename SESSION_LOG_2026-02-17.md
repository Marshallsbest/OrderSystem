# Session Log — February 17, 2026

## Current Version: v0.9.00

## Session Summary
Major infrastructure & admin features session. Added deployment system, header protection, copy management, and UI improvements to the Order System.

---

## Changes Made This Session

### 1. My Info Form — Multi-Field Address (index.html, js.html)
- Replaced the single "Delivery Address" text input in the "Update My Info" modal with four separate fields: **Street**, **City**, **Province**, **Postal Code**
- `showUpdateInfoModal()` parses existing comma-separated addresses into components (handles "ON N0P 1A0" province+postal pattern)
- `submitInfoUpdate()` composes the fields back into a single comma-separated string for storage
- Element IDs: `update-addr-street`, `update-addr-city`, `update-addr-province`, `update-addr-postal`

### 2. Add Customer — Section Checkboxes (admin_form.html, admin_logic.html, Operations.js)
- Added **SECTION ACCESS** checkbox block (4 checkboxes in 2x2 grid) to the Add Customer form
- `getSectionNames()` — reads named ranges `SECTION_A` through `SECTION_D` for display labels (e.g., "Tobacco", "Vape")
- `loadSectionNames()` — populates checkbox labels on form load
- `submitNewClient()` now passes `sections: { A: true/false, ... }` to the backend
- `addNewClient()` writes TRUE/FALSE to section columns and applies `requireCheckbox()` data validation
- Checkboxes default to **checked** and reset on successful submission

### 3. Resilient Column Matching (Operations.js)
- `addNewClient()` now uses **two-tier matching**: exact lowercase first, then `superNormalize()` fallback
- Both `colMap` (normalized) and `colMapRaw` (exact lowercase) are built from headers
- Columns can be renamed, reordered, or reformatted without breaking the system

### 4. Header Protection System (Operations.js, Controller.js)
- **`PROTECTED_SHEETS`** constant — defines which sheets to protect and how many header rows each has
  - CLIENT DATA: 2 rows (row 1 = SECTION_*, row 2 = field headers)
  - PRODUCTS, ORDERS, SETTINGS: 1 row each
- **`backupSheetHeaders()`** — saves golden snapshot to Script Properties as JSON
- **`compareSheetHeaders()`** — generates detailed diff report dialog (🔴 removed, 🟢 added, 🟡 changed, ↔️ moved)
- **`resetSheetHeaders()`** — restores from backup with Yes/No confirmation
- Menu: `Order System → 🛡️ Header Protection`

### 5. Auto-Initialize on First Open (Operations.js, Controller.js)
- `initializeOnFirstOpen_()` runs from `onOpen()` — checks `SYSTEM_INITIALIZED` flag in Script Properties
- On first open: auto-backs up headers, stamps version, records spreadsheet ID and timestamp
- Silent failure — never breaks the menu if something goes wrong

### 6. Deployment & Copy Management System (Operations.js, Controller.js)
- **`registerAsMaster()`** — stamps spreadsheet as master, stores Script ID, initializes copy registry
- **`createCleanCopy()`** — creates sanitized copy for colleagues:
  - Prompts for colleague name
  - Copies spreadsheet via DriveApp
  - Clears: CLIENT DATA rows, ORDERS, DAILY_OPERATIONS, CLIENT_INFO_UPDATES
  - Sanitizes: SETTINGS values containing url/folder/link/drive/export
  - Creates hidden `_SYSTEM_META` sheet with master ID, script ID, version, owner
  - Registers copy in master's `COPY_REGISTRY` (Script Properties)
  - Shows success dialog with clickable link to the new copy
- **`checkForUpdates()`** — context-aware:
  - **Master view**: Deployment Dashboard table showing all copies, owners, versions, ✅/⚠️ status
  - **Copy view**: Shows running version vs. created-from version
- **`showCopyLink()`** — dialog with:
  - Google Sheets `/copy` URL
  - One-click copy-to-clipboard button
  - Auto-generated QR code via Google Charts API
- Menu: `Order System → 📋 Deployment`

### 7. Pull Updates from Master (Operations.js)
- **`pullUpdatesFromMaster()`** — allows copies to self-update without clasp:
  - Reads master Script ID from `_SYSTEM_META` sheet
  - Fetches master's script files via `GET /v1/projects/{id}/content` (Apps Script REST API)
  - Overwrites own files via `PUT /v1/projects/{id}/content`
  - Parses pulled version from Config.js source
  - Updates meta sheet with new version and timestamp
  - Shows success dialog instructing user to reload
  - Handles 403 errors with step-by-step API enablement instructions
- **Prerequisite**: Apps Script API must be enabled at https://script.google.com/home/usersettings
- **OAuth scope added**: `https://www.googleapis.com/auth/script.projects` in appsscript.json

### 8. _README.js — Deploy Guide
- Created `_README.js` — appears first in Apps Script editor (underscore sorts before letters)
- Pure comment block containing:
  - Version number (v0.9.00)
  - Full deployment instructions (step-by-step)
  - Where to save the Web App URL (SETTINGS sheet)
  - How to create QR codes (3 methods: Google Charts API, generator websites, Sheets formula)
  - How to update existing deployments
  - How copies pull updates
  - System files overview

### 9. Git Security Cleanup
- Removed 7 stale `.gs` files from git tracking (`git rm --cached`)
- Expanded `.gitignore`:
  - Added `*.gs` (legacy files)
  - Added `.claspignore`
  - Added `*.secret`
  - Added `_SYSTEM_META/`
  - Added `.idea/`
- Verified: no hardcoded passwords, API keys, or tokens in source code
- `ADMIN_LOGIN` is a runtime named range reference only

---

## Architecture Notes

### Sheet Structure — CLIENT DATA
- **Row 1**: SECTION_A, SECTION_B, SECTION_C, SECTION_D (checkbox columns)
- **Row 2**: CLIENT_ID, Company Name, Type, Phone, Manager, Address, etc.
- **Row 3+**: Data rows

### Key Script Properties (Master)
| Key | Value |
|---|---|
| `IS_MASTER` | `true` |
| `MASTER_ID` | Spreadsheet ID |
| `MASTER_SCRIPT_ID` | Apps Script project ID |
| `MASTER_VERSION` | Current version string |
| `COPY_REGISTRY` | JSON array of copy objects |
| `HEADER_BACKUP` | JSON snapshot of all protected sheet headers |
| `SYSTEM_INITIALIZED` | `true` after first open |

### Key Sheet — _SYSTEM_META (hidden, in copies only)
| Key | Value |
|---|---|
| MASTER_ID | Master spreadsheet ID |
| MASTER_SCRIPT_ID | Master script project ID |
| CREATED_FROM_VERSION | Version at copy time |
| CREATED_AT | ISO timestamp |
| COPY_OWNER | Colleague name |
| MASTER_NAME | Master spreadsheet name |
| LAST_UPDATED | ISO timestamp of last pull |

### Named Ranges Used
- `ADMIN_LOGIN` — admin password (runtime only, not in code)
- `CLIENT_TYPES` — dropdown values for client type
- `SECTION_A` through `SECTION_D` — display names for product sections
- `CFG_SALES_REP` — sales rep name for PDFs

---

## Files Modified This Session
| File | Changes |
|---|---|
| `Config.js` | Version bump to v0.9.00 |
| `Controller.js` | Menu items: Header Protection, Deployment, Copy Link; `initializeOnFirstOpen_()` call; `showCopyLink()` |
| `Operations.js` | `getSectionNames()`, resilient `addNewClient()`, header backup/compare/reset, `initializeOnFirstOpen_()`, `registerAsMaster()`, `createCleanCopy()`, `checkForUpdates()`, `pullUpdatesFromMaster()`, `stampMasterVersion()` |
| `index.html` | Multi-field address in Update My Info modal |
| `js.html` | Address parsing/composing in `showUpdateInfoModal()` and `submitInfoUpdate()` |
| `admin_form.html` | Section checkboxes in Add Customer form |
| `admin_logic.html` | `loadSectionNames()`, section values in `submitNewClient()` |
| `appsscript.json` | Added `script.projects` OAuth scope |
| `_README.js` | NEW — deploy guide, version, QR instructions |
| `.gitignore` | Expanded with *.gs, .claspignore, *.secret, _SYSTEM_META/ |

---

## Pending / Future Considerations
- [ ] **Register as Master** — User still needs to run this from the spreadsheet menu
- [ ] **Test the copy flow end-to-end** — Create a copy, pull updates, verify
- [ ] **Version workflow** — Could create `/update-version-numbers` workflow for consistent bumping
- [ ] **Web App deployment** — After registering as master, user should deploy the web app and save the URL to SETTINGS
- [ ] The Apps Script API enablement is a one-time step for each user who wants to pull updates
