/**
 * Installer.js
 * ============================================================
 * Full spreadsheet bootstrapper for the Order System add-on.
 *
 * Call  runInstaller()  from the Apps Script editor (or trigger it
 * from the add-on's onInstall / onOpen handler) to build a brand-new
 * spreadsheet with every required sheet, table, named range, and
 * default configuration value.
 *
 * Safe to re-run: existing sheets are NOT deleted — only missing
 * pieces are created / filled in.
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
//  ENTRY POINT
// ─────────────────────────────────────────────────────────────

/**
 * Main installer — call this once after binding the script to a new
 * spreadsheet.  Wires up every sheet, named range, and default value
 * the Order System needs to operate.
 */
function runInstaller() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    try {
        Logger.log('=== ORDER SYSTEM INSTALLER STARTED ===');

        _createSheet_Dashboard(ss);
        _createSheet_Settings(ss);
        _createSheet_ClientData(ss);
        _createSheet_Products(ss);
        _createSheet_Orders(ss);
        _createSheet_OrdersExport(ss);
        _createSheet_OrderForm(ss, 1);   // ORDER_FORM_1
        _createSheet_OrderForm(ss, 2);   // ORDER_FORM_2 (spare)
        _createSheet_DailyOperations(ss);
        _createSheet_Welcome(ss);

        _setupNamedRanges(ss);
        _applyDefaultSettings(ss);
        _protectSystemSheets(ss);
        _setSheetOrder(ss);

        Logger.log('=== ORDER SYSTEM INSTALLER COMPLETE ===');

        ui.alert(
            '✅ Order System Installed',
            'All sheets, named ranges, and default settings have been created.\n\n' +
            'Next steps:\n' +
            '1. Set your ADMIN_LOGIN password in the SETTINGS sheet (or name the cell ADMIN_LOGIN).\n' +
            '2. Add your products to the PRODUCTS sheet.\n' +
            '3. Add your clients to the CLIENT DATA sheet.\n' +
            '4. Deploy as a Web App from Extensions → Apps Script → Deploy.',
            ui.ButtonSet.OK
        );

    } catch (e) {
        Logger.log('INSTALLER ERROR: ' + e.message + '\n' + e.stack);
        ui.alert('❌ Installer Error', e.message, ui.ButtonSet.OK);
    }
}


// ─────────────────────────────────────────────────────────────
//  SHEET BUILDERS
// ─────────────────────────────────────────────────────────────

function _createSheet_Dashboard(ss) {
    let sheet = ss.getSheetByName('DASHBOARD');
    if (!sheet) sheet = ss.insertSheet('DASHBOARD');
    sheet.clear();

    const title = [['ORDER SYSTEM — DASHBOARD']];
    sheet.getRange('A1').setValue('ORDER SYSTEM — DASHBOARD')
        .setFontSize(16).setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#1a73e8');
    sheet.getRange('A1:F1').mergeAcross().setHorizontalAlignment('center');

    const actions = [
        ['', 'ACTION', 'DESCRIPTION', '', '', ''],
        ['', '▶ Launch Order Form', 'Open the web app order form for a client', '', '', ''],
        ['', '', '', '', '', ''],
        ['', '➕ Add Product', 'Open the product editor sidebar', '', '', ''],
        ['', '📄 Generate PDF', 'Export a selected order as PDF', '', '', ''],
        ['', '🧹 Cleanup Products', 'Remove duplicate / blank product rows', '', '', ''],
        ['', '🎨 Style Headers', 'Reapply category colour formatting', '', '', ''],
        ['', '📊 Refresh Dashboard', 'Recalculate daily operations summary', '', '', ''],
    ];

    sheet.getRange(2, 1, actions.length, 6).setValues(actions);
    sheet.getRange('B2:C2').setFontWeight('bold').setBackground('#f8f9fa');
    sheet.setColumnWidth(1, 40);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 380);
    sheet.setFrozenRows(2);

    Logger.log('[Installer] DASHBOARD sheet created.');
}


function _createSheet_Settings(ss) {
    let sheet = ss.getSheetByName('SETTINGS');
    if (!sheet) sheet = ss.insertSheet('SETTINGS');

    // Only write if the sheet is essentially empty
    if (sheet.getLastRow() < 2) {
        sheet.clear();
        sheet.getRange('A1:B1').setValues([['SETTING KEY', 'VALUE']])
            .setFontWeight('bold').setBackground('#e8f0fe');

        const defaults = _getDefaultSettings();
        sheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
        sheet.setColumnWidth(1, 240);
        sheet.setColumnWidth(2, 280);
        sheet.setFrozenRows(1);

        Logger.log('[Installer] SETTINGS sheet created with ' + defaults.length + ' defaults.');
    } else {
        // Sheet exists — inject any missing keys
        const existing = sheet.getDataRange().getValues();
        const existingKeys = existing.map(r => String(r[0]).trim().toUpperCase());
        const defaults = _getDefaultSettings();
        const missing = defaults.filter(d => !existingKeys.includes(String(d[0]).toUpperCase()));
        if (missing.length > 0) {
            const nextRow = sheet.getLastRow() + 1;
            sheet.getRange(nextRow, 1, missing.length, 2).setValues(missing);
            Logger.log('[Installer] SETTINGS: injected ' + missing.length + ' missing keys.');
        }
    }
}


function _createSheet_ClientData(ss) {
    let sheet = ss.getSheetByName('CLIENT DATA');
    if (!sheet) sheet = ss.insertSheet('CLIENT DATA');
    if (sheet.getLastRow() >= 2) {
        Logger.log('[Installer] CLIENT DATA exists — skipping overwrite.');
        return;
    }

    sheet.clear();

    // Row 1: section permission headers (merged label + individual section columns)
    const row1 = ['CLIENT DATA', '', '', '', '', '', '', '', 'SECTION ACCESS', '', '', ''];
    sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
    sheet.getRange('A1:H1').merge();
    sheet.getRange('I1').setValue('SECTION ACCESS');
    sheet.getRange('I1:L1').merge().setHorizontalAlignment('center')
        .setBackground('#d9ead3').setFontWeight('bold');

    // Row 2: data column headers
    const headers = [
        'CLIENT_ID', 'Company Name', 'Address', 'Phone', 'Contact Name',
        'Sales Rep', 'Type', 'Min Order',
        'SECTION_A', 'SECTION_B', 'SECTION_C', 'SECTION_D'
    ];
    sheet.getRange(2, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#f3f3f3');

    // Row 3: sample/placeholder row
    const sample = ['CLIENT001', 'Sample Company', '123 Main St', '555-1234',
        'John Doe', 'Rep Name', 'Retail', '100', true, true, false, false];
    sheet.getRange(3, 1, 1, sample.length).setValues([sample]);

    // Checkboxes for section columns (D3:G3 and beyond)
    const checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange('I3:L3').setDataValidation(checkRule);

    // Column widths
    [200, 200, 200, 120, 160, 120, 100, 100, 90, 90, 90, 90].forEach((w, i) => {
        sheet.setColumnWidth(i + 1, w);
    });
    sheet.setFrozenRows(2);

    Logger.log('[Installer] CLIENT DATA sheet created.');
}


function _createSheet_Products(ss) {
    let sheet = ss.getSheetByName('PRODUCTS');
    if (!sheet) sheet = ss.insertSheet('PRODUCTS');
    if (sheet.getLastRow() >= 1) {
        Logger.log('[Installer] PRODUCTS exists — skipping overwrite.');
        return;
    }

    const headers = [
        'Node', 'SKU', 'Brand', 'Product Name', 'Parent Name', 'Category',
        'Variation 1', 'Variation 2', 'Variation 3', 'Variation 4',
        'Price', 'Sale Price', 'MC Price', 'Units per Case', 'On Sale',
        'Description', 'Image URL', 'Color', 'REF', 'Status',
        'Order Form #', 'Section'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#ffe0b2').setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 120);
    sheet.setFrozenRows(1);

    Logger.log('[Installer] PRODUCTS sheet created.');
}


function _createSheet_Orders(ss) {
    let sheet = ss.getSheetByName('ORDERS');
    if (!sheet) sheet = ss.insertSheet('ORDERS');
    if (sheet.getLastRow() >= 1) {
        Logger.log('[Installer] ORDERS exists — skipping overwrite.');
        return;
    }

    const headers = [
        'Timestamp', 'Order ID', 'Client ID', 'Client Name', 'Address',
        'Sales Rep', 'Product REF', 'Product Name', 'Variation', 'Qty Singles',
        'Qty Cases', 'Unit Price', 'Sale Price', 'Total Units', 'Line Total',
        'Commission Rate', 'Commission $', 'Order Total', 'Notes'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#e8f0fe').setFrozenRows(1);
    sheet.setFrozenRows(1);

    Logger.log('[Installer] ORDERS sheet created.');
}


function _createSheet_OrdersExport(ss) {
    let sheet = ss.getSheetByName('ORDERS_EXPORT');
    if (!sheet) sheet = ss.insertSheet('ORDERS_EXPORT');
    if (sheet.getLastRow() >= 1) {
        Logger.log('[Installer] ORDERS_EXPORT exists — skipping overwrite.');
        return;
    }
    sheet.getRange('A1').setValue('ORDER EXPORT — populated automatically when orders are submitted.')
        .setFontStyle('italic').setFontColor('#888888');
    Logger.log('[Installer] ORDERS_EXPORT sheet created.');
}


function _createSheet_OrderForm(ss, num) {
    const name = 'ORDER_FORM_' + num;
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() >= 3) {
        Logger.log('[Installer] ' + name + ' exists — skipping overwrite.');
        return;
    }

    sheet.clear();

    // Title row
    sheet.getRange('A1:F1').merge()
        .setValue('ORDER FORM ' + num + ' — Add products below. Add a row labelled "Shipping" before the last row.')
        .setFontWeight('bold').setBackground('#cc66cc').setFontColor('#ffffff')
        .setHorizontalAlignment('center');

    // Column headers
    const headers = ['REF', 'Product Name', 'Packaging / Description', 'Price', 'MC Price', 'Qty'];
    sheet.getRange(2, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#f3e5f5');

    // Shipping sentinel row (used by addProductToOrderFormSheet to find insert point)
    sheet.getRange(3, 1, 1, headers.length).setValues([['', 'Shipping', '', '', '', '']]);

    [60, 240, 200, 80, 80, 60].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(2);

    Logger.log('[Installer] ' + name + ' sheet created.');
}


function _createSheet_DailyOperations(ss) {
    let sheet = ss.getSheetByName('DAILY_OPERATIONS');
    if (!sheet) sheet = ss.insertSheet('DAILY_OPERATIONS');
    if (sheet.getLastRow() >= 1) {
        Logger.log('[Installer] DAILY_OPERATIONS exists — skipping overwrite.');
        return;
    }
    sheet.getRange('A1').setValue('DAILY OPERATIONS — refreshed automatically.')
        .setFontStyle('italic').setFontColor('#888888');
    Logger.log('[Installer] DAILY_OPERATIONS sheet created.');
}


function _createSheet_Welcome(ss) {
    let sheet = ss.getSheetByName('Welcome');
    if (!sheet) sheet = ss.insertSheet('Welcome');
    if (sheet.getLastRow() >= 2) {
        Logger.log('[Installer] Welcome exists — skipping overwrite.');
        return;
    }

    sheet.clear();
    sheet.getRange('A1').setValue('Welcome to the Order System')
        .setFontSize(20).setFontWeight('bold').setFontColor('#1a73e8');
    sheet.getRange('A2').setValue('Version: ' + (typeof CURRENT_VERSION !== 'undefined' ? CURRENT_VERSION : 'v0.9.21'));
    sheet.getRange('A4').setValue('Quick Start:').setFontWeight('bold');
    const steps = [
        ['1. Fill in your SETTINGS sheet with your business details.'],
        ['2. Add your clients to CLIENT DATA (row 3 onward).'],
        ['3. Add your products to the PRODUCTS sheet.'],
        ['4. Set up your ORDER_FORM_1 sheet with the products you want on the form.'],
        ['5. Deploy the web app from Extensions → Apps Script → Deploy → New Deployment.'],
        ['6. Share the Web App URL with your sales reps.'],
    ];
    sheet.getRange(5, 1, steps.length, 1).setValues(steps);
    sheet.setColumnWidth(1, 600);

    Logger.log('[Installer] Welcome sheet created.');
}


// ─────────────────────────────────────────────────────────────
//  NAMED RANGES
// ─────────────────────────────────────────────────────────────

/**
 * Creates every named range the Order System reads via getRangeByName().
 * Named ranges are defined in SETTINGS so admins can edit values in-cell.
 */
function _setupNamedRanges(ss) {
    const settingsSheet = ss.getSheetByName('SETTINGS');
    if (!settingsSheet) throw new Error('SETTINGS sheet must exist before setting up named ranges.');

    const data = settingsSheet.getDataRange().getValues();

    // Helper: find the row number of a key in SETTINGS (1-indexed)
    const findRow = (key) => {
        for (let i = 0; i < data.length; i++) {
            if (String(data[i][0]).trim().toUpperCase() === key.toUpperCase()) return i + 1;
        }
        return -1;
    };

    // Map: named range name → SETTINGS key whose VALUE cell (col B) it should point to
    const rangeMap = {
        'ADMIN_LOGIN': 'ADMIN_LOGIN',
        'CFG_SALES_REP': 'CFG_SALES_REP',
        'PRIMARY_COLOUR': 'PRIMARY_COLOUR',
        'PRIMARY_COLOUR_TEXT': 'PRIMARY_COLOUR_TEXT',
        'SECONDARY_COLOUR': 'SECONDARY_COLOUR',
        'SECONDARY_COLOUR_TEXT': 'SECONDARY_COLOUR_TEXT',
        'SECTION_A': 'SECTION_A_NAME',
        'SECTION_B': 'SECTION_B_NAME',
        'SECTION_C': 'SECTION_C_NAME',
        'SECTION_D': 'SECTION_D_NAME',
    };

    Object.entries(rangeMap).forEach(([rangeName, settingsKey]) => {
        const row = findRow(settingsKey);
        if (row === -1) {
            Logger.log('[Installer] WARNING: SETTINGS key "' + settingsKey + '" not found — skipping named range "' + rangeName + '".');
            return;
        }
        // Point to column B (value cell) of that row
        const cell = settingsSheet.getRange(row, 2);
        _setNamedRange(ss, rangeName, cell);
    });

    // CLIENT_TYPES — points to a dedicated block of cells in SETTINGS
    _setupClientTypesRange(ss, settingsSheet, data);

    // COLOUR_FORMAT_DEFINITIONS — points to a dedicated table in SETTINGS
    _setupColourDefinitionsRange(ss, settingsSheet, data);

    // VARIATION_GROUPS_AND_VALUES — points to a dedicated table in SETTINGS
    _setupVariationGroupsRange(ss, settingsSheet, data);

    Logger.log('[Installer] Named ranges configured.');
}


/** Create or update a named range safely. */
function _setNamedRange(ss, name, range) {
    try {
        const existing = ss.getRangeByName(name);
        if (existing) {
            // Update to new range
            ss.getNamedRanges().find(nr => nr.getName() === name)?.setRange(range);
        } else {
            ss.setNamedRange(name, range);
        }
        Logger.log('[Installer] Named range set: ' + name + ' → ' + range.getA1Notation());
    } catch (e) {
        Logger.log('[Installer] WARNING: Could not set named range "' + name + '": ' + e.message);
    }
}


/** CLIENT_TYPES — a vertical list of client type strings in SETTINGS. */
function _setupClientTypesRange(ss, settingsSheet, data) {
    // Find the section header
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === 'CLIENT_TYPES') { headerRow = i + 1; break; }
    }
    if (headerRow === -1) {
        Logger.log('[Installer] CLIENT_TYPES header not found in SETTINGS — skipping named range.');
        return;
    }
    // The values start one row below the header and span up to 20 rows
    const range = settingsSheet.getRange(headerRow + 1, 2, 20, 1);
    _setNamedRange(ss, 'CLIENT_TYPES', range);
}


/** COLOUR_FORMAT_DEFINITIONS — a 3-column table (Name | Hex | Text Color) in SETTINGS. */
function _setupColourDefinitionsRange(ss, settingsSheet, data) {
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === 'COLOUR_FORMAT_DEFINITIONS') { headerRow = i + 1; break; }
    }
    if (headerRow === -1) {
        Logger.log('[Installer] COLOUR_FORMAT_DEFINITIONS header not found in SETTINGS — skipping.');
        return;
    }
    const range = settingsSheet.getRange(headerRow + 1, 1, 30, 3);
    _setNamedRange(ss, 'COLOUR_FORMAT_DEFINITIONS', range);
}


/** VARIATION_GROUPS_AND_VALUES — a 3-column table (Group | Var# | Values CSV) in SETTINGS. */
function _setupVariationGroupsRange(ss, settingsSheet, data) {
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === 'VARIATION_GROUPS_AND_VALUES') { headerRow = i + 1; break; }
    }
    if (headerRow === -1) {
        Logger.log('[Installer] VARIATION_GROUPS_AND_VALUES header not found in SETTINGS — skipping.');
        return;
    }
    const range = settingsSheet.getRange(headerRow + 1, 1, 50, 3);
    _setNamedRange(ss, 'VARIATION_GROUPS_AND_VALUES', range);
}


// ─────────────────────────────────────────────────────────────
//  DEFAULT SETTINGS VALUES
// ─────────────────────────────────────────────────────────────

/**
 * Returns the full list of SETTINGS rows as [key, value] pairs.
 * These become the default values written to the SETTINGS sheet.
 * Admins can change values in column B directly.
 */
function _getDefaultSettings() {
    return [
        // ── App Identity ──────────────────────────────────────────
        ['APP_TITLE', 'My Order System'],
        ['WEB_APP_URL', ''],   // Paste your deployed Web App URL here

        // ── Admin Access ──────────────────────────────────────────
        ['ADMIN_LOGIN', 'CHANGE_ME'],  // ⚠️ Change before sharing!

        // ── Branding ──────────────────────────────────────────────
        ['PRIMARY_COLOUR', '#32CD32'],
        ['PRIMARY_COLOUR_TEXT', '#000000'],
        ['SECONDARY_COLOUR', '#625b71'],
        ['SECONDARY_COLOUR_TEXT', '#ffffff'],

        // ── Sales Rep / Commission ────────────────────────────────
        ['CFG_SALES_REP', 'Sales Rep Name'],
        ['CFG_COL_REP', 'Sales Rep'],
        ['CFG_COL_CONTACT', 'Contact Name'],

        // ── Section Names (shown in order form tabs) ───────────────
        ['SECTION_A_NAME', 'Section A'],
        ['SECTION_B_NAME', 'Section B'],
        ['SECTION_C_NAME', 'Section C'],
        ['SECTION_D_NAME', 'Section D'],

        // ── Order Form Template Sheet Mapping ─────────────────────
        ['FORM_1_SHEET', 'ORDER_FORM_1'],
        ['FORM_2_SHEET', 'ORDER_FORM_2'],

        // ── Client Types (used for dropdown in admin) ─────────────
        ['CLIENT_TYPES', '— values listed below —'],
        ['', 'Retail'],
        ['', 'Wholesale'],
        ['', 'Distributor'],
        ['', 'Online'],
        ['', 'Other'],

        // ── Colour Format Definitions ─────────────────────────────
        // Named colors that can be used instead of hex codes in category settings
        // Format: Color Name | Hex Value | Text Color
        ['COLOUR_FORMAT_DEFINITIONS', '— color name, hex, text color —'],
        ['Orange', '#FFA500', '#000000'],
        ['Teal', '#008080', '#ffffff'],
        ['Navy', '#003366', '#ffffff'],
        ['Purple', '#6A0DAD', '#ffffff'],
        ['Green', '#32CD32', '#000000'],
        ['Red', '#cc0000', '#ffffff'],
        ['Blue', '#1a73e8', '#ffffff'],
        ['Grey', '#888888', '#ffffff'],
        ['Gold', '#FFD700', '#000000'],

        // ── Variation Groups & Values ─────────────────────────────
        // Format: Group Name | Variation Number | Values (comma separated)
        ['VARIATION_GROUPS_AND_VALUES', '— group name, var#, values —'],
        ['Strengths', '1', '10mg, 20mg, 50mg'],
        ['Sizes', '2', 'Small, Medium, Large'],
        ['Flavours', '3', 'Original, Mint, Berry'],

        // ── Category Settings table ───────────────────────────────
        // (Managed in the SETTINGS sheet directly — rows below CATEGORIES header)
        ['CATEGORIES', '— name, colour, text colour, order, sale active, section —'],
    ];
}


// ─────────────────────────────────────────────────────────────
//  POST-INSTALL HELPERS
// ─────────────────────────────────────────────────────────────

/**
 * Injects default values into SETTINGS for keys that are already missing
 * (safe to call multiple times — won't overwrite existing values).
 */
function _applyDefaultSettings(ss) {
    const sheet = ss.getSheetByName('SETTINGS');
    if (!sheet) return;
    // Already handled in _createSheet_Settings — nothing extra needed.
    Logger.log('[Installer] Default settings applied.');
}


/**
 * Lightly protects system sheets so accidental edits trigger a warning.
 * Does NOT lock them — admins can still edit.
 */
function _protectSystemSheets(ss) {
    const systemSheets = ['ORDERS', 'ORDERS_EXPORT', 'DAILY_OPERATIONS', 'DELETED_PRODUCTS'];
    systemSheets.forEach(name => {
        const sheet = ss.getSheetByName(name);
        if (!sheet) return;
        const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        if (existing.length === 0) {
            const prot = sheet.protect().setDescription('Order System — managed automatically');
            prot.setWarningOnly(true);
            Logger.log('[Installer] Warning-protection applied to: ' + name);
        }
    });
}


/**
 * Reorders sheets so the most-used ones appear first.
 */
function _setSheetOrder(ss) {
    const order = [
        'DASHBOARD', 'Welcome', 'CLIENT DATA', 'PRODUCTS',
        'ORDER_FORM_1', 'ORDER_FORM_2',
        'ORDERS', 'ORDERS_EXPORT', 'SETTINGS',
        'DAILY_OPERATIONS'
    ];
    order.forEach((name, idx) => {
        const sheet = ss.getSheetByName(name);
        if (sheet) ss.setActiveSheet(sheet).moveActiveSheet(idx + 1);
    });
    Logger.log('[Installer] Sheet order applied.');
}


// ─────────────────────────────────────────────────────────────
//  ADD-ON HOOKS (for when this runs as a Google Workspace Add-on)
// ─────────────────────────────────────────────────────────────

/**
 * onInstall — fires automatically when a user installs the add-on.
 * Runs the installer and adds the Order System menu.
 */
function onInstall(e) {
    onOpen(e);
    runInstaller();
}

/**
 * onOpen — fires every time the spreadsheet is opened.
 * Adds the Order System menu.
 */
function onOpen(e) {
    SpreadsheetApp.getUi()
        .createAddonMenu()        // Use createMenu('Order System') for editor add-ons
        .addItem('Launch Order Form', 'showOrderFormDialog')
        .addItem('Add / Edit Products', 'showAddProductSidebar')
        .addSeparator()
        .addItem('📄 Generate PDF for Selected Order', 'generateSelectedOrderPdf')
        .addSeparator()
        .addItem('⚙️  Run Installer / Repair', 'runInstaller')
        .addItem('📊 Refresh Daily Dashboard', 'refreshDailyOperationsDashboard')
        .addToUi();
}
