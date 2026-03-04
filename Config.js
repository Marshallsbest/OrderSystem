/**
 * Config.gs
 * Global Constants & Shared Utilities
 */

const APP_TITLE = "Order System";
const CURRENT_VERSION = "v0.9.21";

const SHEET_NAMES = {
    DASHBOARD: "DASHBOARD",
    WELCOME: "Welcome",
    ORDERS_EXPORT: "ORDERS_EXPORT",
    SETTINGS: "SETTINGS",
    ORDERS: "ORDERS",
    CLIENT_DATA: "CLIENT DATA",
    PRODUCTS: "PRODUCTS",
    DELETED_PRODUCTS: "DELETED_PRODUCTS",
    DAILY_OPERATIONS: "DAILY_OPERATIONS",
    EXPORT_SUMMARY: "EXPORT",
    CLIENT_INFO_UPDATES: "CLIENT_INFO_UPDATES"
};

/**
 * Order Form PDF — Brand Colour Config
 * Edit these to change the look of all generated Order Form PDFs.
 * Layout is still read dynamically from ORDER_FORM_1; only colours are fixed here.
 */
const ORDER_FORM_COLORS = {
    categoryBg: '#cc66cc', // Pink/magenta category header background
    categoryText: '#ffffff', // White text on category headers
    mcPriceBg: '#ffff00', // Yellow highlight for MC price cells
    totalText: '#1a6b2a', // Green for Total $ values
    accentBorder: '#b050b0', // Darker purple for borders / header outline
};

/**
 * Shared Utilities
 */
const superNormalize = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, '').replace(/s$/, '');

function columnToLetter(column) {
    if (column < 1) return "A";
    let temp, letter = "";
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

function getSheet(sheetName) {
    // In web app context, getActiveSpreadsheet() may fail or return wrong spreadsheet.
    // For container-bound scripts, we can use getActive() which should work.
    // If this fails, fall back to opening by ID from script properties.
    let ss;
    try {
        ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
        // Fallback: Try to get the spreadsheet this script is bound to
        const scriptId = ScriptApp.getScriptId();
        const file = DriveApp.getFileById(scriptId);
        const parentId = file.getParents().next().getId();
        ss = SpreadsheetApp.openById(parentId);
    }

    if (!ss) {
        throw new Error("Could not access the spreadsheet. Please ensure the script is bound to a spreadsheet.");
    }

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        // Fallback: Case-insensitive search
        const sheets = ss.getSheets();
        sheet = sheets.find(s => s.getName().toLowerCase() === sheetName.toLowerCase());
    }
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found. Available: ${ss.getSheets().map(s => s.getName()).join(', ')}`);
    return sheet;
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Return the sheet name for a given Order Form number.
 * Reads from SETTINGS: key = "FORM_{n}_SHEET", value = sheet name.
 * Falls back to "ORDER_FORM_{n}" if not configured.
 */
function getOrderFormSheetName(formNum) {
    const key = 'FORM_' + formNum + '_SHEET';
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
        if (settingsSheet) {
            const data = settingsSheet.getDataRange().getValues();
            for (let i = 0; i < data.length; i++) {
                if (String(data[i][0]).trim().toUpperCase() === key.toUpperCase()) {
                    const val = String(data[i][1] || '').trim();
                    if (val) return val;
                }
            }
        }
    } catch (e) { /* ignore — fall through to default */ }
    return 'ORDER_FORM_' + formNum;
}

/**
 * Return all configured Order Form template mappings for the Admin UI.
 * Reads every "FORM_N_SHEET" row from SETTINGS.
 * Always includes at least Form 1 as a default.
 */
function getOrderFormTemplates() {
    const templates = [];
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
        if (settingsSheet) {
            const data = settingsSheet.getDataRange().getValues();
            data.forEach(row => {
                const key = String(row[0] || '').trim();
                const val = String(row[1] || '').trim();
                const m = key.match(/^FORM_(\d+)_SHEET$/i);
                if (m && val) {
                    templates.push({
                        formNum: m[1], sheetName: val,
                        label: 'Form ' + m[1] + ' \u2014 ' + val
                    });
                }
            });
        }
    } catch (e) { /* ignore */ }
    if (templates.length === 0) {
        templates.push({
            formNum: '1', sheetName: 'ORDER_FORM_1',
            label: 'Form 1 \u2014 ORDER_FORM_1'
        });
    }
    return templates;
}
