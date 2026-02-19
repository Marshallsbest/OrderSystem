/**
 * Config.gs
 * Global Constants & Shared Utilities
 */

const APP_TITLE = "Order System";
const CURRENT_VERSION = "v0.9.16";

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
