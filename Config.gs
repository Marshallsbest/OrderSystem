/**
 * Config.gs
 * Global Constants & Shared Utilities
 */

const APP_TITLE = "Order System";
const CURRENT_VERSION = "v1.8.00";

const SHEET_NAMES = {
    DASHBOARD: "DASHBOARD",
    WELCOME: "Welcome",
    ORDERS_EXPORT: "ORDERS_EXPORT",
    SETTINGS: "SETTINGS",
    ORDERS: "ORDERS",
    ORDER_PLACING: "ORDER_PLACING",
    CLIENT_DATA: "CLIENT DATA",
    PRODUCTS: "PRODUCTS",
    ORDER_DATA: "ORDER_DATA",
    DELETED_PRODUCTS: "DELETED_PRODUCTS",
    DAILY_OPERATIONS: "DAILY_OPERATIONS",
    EXPORT_SUMMARY: "EXPORT",
    STAGING_SOURCE: "STAGING_SOURCE"
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
    return sheet;
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
