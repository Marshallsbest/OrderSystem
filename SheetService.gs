/**
 * SheetService.gs
 * Handles all direct interactions with the Google Spreadsheet
 */

const SHEET_NAMES = {
    WELCOME: "Welcome",
    ORDERS_EXPORT: "ORDERS_EXPORT",
    SETTINGS: "SETTINGS",
    ORDERS: "ORDERS",
    ORDER_PLACING: "ORDER_PLACING",
    CLIENT_DATA: "CLIENT DATA",
    PRODUCTS: "PRODUCTS"
};

/**
 * Get specific sheet by name
 */
function getSheet(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found. Please ensure the spreadsheet is set up correctly.`);
    }
    return sheet;
}

/**
 * Get all data from CLIENT DATA sheet
 * Assumes headers are in row 1
 * Returns array of objects keyed by header name
 */
function getClientData() {
    const sheet = getSheet(SHEET_NAMES.CLIENT_DATA);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove headers

    if (data.length === 0) return [];

    return data.map(row => {
        let client = {};
        headers.forEach((header, index) => {
            client[header] = row[index];
        });
        return client;
    });
}

/**
 * Find a client by ID
 */
function getClientById(clientId) {
    const clients = getClientData();
    // Adjust 'ClientID' to match exact header name in your sheet, assuming "ClientID"
    return clients.find(c => String(c['ClientID']) === String(clientId));
}
