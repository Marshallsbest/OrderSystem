/**
 * Setup.gs
 * Version: v1.8.50
 * Structural setup functions removed as per user directive.
 */

/**
 * Handle External Template Creation
 */
function createExternalTemplates() {
    try {
        const clientSs = SpreadsheetApp.create("Order System - Client CRM Template");
        const clientSheet = clientSs.getSheets()[0];
        const clientHeaders = ["ClientID", "Company Name", "Address", "Sales Rep", "Contact Name", "Min Order"];
        clientSheet.getRange(2, 1, 1, clientHeaders.length).setValues([clientHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
        clientSheet.getRange(1, 1).setValue("CLIENT CRM DATABASE (DATA STARTS ON ROW 3)").setFontWeight("bold").setFontSize(12);

        const productSs = SpreadsheetApp.create("Order System - Product Master Template");
        const productSheet = productSs.getSheets()[0];
        const productHeaders = ["Node", "SKU", "Brand", "Product Name", "Parent Name", "Category", "Variation", "Variation 2", "Variation 3", "Price", "Sale Price", "Units per Case", "On Sale", "Description", "Image", "Color", "Ref", "Status", "PDF Range Name"];
        productSheet.getRange(1, 1, 1, productHeaders.length).setValues([productHeaders]).setFontWeight("bold").setBackground("#f3f3f3");

        return { success: true, clientUrl: clientSs.getUrl(), productUrl: productSs.getUrl() };
    } catch (e) {
        console.error("Template Creation Error:", e);
        return { success: false, error: e.message };
    }
}

/**
 * Initialize DAILY_OPERATIONS sheet
 * Used as a fallback by AnalyticsService
 */
function setupDailyOperationsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.DAILY_OPERATIONS);
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAMES.DAILY_OPERATIONS);
    }
    sheet.clear();
    refreshDailyOperationsDashboard();
    return sheet;
}
