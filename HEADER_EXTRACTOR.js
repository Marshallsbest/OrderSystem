/**
 * HEADER_EXTRACTOR.gs
 * Run this to see exactly what headers the system is detecting.
 */
function getLiveProductHeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("PRODUCTS");
    if (!sheet) {
        console.error("Sheet 'PRODUCTS' not found.");
        return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log("LIVE HEADERS DETECTED:");
    console.log(JSON.stringify(headers));

    // Also put them in a UI alert for easy copying
    SpreadsheetApp.getUi().alert("Live Headers:\n\n" + JSON.stringify(headers));
}
