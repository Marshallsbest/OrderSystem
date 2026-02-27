function CHECK_ORDERS_HEADERS() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("ORDERS");
    if (!sheet) {
        console.log("Sheet ORDERS not found!");
        return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log("HEADERS: " + JSON.stringify(headers));

    // Also check column J for the product format
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const sample = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
        console.log("SAMPLE ROW: " + JSON.stringify(sample));
    }
}
