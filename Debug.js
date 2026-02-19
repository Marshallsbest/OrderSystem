/**
 * DEBUG / VERIFICATION PROTCOL
 * Runs a simulated order to verify the multi-cell save fix.
 */
function VERIFY_ORDER_SAVE_FIX() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // 1. Setup dummy data
    const testOrderId = "TEST-" + new Date().getTime();
    const testData = {
        clientId: "100001",
        clientName: "Gravity Test Client",
        clientComments: "Checking multi-cell placement",
        items: [
            { sku: "TEST-SKU-A", quantity: 5, price: 10.50 },
            { sku: "TEST-SKU-B", quantity: 2, price: 50.00 }
        ]
    };

    ui.alert("Starting Verification", "This will generate a test order in the ORDERS sheet. Review Column J onwards.", ui.ButtonSet.OK);

    try {
        // 2. Execute processOrder
        // Mocking the clientId lookup if needed, but and processOrder handles fallback
        const result = processOrder(testData);

        // 3. Check the sheet
        const sheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
        const lastRow = sheet.getLastRow();
        const rowValues = sheet.getRange(lastRow, 1, 1, 12).getValues()[0]; // Check first 12 cols

        const msg = [
            "✅ Process Complete!",
            "Order ID: " + result.orderId,
            "Row Index: " + lastRow,
            "Col J (10): " + rowValues[9],
            "Col K (11): " + rowValues[10],
            "\nIf Col J and Col K both contain '{#...}', the fix is working!"
        ].join("\n");

        ui.alert("Verification Result", msg, ui.ButtonSet.OK);

    } catch (e) {
        ui.alert("Verification Failed", e.toString(), ui.ButtonSet.OK);
    }
}

/**
 * DEBUG: Test Order History Retrieval
 * Run this from Extensions > Apps Script > Run to diagnose history loading.
 */
function DEBUG_testOrderHistoryRetrieval() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
        let sheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
        if (!sheet) {
            const sheets = ss.getSheets();
            sheet = sheets.find(s => s.getName().toLowerCase() === SHEET_NAMES.ORDERS.toLowerCase());
        }

        if (!sheet) {
            const allNames = ss.getSheets().map(s => s.getName()).join(", ");
            ui.alert("Sheet Not Found", `Could not find sheet named "${SHEET_NAMES.ORDERS}".\n\nAvailable sheets: ${allNames}`, ui.ButtonSet.OK);
            return;
        }

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const rows = data.slice(1);

        // Show first 8 header values
        const headerReport = headers.slice(0, 8).map((h, i) => `[${i}]: "${h}"`).join(" | ");

        // Show raw values at NEW indices (headers say: 1=Invoice, 6=Client, 5=Total)
        // AND old indices (6=Invoice, 7=Client, 5=Total)
        let sampleRows = "";
        for (let i = 0; i < Math.min(3, rows.length); i++) {
            const row = rows[i];
            sampleRows += `Row ${i + 2}:\n`;
            sampleRows += `  NEW: Invoice[1]="${row[1]}", Client[6]="${row[6]}", Total[5]="${row[5]}"\n`;
            sampleRows += `  OLD: Invoice[6]="${row[6]}", Client[7]="${row[7]}", Total[5]="${row[5]}"\n`;
        }

        const result = getOrdersByClient('');

        const msg = [
            `Sheet: "${sheet.getName()}" | Data Rows: ${rows.length}`,
            ``,
            `Headers: ${headerReport}`,
            ``,
            `--- Sample Data ---`,
            sampleRows,
            ``,
            `getOrdersByClient('') returned: ${result ? result.length : 0} orders`
        ].join("\n");

        ui.alert("Order History Debug", msg, ui.ButtonSet.OK);

    } catch (e) {
        ui.alert("Debug Error", e.toString() + "\n\nStack: " + e.stack, ui.ButtonSet.OK);
    }
}

