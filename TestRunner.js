/**
 * TEST: Verify PDF Folder Structure and Grid Layout
 * Run this function from the Apps Script editor to verify v0.9.17 changes.
 */
function TEST_PDF_OUTPUT() {
    const testData = {
        id: "TEST-INV-999",
        clientName: "Antigravity Test Corp",
        clientAddress: "123 Starship Drive, Mars Colony 1, 99999",
        clientComments: "This is a test of the new monthly folder structure and grid layout.",
        salesRep: "Luke (AI Test)",
        date: new Date(),
        total: 155.50,
        items: [
            { sku: "SOL-Pnl-400W", quantity: 2, price: 50.00 },
            { sku: "BATT-LIT-100", quantity: 1, price: 55.50 }
        ]
    };

    console.log("Starting PDF Test...");
    try {
        const pdfUrl = generateOrderPdf(testData);
        console.log("Test Success! PDF generated at: " + pdfUrl);

        // Show the results to the user if running in a container
        try {
            const ui = SpreadsheetApp.getUi();
            ui.alert("Test Complete", "PDF generated and saved to monthly folder.\nURL: " + pdfUrl, ui.ButtonSet.OK);
        } catch (e) {
            // Not running in Spreadsheet UI context, just log it
        }

        return pdfUrl;
    } catch (e) {
        console.error("Test Failed: " + e.toString());
        throw e;
    }
}
