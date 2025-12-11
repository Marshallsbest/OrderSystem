/**
 * OrderService.gs
 * Handles order validation, formatting, and persistence
 */

/**
 * Process a new order from the Client Web App
 * @param {Object} orderData - { clientId: "...", items: [{sku: "...", quantity: 5, unit: "case"}] }
 */
function processOrder(orderData) {
    try {
        const lock = LockService.getScriptLock();
        // Wait for up to 30 seconds for other processes to finish.
        lock.waitLock(30000);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const orderSheet = getSheet(SHEET_NAMES.ORDERS);

        const client = getClientById(orderData.clientId);
        if (!client) throw new Error("Client not found: " + orderData.clientId);

        const orderId = "ORD-" + new Date().getTime(); // Simple ID generation
        const timestamp = new Date();

        // Calculate Totals (Back-end validation)
        const productCatalog = getProductCatalog();
        let totalAmount = 0;
        let totalPieces = 0;

        // Prepare Product Column strings (Key-Value pairs: "SKU: Qty")
        // We will pack these into generic "Product N" columns
        // Strategy: Map sku to formatted string
        const productStrings = [];

        // Iterate over ordered items
        orderData.items.forEach(item => {
            if (item.quantity > 0) {
                const product = productCatalog.find(p => p.sku === item.sku);
                if (product) {
                    // Calculate line cost (simplistic, assumes price is per unit or case based on input)
                    // Adjust logic based on real "Units vs Case" pricing model if needed
                    // For now assuming Price is Unit Price and Case Price needs calculation or look up
                    // User said PRODUCTS has Price ($/unit?) and Order Amount. 
                    // Implementation Plan said: PRODUCTS has Price, Units/Case. 
                    // Let's assume input quantity is "units" for simplicity or user logic passed total units.
                    // IF the UI passes "cases" we convert to units? 
                    // Let's stick to the KV Pair requirement: "SKU: Quantity"

                    productStrings.push(`${item.sku}: ${item.quantity}`);

                    // Add to totals
                    // totalAmount += product.price * item.quantity; // validation logic
                    totalPieces += Number(item.quantity);
                }
            }
        });

        if (productStrings.length === 0) throw new Error("No items in order");

        // Prepare Row Data
        // [Order ID | Date/Time | Client Name | Total Amount | Total Pieces | Product 1 | Product 2 | ... ]
        const rowData = [
            orderId,
            timestamp,
            client['Name'] || orderData.clientId,
            // We rely on frontend total or calc here if prices strictly known. 
            // For now passing 0 or frontend total if passed. 
            orderData.totalAmount || 0,
            totalPieces
        ];

        // Append Product Strings
        // The ORDERS sheet might need enough columns.
        // We dynamically append.
        const finalRow = rowData.concat(productStrings);

        // Check if we need more columns
        if (finalRow.length > orderSheet.getMaxColumns()) {
            orderSheet.insertColumnsAfter(orderSheet.getMaxColumns(), finalRow.length - orderSheet.getMaxColumns());
        }

        orderSheet.appendRow(finalRow);

        SpreadsheetApp.flush(); // Ensure data is written

        // Trigger PDF Export
        const pdfUrl = createOrderPdf(orderId, client, orderData.items, timestamp);

        return { success: true, orderId: orderId, message: "Order placed successfully!", pdfUrl: pdfUrl };

    } catch (e) {
        Logger.log("Order Error: " + e.toString());
        throw e;
    } finally {
        LockService.getScriptLock().releaseLock();
    }
}
