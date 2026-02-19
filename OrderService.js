/**
 * OrderService.gs
 * Version: v0.8.66 - Added ADDRESS column support
 * 
 * CURRENT HEADERS:
 * A(0)=Version | B(1)=INVOICE_NUMBER | C(2)=TIME STAMP | D(3)=TOTAL UNITS
 * E(4)=COMISSION | F(5)=TOTAL | G(6)=CLIENT | H(7)=COMMENT | I(8)=ADDRESS | J(9+)=Products
 */

const ORDER_COL = {
    VERSION: 0,
    INVOICE_NUMBER: 1,
    TIME_STAMP: 2,
    TOTAL_UNITS: 3,
    COMMISSION: 4,
    TOTAL: 5,
    CLIENT: 6,
    COMMENT: 7,
    ADDRESS: 8,
    PRODUCTS_START: 9
};

function processOrder(orderData) {
    try {
        const lock = LockService.getScriptLock();
        lock.waitLock(30000);

        const orderSheet = getSheet(SHEET_NAMES.ORDERS);
        const client = getClientById(orderData.clientId);
        if (!client) throw new Error("Client not found: " + orderData.clientId);

        const productCatalog = getProductCatalog();
        const orderModel = createOrderModel(orderData);

        let totalAmount = 0;
        let totalPieces = 0;
        let totalCommission = 0;
        let hasSale = false;
        const itemsToStaging = [];

        orderModel.line_items.forEach(item => {
            if (item.quantity > 0) {
                const normalizedItemSku = String(item.sku || "").trim().toUpperCase();
                const product = productCatalog.find(p => String(p.sku || "").trim().toUpperCase() === normalizedItemSku);
                if (product) {
                    const isProductOnSale = product.onSale && (product.salePrice > 0);
                    const finalPrice = isProductOnSale ? product.salePrice : product.price;
                    if (isProductOnSale) hasSale = true;
                    totalAmount += finalPrice * item.quantity;

                    const units = parseInt(product.unitsPerCase) || 1;
                    totalPieces += units * item.quantity;

                    const rate = isProductOnSale
                        ? (parseFloat(product.saleCommission) || 1.0)
                        : (parseFloat(product.commissionRate) || 1.5);
                    totalCommission += rate * item.quantity * units;

                    itemsToStaging.push({
                        sku: item.sku,
                        quantity: item.quantity,
                        price: finalPrice
                    });
                }
            }
        });

        if (itemsToStaging.length === 0) throw new Error("No items in order");

        const targetRow = orderSheet.getLastRow() + 1;

        // Determine version label (Original vs Rev:X)
        let versionLabel = "Original";
        let finalOrderId = orderModel.id;

        // Check if this is an edit of an existing order
        const editingOrderId = orderData.editOrderId || orderData.originalOrderId;
        if (editingOrderId) {
            // This is a revision - count existing revisions for this order
            const allData = orderSheet.getDataRange().getValues();
            const baseInvoice = editingOrderId.replace(/^Rev:\d+\s*/, '').trim(); // Strip any existing Rev: prefix

            let revisionCount = 0;
            for (let i = 1; i < allData.length; i++) {
                const rowVersion = String(allData[i][ORDER_COL.VERSION] || '');
                const rowInvoice = String(allData[i][ORDER_COL.INVOICE_NUMBER] || '');

                // Count rows that are revisions of this order OR the original
                if (rowInvoice === baseInvoice || rowInvoice === editingOrderId) {
                    if (rowVersion.startsWith('Rev:')) {
                        const revNum = parseInt(rowVersion.replace('Rev:', '')) || 0;
                        if (revNum > revisionCount) revisionCount = revNum;
                    }
                }
            }

            versionLabel = "Rev:" + (revisionCount + 1);
            finalOrderId = baseInvoice; // Keep same invoice number for tracking
        }

        // Build row matching current headers
        const baseRowData = [
            versionLabel,                      // A(0): Version (Original or Rev:X)
            finalOrderId,                      // B(1): INVOICE_NUMBER
            new Date(),                        // C(2): TIME STAMP
            totalPieces,                       // D(3): TOTAL UNITS
            totalCommission,                   // E(4): COMISSION
            totalAmount,                       // F(5): TOTAL
            orderData.clientName || "Unknown", // G(6): CLIENT
            orderData.clientComments || "",    // H(7): COMMENT
            orderData.clientAddress || ""      // I(8): ADDRESS
        ];

        const itemStrings = itemsToStaging.map(item => {
            const product = productCatalog.find(p => p.sku === item.sku);
            const isSaleFlag = (product && product.onSale) ? 'T' : 'F';
            return `[${item.quantity}|@${item.sku}|$${item.price.toFixed(2)}|${isSaleFlag}]`;
        });

        const finalRowData = baseRowData.concat(itemStrings);
        orderSheet.getRange(targetRow, 1, 1, finalRowData.length).setValues([finalRowData]);

        let pdfUrl = "";
        try {
            pdfUrl = generateOrderPdf({
                id: finalOrderId,
                clientName: orderData.clientName || "Unknown Client",
                clientAddress: orderData.clientAddress || "",
                clientComments: orderData.clientComments || "",
                date: new Date(),
                total: totalAmount,
                items: itemsToStaging
            });
        } catch (e) { }

        SpreadsheetApp.flush();
        return { success: true, orderId: finalOrderId, total: totalAmount.toFixed(2), pdfUrl: pdfUrl };
    } catch (e) {
        throw e;
    } finally {
        LockService.getScriptLock().releaseLock();
    }
}

function getOrderById(orderId) {
    if (!orderId) return null;
    const sheet = getSheet(SHEET_NAMES.ORDERS);
    const data = sheet.getDataRange().getValues();

    // Search by Invoice (Index 1)
    let rowContent = data.find(r => String(r[ORDER_COL.INVOICE_NUMBER]).trim() === String(orderId));

    // Fallback: broad search
    if (!rowContent) {
        rowContent = data.find(r => r.some(cell => String(cell).trim() === String(orderId)));
    }
    if (!rowContent) return null;

    const items = {};
    for (let i = ORDER_COL.PRODUCTS_START; i < rowContent.length; i++) {
        const cell = String(rowContent[i]);
        if (cell.includes("[") && cell.includes("|")) {
            try {
                const parts = cell.replace(/[\[\]]/g, '').split('|');
                if (parts.length >= 2) {
                    const qty = parseInt(parts[0]);
                    const sku = parts[1].replace('@', '').trim();
                    const price = parts[2] ? parseFloat(parts[2].replace('$', '')) : 0;
                    if (sku && !isNaN(qty)) items[sku] = { qty: qty, price: price };
                }
            } catch (e) { }
        }
    }

    return {
        id: rowContent[ORDER_COL.INVOICE_NUMBER],
        items: items,
        clientName: rowContent[ORDER_COL.CLIENT] || "",
        clientComments: rowContent[ORDER_COL.COMMENT] || "",
        clientAddress: rowContent[ORDER_COL.ADDRESS] || ""
    };
}

function getOrdersByClient(clientName) {
    try {
        const sheet = getSheet(SHEET_NAMES.ORDERS);
        const data = sheet.getDataRange().getValues();
        const rows = data.slice(1);
        const allOrders = [];

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const invoiceId = String(row[ORDER_COL.INVOICE_NUMBER] || '').trim();
            const client = String(row[ORDER_COL.CLIENT] || '').trim();
            const total = parseFloat(String(row[ORDER_COL.TOTAL] || 0).replace(/[$,]/g, '')) || 0;

            if (!invoiceId && !client && total === 0) continue;
            if (clientName && client.toLowerCase() !== clientName.toLowerCase()) continue;

            allOrders.push({
                id: invoiceId || ('ROW-' + (i + 2)),
                clientName: client || 'Unknown',
                total: total,
                pieces: parseInt(row[ORDER_COL.TOTAL_UNITS]) || 0,
                timestamp: row[ORDER_COL.TIME_STAMP] instanceof Date
                    ? row[ORDER_COL.TIME_STAMP].toISOString()
                    : String(row[ORDER_COL.TIME_STAMP] || new Date().toISOString()),
                comments: String(row[ORDER_COL.COMMENT] || ''),
                address: String(row[ORDER_COL.ADDRESS] || ''),
                state: 'New'
            });
        }

        allOrders.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        console.log(`[getOrdersByClient] Returning ${allOrders.length} orders`);
        return allOrders;
    } catch (e) {
        console.error("getOrdersByClient Error:", e);
        return [];
    }
}
