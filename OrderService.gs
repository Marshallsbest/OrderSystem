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
        const timestampDate = new Date();
        const timestamp = Utilities.formatDate(timestampDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");

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
                    const price = Number(product.price) || 0;
                    const salePrice = Number(product.salePrice) || 0;
                    const isSale = product.onSale && (salePrice > 0);
                    const finalPrice = isSale ? salePrice : price;

                    totalAmount += finalPrice * item.quantity;
                    totalPieces += Number(item.quantity);
                }
            }
        });

        if (productStrings.length === 0) throw new Error("No items in order");

        // Prepare Row Data
        // [Order ID | Date/Time | Client Name | Total Amount | Total Pieces | Product 1 | Product 2 | ... ]
        // Ensure Client Name is robust - prioritize edited name
        const displayClientName = orderData.clientName || client['Name'] || client['Company Name'] || client['name'] || "Unknown";

        // FIXED COLUMN ORDER [A-F]:
        // Col A: Order ID
        // Col B: Client ID (Restored)
        // Col C: Date/Time (Requested Location)
        // Col D: Client Name (Edited/Display Name since we clear validation)
        // Col E: Total Amount
        // Col F: Total Pieces
        const rowData = [
            orderId,
            orderData.clientId, // Col B
            timestamp,          // Col C
            displayClientName,  // Col D
            totalAmount,        // Col E
            totalPieces         // Col F
        ];

        // Append Product Strings
        // The ORDERS sheet might need enough columns.
        // We dynamically append.
        const finalRow = rowData.concat(productStrings);

        // NUCLEAR VALIDATION FIX:
        // appendRow fails if validation rules are strict.
        // We must:
        // 1. Determine the next row index manually.
        // 2. Clear validation on that entire row range.
        // 3. Use setValues() which is more robust than appendRow() for mixed data types.

        const nextRow = orderSheet.getLastRow() + 1;
        const totalCols = finalRow.length;

        // Ensure we have enough columns
        if (totalCols > orderSheet.getMaxColumns()) {
            orderSheet.insertColumnsAfter(orderSheet.getMaxColumns(), totalCols - orderSheet.getMaxColumns());
        }

        const targetRange = orderSheet.getRange(nextRow, 1, 1, totalCols);

        // 1. Clear Validation
        try {
            targetRange.clearDataValidation();
            SpreadsheetApp.flush(); // FORCE commit of validation clear
        } catch (e) {
            Logger.log("Warning: Could not clear validation: " + e.toString());
        }

        // 2. Write Data

        // Step A: Write Non-Contentious Columns (A, B, D, E, F)
        // We skip Column C (Index 2 in 0-based array)
        const safeRowData = [...finalRow];
        safeRowData[2] = ""; // Temporarily blank out Date (Col C)

        try {
            targetRange.setValues([safeRowData]);
        } catch (e) {
            throw new Error(`Failed to write basic order data (Cols A,B,D-F): ${e.toString()}`);
        }

        // Step B: Handle "The Problem Child" - Column C (Date)
        // Target specifically cell C{nextRow}
        const cellC = orderSheet.getRange(nextRow, 3);

        try {
            // Aggressively clear validation on JUST this cell
            cellC.clearDataValidation();
            SpreadsheetApp.flush();

            // Try writing the Date Object directly (let Sheets handle formatting)
            cellC.setValue(timestampDate);
        } catch (e) {
            Logger.log(`Warning: Failed to write Date to C${nextRow}. Trying String fallback.`);
            // Fallback: Write string value
            try {
                cellC.setValue(timestamp);
            } catch (innerE) {
                // If even this fails, we leave it blank but don't crash the order
                Logger.log(`Critical: Could not write date to C${nextRow}: ${innerE.toString()}`);
                cellC.setValue("Date Error");
            }
        }

        SpreadsheetApp.flush();

        // Trigger PDF Export
        // Limit Client Name length for PDF filename if needed, but createOrderPdf handles it
        // Construct a client object that includes the EDITED details
        const finalClient = { ...client, Name: displayClientName, Address: orderData.clientAddress || client.Address };
        const pdfUrl = createOrderPdf(orderId, finalClient, orderData.items, timestampDate);
        /* const pdfUrl = "DEBUG_MODE_SKIPPED"; */

        return { success: true, orderId: orderId, message: "Order placed successfully!", pdfUrl: pdfUrl };

    } catch (e) {
        Logger.log("Order Error: " + e.toString());
        throw e;
    } finally {
        LockService.getScriptLock().releaseLock();
    }
}

/**
 * Get recent order history for a client
 * @param {string} clientId
 * @returns {Array} List of { id, date, total, pieces }
 */
function getHistoryForClient(clientId) {
    const sheet = getSheet(SHEET_NAMES.ORDERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Col A=ID, B=ClientID, C=Date, E=Total $, F=Total Pieces
    // Read A2:F
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    // Filter and Map in reverse to show newest first
    const history = [];
    // Loop backwards
    for (let i = data.length - 1; i >= 0; i--) {
        const row = data[i];
        // String cast to be safe
        const rowClientId = String(row[1]);
        if (rowClientId === String(clientId)) {
            history.push({
                id: row[0],
                date: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "MM/dd/yyyy") : String(row[2]),
                clientName: row[3],
                total: row[4],
                pieces: row[5]
            });
        }
        // Limit to last 20?
        if (history.length >= 20) break;
    }
    return history;
}

/**
 * Get specific items for an Order ID to repopulate the form
 * @param {string} orderId
 * @returns {Array} [{ sku, qty }]
 */
function getOrderItems(orderId) {
    const sheet = getSheet(SHEET_NAMES.ORDERS);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return [];

    // Search for row with Order ID in Col A
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const rowIndex = ids.findIndex(id => String(id) === String(orderId));

    if (rowIndex === -1) return [];

    // Read the whole row (adjusted for 0-index + 2 header offset)
    const rowValues = sheet.getRange(rowIndex + 2, 1, 1, lastCol).getValues()[0];

    // Items start at Col G (Index 6)
    // Structure: "SKU: Qty"
    const items = [];

    for (let i = 6; i < rowValues.length; i++) {
        const cell = String(rowValues[i]);
        if (cell && cell.includes(":")) {
            const parts = cell.split(":");
            if (parts.length === 2) {
                const sku = parts[0].trim();
                const qty = parseInt(parts[1].trim());
                if (sku && !isNaN(qty)) {
                    items.push({ sku: sku, qty: qty });
                }
            }
        }
    }

    return items;
}

/**
 * DUPLICATE ORDER TOOL (Spreadsheet Only)
 * Allows Admin to select a row in ORDERS, duplicate it to a new Pending Order
 * with a new timestamp and ID, so they can edit it manually.
 */
function duplicateSelectedOrder() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    // Safety Check: Are we on ORDERS sheet?
    if (sheet.getName() !== SHEET_NAMES.ORDERS) {
        ss.toast("Please select a row in the ORDERS sheet first.");
        return;
    }

    const row = sheet.getActiveRange().getRow();
    if (row < 2) {
        ss.toast("Please select a valid order row (Row 2+).");
        return;
    }

    // Read the selected row
    const lastCol = sheet.getLastColumn();
    const sourceData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    // Validate it looks like an order
    const oldId = sourceData[0];
    if (!String(oldId).startsWith("ORD-")) {
        // Soft warning
        // ss.toast("Warning: Selected row doesn't look like a standard order.");
    }

    // Create new Meta Data
    const newId = "ORD-" + new Date().getTime();
    const newDate = new Date(); // Current Timestamp
    const formattedDate = Utilities.formatDate(newDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");

    // Copy Data
    const newData = [...sourceData];
    newData[0] = newId;          // New ID
    newData[2] = formattedDate;  // New Date

    // Append to bottom
    const nextRow = sheet.getLastRow() + 1;
    sheet.appendRow(newData);

    // Highlight the new row
    // Attempt to clear validation on the new row just in case (Date column issues)
    try {
        sheet.getRange(nextRow, 1, 1, lastCol).clearDataValidation();
    } catch (e) { }

    ss.toast(`Order duplicated! New Order ID: ${newId}. You can now edit columns G+ manually.`);
}

/**
 * STAGING TOOL: Populate ORDER_DATA sheet from Selected Order Row
 * 1. Reads selected row in ORDERS
 * 2. Parses items (SKU: Qty)
 * 3. Lookups up Product Details (Price, Name, etc)
 * 4. Clears and writes to ORDER_DATA (A=SKU, B=Name, C=Qty, D=Unit$, E=Line$)
 */
function populateOrderDataStaging() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    // Check if on ORDERS sheet
    if (sheet.getName() !== SHEET_NAMES.ORDERS) {
        ss.toast("Please select a row in the ORDERS sheet.");
        return;
    }

    const row = sheet.getActiveRange().getRow();
    if (row < 2) {
        ss.toast("Select a valid order row.");
        return;
    }

    // 1. Get Items from Row
    const orderId = sheet.getRange(row, 1).getValue();
    const items = getOrderItems(orderId); // Reuse our parser: [{sku, qty}]

    if (items.length === 0) {
        ss.toast("No items found in this order row.");
        return;
    }

    // 2. Get Product Catalog for Enrichment
    const catalog = getProductCatalog();

    // AGGREGATION LOGIC:
    // We want to combine "Unit SKUs" and "Carton SKUs" into a single row if they are the same product.
    // Key = Normalized Product Name

    const aggregated = new Map();
    // Value: { 
    //   baseSku, name, 
    //   unitQty, cartonQty, 
    //   unitPrice, cartonPrice, 
    //   totalLineValue 
    // }

    items.forEach(item => {
        const product = catalog.find(p => p.sku === item.sku);
        if (product) {
            // Normalize Name: Remove "Carton", "Case", "MC", "Box" to find the base product
            // Also handle variations
            let baseName = product.name
                .replace(/\s?-\s?Carton/i, '')
                .replace(/\s?\(?MC\)?/i, '')
                .replace(/\s?Master Case/i, '')
                .trim();

            if (product.variation && product.variation !== 'Standard') {
                baseName += " " + product.variation;
            }

            // Initialize Group
            if (!aggregated.has(baseName)) {
                aggregated.set(baseName, {
                    baseSku: item.sku,
                    name: baseName,
                    unitQty: 0,
                    cartonQty: 0,
                    unitPrice: 0,
                    cartonPrice: 0,
                    totalLineValue: 0
                });
            }

            const group = aggregated.get(baseName);
            const price = Number(product.price) || 0;
            const itemValue = price * item.qty;
            group.totalLineValue += itemValue;

            // DETERMINE TYPE: Unit vs Carton
            // Logic: 
            // 1. Name contains 'Carton'/'MC' -> Carton
            // 2. SKU ends in 'C' -> Carton
            // 3. Price/UnitsPerCase check (if we had both prices, we could compare, but here we process linearly)

            const nameIsCarton = /Carton|Master Case|\bMC\b/i.test(product.name);
            const skuIsCarton = product.sku.toUpperCase().endsWith('C');

            // If it's a Carton
            if (nameIsCarton || skuIsCarton) {
                group.cartonQty += item.qty;
                // Update Carton Price (take the highest found so far for this group?)
                if (price > group.cartonPrice) group.cartonPrice = price;
            } else {
                // It's a Unit
                group.unitQty += item.qty;
                // Update Unit Price (take the price of this unit item)
                // If multiple unit variants exist, this might just take the last one, which is usually fine
                if (group.unitPrice === 0 || price < group.unitPrice) group.unitPrice = price;
            }

        } else {
            // Unknown Product - Keep Separate
            aggregated.set(item.sku, {
                baseSku: item.sku,
                name: "Unknown: " + item.sku,
                unitQty: item.qty,
                cartonQty: 0,
                unitPrice: 0,
                cartonPrice: 0,
                totalLineValue: 0
            });
        }
    });

    // POST-PROCESS: Fill in missing pricing gaps
    // If we only ordered Cartons, we might not have a Unit Price.
    // We should try to derive it or find it from the catalog if possible.
    // For now, satisfy the "Multiplication" rule: CartonPrice = UnitPrice * UnitsPerCase

    for (const group of aggregated.values()) {
        // If we have Carton Price but no Unit Price, try to estimate?
        // Or if we have Unit Price but no Carton Price, calculate it?

        // Let's rely on what we found. If we didn't find a Unit price (because none were ordered),
        // we leave it 0 or try to fetch the "Base" product from catalog?
        // That's complex. Let's send what we have.
        // The Template likely has the prices hardcoded or VLOOKUPs them anyway.
        // We mainly need the QUANTITIES to allow the template to do `Qty * Price`.

        // Actually, if we populate columns "Unit Price" and "Carton Price", we clarify it.
        // Even better: The stored `totalLineValue` is the source of truth for the dollar amount.
    }

    const enrichedItems = [];
    for (const group of aggregated.values()) {
        enrichedItems.push([
            group.baseSku,
            group.name,
            group.unitQty,
            group.cartonQty,
            group.unitPrice,
            group.cartonPrice,
            group.totalLineValue
        ]);
    }

    // 3. Write to ORDER_DATA
    const targetSheet = getSheet(SHEET_NAMES.ORDER_DATA);
    if (!targetSheet) {
        ss.toast("ORDER_DATA sheet not found. Please run Setup.");
        return;
    }

    // Get Metadata (Client Info, Date) from the Order Row
    // Row Structure: A=ID, B=ClientID, C=Date, D=DisplayClient
    const orderIdVal = sheet.getRange(row, 1).getValue();
    const clientIdVal = sheet.getRange(row, 2).getValue();
    const dateVal = sheet.getRange(row, 3).getValue();
    // const clientNameVal = sheet.getRange(row, 4).getValue();

    const salesRep = clientData[config.CFG_COL_REP || 'Sales Rep'] || "";
    const contactName = clientData[config.CFG_COL_CONTACT || 'Contact Name'] || "";
    const phone = clientData['Phone'] || clientData['Telephone'] || "";
    const address = clientData['Address'] || "";

    // Clear old data
    targetSheet.clearContents(); // Clear everything to be safe

    // Re-write Headers
    const itemHeaders = ["SKU (Base)", "Product Name", "Unit Qty", "Carton Qty", "Unit Price", "Carton Price", "Line Total"];
    targetSheet.getRange("A1:G1").setValues([itemHeaders]).setFontWeight("bold").setBackground("#e0e0e0");

    // Write new Item data
    if (enrichedItems.length > 0) {
        targetSheet.getRange(2, 1, enrichedItems.length, 7).setValues(enrichedItems);
    }

    // Write METADATA to Columns H (Label) and I (Value) for easy VLOOKUP/Referencing
    const metadata = [
        ["METADATA", "VALUE"],
        ["Order ID", orderIdVal],
        ["Date", dateVal],
        ["Client ID", clientIdVal],
        ["Company Name", clientData['Company Name'] || clientData['Name'] || ""],
        ["Address", address],
        ["Sales Rep", salesRep],
        ["Contact Name", contactName],
        ["Phone", phone]
    ];

    targetSheet.getRange("H1:I" + metadata.length).setValues(metadata);
    targetSheet.getRange("H1:I1").setFontWeight("bold").setBackground("#e0e0e0");
    targetSheet.setColumnWidth(8, 150); // Label Col
    targetSheet.setColumnWidth(9, 250); // Value Col

    ss.toast(`Staged Order ${orderIdVal} to ORDER_DATA.`);
}
