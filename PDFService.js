/**
 * PDFService.gs
 * Generates styled PDF invoices for orders
 * Version: v1.8.25
 */

// ============================================================================
// FOLDER MANAGEMENT
// ============================================================================

/**
 * Get or create the Orders folder for PDF storage
 * Checks SETTINGS for custom folder URL first, otherwise creates "Orders" folder
 */
function getOrdersFolder() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Check for custom folder URL in SETTINGS
    const customFolderUrl = getSettingValue("PDF_FOLDER_URL");
    if (customFolderUrl && customFolderUrl.trim() !== "") {
        try {
            // Extract folder ID from URL
            const match = customFolderUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
            if (match && match[1]) {
                const folder = DriveApp.getFolderById(match[1]);
                return folder;
            }
        } catch (e) {
            Logger.log("Custom folder URL invalid, falling back to default: " + e.message);
        }
    }

    // Default: Create/get "Orders" folder in same location as spreadsheet
    const ssFile = DriveApp.getFileById(ss.getId());
    const parents = ssFile.getParents();
    let parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

    const folderName = "Orders";
    const folders = parentFolder.getFoldersByName(folderName);

    if (folders.hasNext()) {
        return folders.next();
    } else {
        const newFolder = parentFolder.createFolder(folderName);
        Logger.log("Created 'Orders' folder: " + newFolder.getUrl());
        return newFolder;
    }
}

/**
 * Create the Orders folder during initial setup
 */
function setupOrdersFolder() {
    const folder = getOrdersFolder();
    SpreadsheetApp.getActiveSpreadsheet().toast("Orders folder ready: " + folder.getName());
    return folder;
}

/**
 * Helper: Get a setting value from the SETTINGS sheet
 */
function getSettingValue(key) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === key.toUpperCase()) {
            return data[i][1];
        }
    }
    return null;
}

// ============================================================================
// PDF GENERATION
// ============================================================================

/**
 * Generate a PDF invoice for an order
 * @param {Object} orderData - The order data object
 * @returns {string} URL of the generated PDF
 */
function generateOrderPdf(orderData) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
    } catch (e) {
        throw new Error("System is busy. Please try again in a few seconds.");
    }

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const folder = getOrdersFolder();

        // Format date: "Feb 4th 2026"
        const orderDate = orderData.date instanceof Date ? orderData.date : new Date();
        const formattedDate = formatDateWithOrdinal(orderDate);

        // File name: "Client Name - Feb 4th 2026.pdf"
        const clientName = String(orderData.clientName || "Unknown Client").trim();
        const fileName = `${clientName} - ${formattedDate}.pdf`;

        // Build the PDF content HTML
        const htmlContent = buildInvoiceHtml(orderData, formattedDate);

        // Create temporary Google Doc, convert to PDF
        const pdfBlob = convertHtmlToPdf(htmlContent, fileName);

        // Save to Orders folder
        const pdfFile = folder.createFile(pdfBlob);

        Logger.log("PDF created: " + pdfFile.getUrl());
        return pdfFile.getUrl();

    } finally {
        lock.releaseLock();
    }
}

/**
 * Format date with ordinal suffix: "Feb 4th 2026"
 */
function formatDateWithOrdinal(date) {
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const day = date.getDate();
    const month = months[date.getMonth()];
    const year = date.getFullYear();

    // Add ordinal suffix
    let suffix = "th";
    if (day === 1 || day === 21 || day === 31) suffix = "st";
    else if (day === 2 || day === 22) suffix = "nd";
    else if (day === 3 || day === 23) suffix = "rd";

    return `${month} ${day}${suffix} ${year}`;
}

/**
 * Format an address string for HTML display with proper line breaks.
 * Handles Canadian addresses where components may be concatenated without separators.
 * Input:  "140 little st SBlenheimONN0P1A0" or "140 Little St S, Blenheim ON, N0P 1A0"
 * Output: "140 Little St S,<br>Blenheim, ON<br>N0P 1A0"
 */
function formatAddress(address) {
    if (!address) return "";
    let addr = String(address).trim();

    // If already has newlines, convert to <br> and return
    if (addr.includes("\n")) {
        return addr.split("\n").map(s => s.trim()).filter(s => s).join("<br>");
    }

    // If already has commas, split on commas and format
    if (addr.includes(",")) {
        return addr.split(",").map(s => s.trim()).filter(s => s).join(",<br>");
    }

    // Canadian postal code pattern: A1A 1A1 (with or without space)
    const postalMatch = addr.match(/([A-Za-z]\d[A-Za-z])\s*(\d[A-Za-z]\d)\s*$/);
    if (postalMatch) {
        const postalCode = postalMatch[1].toUpperCase() + " " + postalMatch[2].toUpperCase();
        let beforePostal = addr.substring(0, postalMatch.index).trim();

        // Canadian province abbreviations (2-letter)
        const provinces = ['ON', 'QC', 'BC', 'AB', 'MB', 'SK', 'NB', 'NS', 'PE', 'NL', 'NT', 'NU', 'YT'];

        // Try to extract province abbreviation right before postal code
        let province = "";
        let city = "";
        let street = beforePostal;

        for (const prov of provinces) {
            // Check if province appears at end of beforePostal (case-insensitive)
            const provRegex = new RegExp('\\b(' + prov + ')\\s*$', 'i');
            const provMatch = beforePostal.match(provRegex);
            if (provMatch) {
                province = prov;
                street = beforePostal.substring(0, provMatch.index).trim();
                break;
            }
            // Also check without word boundary (concatenated case like "BlenheimON")
            const concatRegex = new RegExp('(' + prov + ')\\s*$', 'i');
            const concatMatch = beforePostal.match(concatRegex);
            if (concatMatch) {
                province = prov;
                street = beforePostal.substring(0, concatMatch.index).trim();
                break;
            }
        }

        // Try to split street from city
        // Look for common street suffixes to find where street ends and city begins
        const streetSuffixes = /\b(st|street|ave|avenue|blvd|boulevard|dr|drive|rd|road|cres|crescent|ct|court|pl|place|way|ln|lane|hwy|highway|cir|circle|pkwy|parkway)\b\s*(N|S|E|W|NE|NW|SE|SW)?\s*/i;
        const streetSuffixMatch = street.match(streetSuffixes);

        if (streetSuffixMatch) {
            // Find end of the street suffix (including direction if present)
            const suffixEnd = streetSuffixMatch.index + streetSuffixMatch[0].length;
            const streetPart = street.substring(0, suffixEnd).trim();
            const cityPart = street.substring(suffixEnd).trim();

            let parts = [streetPart];
            if (cityPart) {
                if (province) {
                    parts.push(cityPart + ", " + province);
                } else {
                    parts.push(cityPart);
                }
            } else if (province) {
                parts.push(province);
            }
            parts.push(postalCode);
            return parts.join("<br>");
        }

        // Fallback: Just split street + province + postal
        let parts = [street];
        if (province) parts[parts.length - 1] += ", " + province;
        parts.push(postalCode);
        return parts.join("<br>");
    }

    // No postal code detected - just return as-is
    return addr;
}

/**
 * Build the HTML content for the invoice
 */
function buildInvoiceHtml(orderData, formattedDate) {
    const clientName = orderData.clientName || "Unknown Client";
    const clientAddress = formatAddress(orderData.clientAddress || "");
    const comments = orderData.clientComments || "";
    const items = orderData.items || [];
    const salesRep = orderData.salesRep || "";

    // Get product catalog for full product details
    const catalog = getProductCatalog();

    // Build enriched items with product details
    const enrichedItems = items.map(item => {
        const product = catalog.find(p =>
            String(p.sku || "").trim().toUpperCase() === String(item.sku || "").trim().toUpperCase()
        );

        if (!product) {
            return {
                sku: item.sku,
                name: item.sku,
                variation: "",
                category: "Unknown",
                quantity: item.quantity,
                price: item.price || 0,
                salePrice: 0,
                onSale: false,
                sortKey: "ZZZ_Unknown"
            };
        }

        return {
            sku: product.sku,
            name: product.name,
            variation: [product.variation, product.variation2, product.variation3].filter(v => v).join(" "),
            category: product.category || "Uncategorized",
            quantity: item.quantity,
            price: product.price || 0,
            salePrice: product.salePrice || 0,
            onSale: !!product.onSale,
            sortKey: `${product.name}_${product.category}`
        };
    });

    // Sort: Group by Product Name, then by Category
    enrichedItems.sort((a, b) => a.sortKey.localeCompare(b.sortKey));

    // Calculate order total
    let orderTotal = 0;
    enrichedItems.forEach(item => {
        const finalPrice = item.onSale ? item.salePrice : item.price;
        orderTotal += finalPrice * item.quantity;
    });

    // Build item rows HTML
    // Columns: Item Name | Regular Price | Sale Price | Qty | Line Subtotal
    let itemRowsHtml = "";
    enrichedItems.forEach(item => {
        const productDisplay = item.name + (item.variation ? " " + item.variation : "");
        const finalPrice = item.onSale ? item.salePrice : item.price;
        const lineSubtotal = finalPrice * item.quantity;

        // Regular price: crossed out if on sale, normal otherwise
        const regularPriceDisplay = item.onSale
            ? `<span style="text-decoration:line-through;color:#999;">$${item.price.toFixed(2)}</span>`
            : `$${item.price.toFixed(2)}`;

        // Sale price: shown only if on sale, blank otherwise
        const salePriceDisplay = item.onSale
            ? `<strong style="color:#e53935;">$${item.salePrice.toFixed(2)}</strong>`
            : '';

        itemRowsHtml += `
            <tr>
                <td style="padding:8px;">${productDisplay}</td>
                <td style="padding:8px;text-align:right;">${regularPriceDisplay}</td>
                <td style="padding:8px;text-align:right;">${salePriceDisplay}</td>
                <td style="padding:8px;text-align:center;">${item.quantity}</td>
                <td style="padding:8px;text-align:right;">$${lineSubtotal.toFixed(2)}</td>
            </tr>
        `;
    });

    // Sales Rep display
    const salesRepHtml = salesRep
        ? `<div class="sales-rep">Sales Rep: ${salesRep}</div>`
        : '';

    // Build full HTML document
    const html = `
<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            font-family: 'Roboto', Arial, sans-serif;
            margin: 40px;
            color: #333;
        }
        .header {
            margin-bottom: 30px;
            border-bottom: 2px solid #006c4c;
            padding-bottom: 20px;
        }
        .client-name {
            font-size: 24px;
            font-weight: bold;
            color: #006c4c;
            margin-bottom: 5px;
        }
        .client-address {
            font-size: 14px;
            color: #666;
        }
        .order-date {
            font-size: 12px;
            color: #999;
            margin-top: 10px;
        }
        .sales-rep {
            font-size: 13px;
            color: #006c4c;
            font-weight: 600;
            margin-top: 8px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th {
            background: #f5f5f5;
            padding: 12px 8px;
            text-align: left;
            border-bottom: 2px solid #ddd;
            font-weight: 600;
        }
        th:nth-child(2) { width: 100px; text-align: right; }
        th:nth-child(3) { width: 100px; text-align: right; }
        th:nth-child(4) { width: 60px; text-align: center; }
        th:nth-child(5) { width: 110px; text-align: right; }
        td {
            border-bottom: 1px solid #eee;
        }
        .subtotal-row {
            background: #f9f9f9;
            font-weight: bold;
        }
        .subtotal-row td {
            padding: 12px 8px;
            border-top: 2px solid #006c4c;
        }
        .comments-section {
            margin-top: 30px;
            padding: 15px;
            background: #f5f5f5;
            border-radius: 8px;
        }
        .comments-label {
            font-weight: bold;
            color: #006c4c;
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="client-name">${clientName}</div>
        <div class="client-address">${clientAddress}</div>
        <div class="order-date">Order Date: ${formattedDate}</div>
        ${salesRepHtml}
    </div>
    
    <table>
        <thead>
            <tr>
                <th>Item</th>
                <th>Price</th>
                <th>Sale Price</th>
                <th>Qty</th>
                <th>Subtotal</th>
            </tr>
        </thead>
        <tbody>
            ${itemRowsHtml}
            <tr class="subtotal-row">
                <td style="text-align:left; font-size:12px; color:#666;">Line Items: ${enrichedItems.length}</td>
                <td colspan="3" style="text-align:right;">Order Total:</td>
                <td style="text-align:right;">$${orderTotal.toFixed(2)}</td>
            </tr>
        </tbody>
    </table>

    ${comments ? `
    <div class="comments-section">
        <span class="comments-label">Comments:</span> ${comments}
    </div>
    ` : ''}
</body>
</html>
    `;

    return html;
}

/**
 * Convert HTML to PDF blob
 */
function convertHtmlToPdf(htmlContent, fileName) {
    // Create a temporary Google Doc
    const doc = DocumentApp.create("TempInvoice_" + new Date().getTime());
    const docId = doc.getId();

    try {
        // For HTML styling, we'll use a different approach:
        // Create a Blob from HTML and use Drive API to convert
        const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'invoice.html');

        // Use URL fetch to convert via Google's export endpoint
        // First, we need to create a temporary spreadsheet for this approach
        // Actually, let's use a simpler method - create HTML as a standalone page

        // Simplified approach: Create PDF from HTML blob directly
        const pdfBlob = htmlBlob.getAs('application/pdf').setName(fileName);

        return pdfBlob;

    } finally {
        // Cleanup temp doc
        try {
            DriveApp.getFileById(docId).setTrashed(true);
        } catch (e) { }
    }
}

// ============================================================================
// TRIGGER FUNCTIONS (Called from Menu/UI)
// ============================================================================

/**
 * Generate PDF for the selected order in ORDERS sheet
 */
function generateSelectedOrderPdf() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== SHEET_NAMES.ORDERS) {
        ss.toast("Please run this tool from the ORDERS sheet.");
        return;
    }

    const activeRow = sheet.getActiveCell().getRow();
    if (activeRow < 2) {
        ss.toast("Please select an order row.");
        return;
    }

    const rowData = sheet.getRange(activeRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Column mapping matching OrderService.js ORDER_COL constants:
    // A(0)=Version | B(1)=INVOICE_NUMBER | C(2)=TIME STAMP | D(3)=TOTAL UNITS
    // E(4)=COMMISSION | F(5)=TOTAL | G(6)=CLIENT | H(7)=COMMENT | I(8)=ADDRESS | J(9+)=Products
    const orderId = String(rowData[ORDER_COL.INVOICE_NUMBER] || "").trim();       // Column B
    const clientName = String(rowData[ORDER_COL.CLIENT] || "").trim();            // Column G
    const clientComments = String(rowData[ORDER_COL.COMMENT] || "").trim();       // Column H
    const orderDate = (rowData[ORDER_COL.TIME_STAMP] instanceof Date) ? rowData[ORDER_COL.TIME_STAMP] : new Date();  // Column C
    const totalAmount = parseFloat(rowData[ORDER_COL.TOTAL]) || 0;               // Column F

    if (!orderId) {
        ss.toast("Could not find Invoice Number in the selected row.");
        return;
    }

    ss.toast(`Generating PDF for Order ${orderId}...`);

    // Parse product data from column J (index 9) onwards
    const items = [];
    for (let i = ORDER_COL.PRODUCTS_START; i < rowData.length; i++) {
        const cellValue = String(rowData[i] || "").trim();
        if (!cellValue) continue;

        // Parse format: [1|@SKU|$33.00|T]
        const match = cellValue.match(/\[(\d+)\|@?([^\|]+)\|\$?([\d.]+)\|([TF])\]/);
        if (match) {
            items.push({
                quantity: parseInt(match[1]) || 0,
                sku: match[2].trim(),
                price: parseFloat(match[3]) || 0,
                onSale: match[4] === 'T'
            });
        }
    }

    if (items.length === 0) {
        ss.toast("No products found in this order.");
        return;
    }

    // Get client address - prefer the address stored in the order row, fall back to CLIENT DATA sheet
    let clientAddress = String(rowData[ORDER_COL.ADDRESS] || "").trim();
    if (!clientAddress) {
        try {
            const clients = getClientData();
            const client = clients.find(c => {
                const cName = c['Company Name'] || c['Name'] || "";
                return String(cName).trim().toLowerCase() === clientName.toLowerCase();
            });
            if (client) {
                clientAddress = client['Address'] || client['Street Address'] || "";
            }
        } catch (e) {
            Logger.log("Could not fetch client address: " + e.message);
        }
    }

    // Fetch Sales Rep name from CFG_SALES_REP named range
    let salesRepFirstName = "";
    try {
        const salesRepRange = ss.getRangeByName("CFG_SALES_REP");
        if (salesRepRange) {
            const fullName = String(salesRepRange.getValue() || "").trim();
            // Extract first name only (everything before the first space)
            salesRepFirstName = fullName.split(/\s+/)[0] || fullName;
        }
    } catch (e) {
        Logger.log("Could not read CFG_SALES_REP: " + e.message);
    }

    // Build order data object
    const orderData = {
        id: orderId,
        clientName: clientName,
        clientAddress: clientAddress,
        clientComments: clientComments,
        date: orderDate,
        total: totalAmount,
        items: items,
        salesRep: salesRepFirstName
    };

    try {
        const pdfUrl = generateOrderPdf(orderData);

        const html = HtmlService.createHtmlOutput(`
            <div style="font-family:Arial,sans-serif;padding:20px;">
                <p style="color:#006c4c;font-size:18px;font-weight:bold;">✓ PDF Generated Successfully!</p>
                <p><a href="${pdfUrl}" target="_blank" style="color:#006c4c;font-weight:bold;">Click here to open PDF</a></p>
            </div>
        `).setWidth(350).setHeight(150);

        SpreadsheetApp.getUi().showModalDialog(html, 'Order PDF Created');

    } catch (e) {
        SpreadsheetApp.getUi().alert("Error generating PDF: " + e.message);
        Logger.log("PDF Error: " + e.toString());
    }
}

/**
 * Open the Orders folder in a new tab
 */
function openOrdersFolder() {
    const folder = getOrdersFolder();
    const url = folder.getUrl();
    const html = `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`;
    SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createHtmlOutput(html),
        "Opening Orders Folder..."
    );
}
