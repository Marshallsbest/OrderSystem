/**
 * PDFService.gs
 * Generates styled PDF invoices for orders
 * Version: v1.8.32
 */

// ============================================================================
// FOLDER MANAGEMENT
// ============================================================================

/**
 * Get or create the Orders folder for PDF storage
 * Checks SETTINGS for custom folder URL first, otherwise creates "Orders" folder
 */
/**
 * Helper to find or create a subfolder by name
 * @param {GoogleAppsScript.Drive.Folder} parentFolder
 * @param {string} folderName
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateSubfolder(parentFolder, folderName) {
    const folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
        return folders.next();
    } else {
        const newFolder = parentFolder.createFolder(folderName);
        Logger.log(`Created folder: ${folderName} in ${parentFolder.getName()}`);
        return newFolder;
    }
}

/**
 * Get or create the monthly folder for PDF storage
 * Structure: [Main Folder] > Order System > [Current Month Year]
 */
function getOrdersFolder() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const parents = ssFile.getParents();
    let parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

    // Check if parent is actually "Order System". If not, look for it/create it inside parent.
    let systemFolder;
    if (parentFolder.getName() === "Order System") {
        systemFolder = parentFolder;
    } else {
        systemFolder = getOrCreateSubfolder(parentFolder, "Order System");
    }

    // Get or Create Monthly folder (e.g., "February 2026")
    const now = new Date();
    const monthFolderName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM YYYY");
    const monthFolder = getOrCreateSubfolder(systemFolder, monthFolderName);

    return monthFolder;
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
/**
 * Build the HTML content for the invoice
 */
function buildInvoiceHtml(orderData, formattedDate) {
    const clientName = orderData.clientName || "Unknown Client";
    const clientAddress = formatAddress(orderData.clientAddress || "");
    const comments = orderData.clientComments || "";
    const items = orderData.items || [];

    // Fetch Sales Rep from settings - Prioritize CFG_SALES_REP
    const settingsRep = getSettingValue("CFG_SALES_REP") || getSettingValue("Sales Rep") || getSettingValue("SALES_REP");
    const salesRep = orderData.salesRep || settingsRep || "Admin";

    // Extract first name only if it exists (for a cleaner look)
    const displayRep = String(salesRep).split(/\s+/)[0] || salesRep;

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
    // Columns: Product | Qty | Price | Sale Price | Subtotal
    let itemRowsHtml = "";
    enrichedItems.forEach(item => {
        const productDisplay = item.name + (item.variation ? " " + item.variation : "");
        const finalPrice = item.onSale ? item.salePrice : item.price;
        const lineSubtotal = finalPrice * item.quantity;

        // Regular price
        const regularPriceDisplay = `$${item.price.toFixed(2)}`;

        // Sale price: shown only if on sale, blank otherwise
        const salePriceDisplay = item.onSale
            ? `$${item.salePrice.toFixed(2)}`
            : '-';

        itemRowsHtml += `
            <tr>
                <td style="padding:4px 8px; border:1px solid #eee; font-size:12px;">${productDisplay}</td>
                <td style="padding:4px 8px; border:1px solid #eee; text-align:center; font-size:12px;">${item.quantity}</td>
                <td style="padding:4px 8px; border:1px solid #eee; text-align:right; font-size:12px;">${regularPriceDisplay}</td>
                <td style="padding:4px 8px; border:1px solid #eee; text-align:right; font-size:12px; color: ${item.onSale ? '#e53935' : '#333'};">
                    ${salePriceDisplay}
                </td>
                <td style="padding:4px 8px; border:1px solid #eee; text-align:right; font-weight:600; font-size:12px;">$${lineSubtotal.toFixed(2)}</td>
            </tr>
        `;
    });

    // Build full HTML document
    const html = `
<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            font-family: 'Inter', 'Roboto', Arial, sans-serif;
            margin: 30px;
            color: #1a1c1e;
            background-color: #ffffff;
            line-height: 1.3;
        }
        .header-container {
            margin-bottom: 20px;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 8px;
            border: 1px solid #e1e3e5;
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
        }
        .info-label {
            font-size: 10px;
            font-weight: 700;
            color: #5e6066;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 2px;
        }
        .sales-rep-section {
            text-align: right;
        }
        .sales-rep-name {
            font-size: 16px;
            font-weight: 700;
            color: #006c4c;
        }
        .client-info {
            flex-grow: 1;
        }
        .client-name {
            font-size: 18px;
            font-weight: 800;
            color: #1a1c1e;
            margin-bottom: 2px;
        }
        .client-address {
            font-size: 12px;
            color: #44474e;
            max-width: 400px;
        }
        .order-meta {
            margin-top: 10px;
            font-size: 11px;
            color: #74777f;
            display: flex;
            justify-content: space-between;
            border-top: 1px solid #eee;
            padding-top: 8px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
            border: 1px solid #e1e3e5;
        }
        th {
            background: #f1f3f4;
            padding: 8px 10px;
            text-align: left;
            font-size: 11px;
            font-weight: 700;
            color: #44474e;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            border-bottom: 2px solid #e1e3e5;
        }
        th:nth-child(2) { text-align: center; width: 50px; }
        th:nth-child(3) { text-align: right; width: 80px; }
        th:nth-child(4) { text-align: right; width: 80px; }
        th:nth-child(5) { text-align: right; width: 100px; }
        
        .total-section {
            margin-top: 15px;
            padding: 12px 20px;
            background: #006c4c;
            color: white;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .total-label {
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .total-value {
            font-size: 20px;
            font-weight: 800;
        }
        .comments-box {
            margin-top: 20px;
            padding: 12px;
            background: #f8f9fa;
            border-left: 4px solid #006c4c;
            border-radius: 4px;
        }
        .comments-title {
            font-size: 11px;
            font-weight: 700;
            color: #006c4c;
            margin-bottom: 4px;
            text-transform: uppercase;
        }
        .comments-text {
            font-size: 12px;
            color: #44474e;
            font-style: italic;
        }
    </style>
</head>
<body>
    <div class="header-container">
        <div class="client-info">
            <div class="info-label">Client Company</div>
            <div class="client-name">${clientName}</div>
            <div class="client-address">${clientAddress}</div>
        </div>

        <div class="sales-rep-section">
            <div class="info-label">Sales Representative</div>
            <div class="sales-rep-name">${displayRep}</div>
        </div>
    </div>
    
    <div class="order-meta">
        <span>Order Date: <strong>${formattedDate}</strong></span>
        <span>Order ID: #${orderData.id || 'N/A'}</span>
    </div>
    
    <table>
        <thead>
            <tr>
                <th>Product</th>
                <th>Qty</th>
                <th>Price</th>
                <th>Sale Price</th>
                <th>Subtotal</th>
            </tr>
        </thead>
        <tbody>
            ${itemRowsHtml}
        </tbody>
    </table>

    <div class="total-section">
        <span class="total-label">Total Order Amount</span>
        <span class="total-value">$${orderTotal.toFixed(2)}</span>
    </div>

    ${comments ? `
    <div class="comments-box">
        <div class="comments-title">Special Instructions / Comments</div>
        <div class="comments-text">"${comments}"</div>
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
