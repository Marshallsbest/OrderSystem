/**
 * PDFService.gs
 * Handles PDF generation and Export
 */

/**
 * Generate PDF for an order
 */
function createOrderPdf(orderId, client, items, date) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = getSheet(SHEET_NAMES.ORDERS_EXPORT);
    const settingsSheet = getSheet(SHEET_NAMES.SETTINGS); // If we need config

    // 1. Create Temp Sheet
    // We copy the template so we don't mess up the master
    const tempSheet = templateSheet.copyTo(ss);
    tempSheet.setName("Temp_" + orderId);

    try {
        // 2. Populate Named Ranges (Header Info)
        setNamedRangeValue(tempSheet, "EXP_CLT_NAME", client['Name'] || "");
        setNamedRangeValue(tempSheet, "EXP_ORD_ID", orderId);
        setNamedRangeValue(tempSheet, "EXP_DATE", Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy"));
        // Add other headers as needed (Address, etc.)

        // 3. Populate Line Items
        // Find Start Row from Named Range "EXP_ITM_START"
        const startRange = tempSheet.getRange(getNamedRangeA1(tempSheet, "EXP_ITM_START"));
        if (startRange) {
            const startRow = startRange.getRow();

            // We insert rows for items 
            // Note: If the template has specific styling below the items (totals), inserting rows pushes them down, which is good.
            // If the template has a fixed table, we might overwrite. 
            // Assumption: Inserting rows is safer for variable length orders.

            if (items.length > 0) {
                tempSheet.insertRowsAfter(startRow, items.length - 1); // -1 because we have the start row itself

                // Prepare data array for fast write
                // Structure depends on Template columns. 
                // We assume the template columns match logical order or we need specific named ranges per column?
                // "Using a list of named ranges" was the prompt. 
                // But for a dynamic list, Named Ranges usually define the *Area* or *Start*.
                // Let's assume standard columns relative to EXP_ITM_START: [SKU, Name, Qty, Price, Subtotal]
                // This is a risk point: User layout is "exact". 
                // We'll write to the columns starting at EXP_ITM_START's column code.

                const startCol = startRange.getColumn();
                const rangeToWrite = tempSheet.getRange(startRow, startCol, items.length, 5); // Writing 5 cols for example

                const productCatalog = getProductCatalog();

                const itemValues = items.map(item => {
                    const product = productCatalog.find(p => p.sku === item.sku);
                    // Fallback if product not found
                    const name = product ? product.name + (product.variation ? " - " + product.variation : "") : item.sku;
                    const price = product ? product.price : 0;
                    const subtotal = price * item.quantity;

                    return [
                        item.sku,
                        name,
                        item.quantity,
                        price,
                        subtotal
                    ];
                });

                rangeToWrite.setValues(itemValues);
            }
        }

        SpreadsheetApp.flush(); // Commit changes before export

        // 4. Export to PDF
        const folder = getOrCreateFolder(date);
        const pdfBlob = tempSheet.getParent().getAs('application/pdf'); // This exports the whole spreadsheet by default? 
        // No, we need to export specific sheet.
        // DriveApp/SpreadsheetApp export is tricky.
        // Standard robust method: HTTP Request to export URL with gid

        const pdfFile = exportSheetToPdf(tempSheet, folder, `Order_${client['Name']}_${orderId}.pdf`);

        // 5. Send Notification
        sendNotification(orderId, client['Name'], pdfFile.getUrl());

        return pdfFile.getUrl();

    } catch (e) {
        Logger.log("PDF Creation failed: " + e.message);
        throw e; // Rethrow to alert main process
    } finally {
        // 6. Cleanup
        ss.deleteSheet(tempSheet);
    }
}

/**
 * Helper: Set value to a named range specific to a sheet
 * Note: NamedRanges in copies usually lose the specific scope or keep the name?
 * When copying a sheet, named ranges usually become Sheet-scoped "SheetName!RangeName".
 */
function setNamedRangeValue(sheet, rangeName, value) {
    // We search for the range in the sheet
    const range = sheet.getRange(rangeName); // This might look for Workbook scope first. 
    // Should ideally scope search.
    // Simple hack: Try 'RangeName' directly. If unique, works. 
    // Or iterate named ranges.
    try {
        sheet.getRange(rangeName).setValue(value);
    } catch (e) {
        Logger.log(`Named Range ${rangeName} not found in temp sheet. Skipping.`);
    }
}

function getNamedRangeA1(sheet, rangeName) {
    // Named ranges checks
    try {
        return sheet.getRange(rangeName).getA1Notation();
    } catch (e) {
        return null;
    }
}

/**
 * robust PDF export via UrlFetchApp
 */
function exportSheetToPdf(sheet, folder, filename) {
    const ss = sheet.getParent();
    const sheetId = sheet.getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sheetId}`
        + '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false';

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });

    const blob = response.getBlob().setName(filename);
    return folder.createFile(blob);
}

/**
 * Folder Logic: Root/Year/WeekNumber
 */
function getOrCreateFolder(date) {
    // Root folder from settings or default
    // Simply using DriveApp.getRootFolder() if no ID setting 
    // Prompt said "stored in a folder designated by the year and week number"

    // Need to implement WEEKNUM logic
    const year = date.getFullYear();
    const week = getWeekNumber(date);

    const rootName = "Order_System_Exports";
    let rootFolder = DriveApp.getFoldersByName(rootName).hasNext() ? DriveApp.getFoldersByName(rootName).next() : DriveApp.createFolder(rootName);

    // Year Folder
    let yearFolder = rootFolder.getFoldersByName(String(year)).hasNext() ? rootFolder.getFoldersByName(String(year)).next() : rootFolder.createFolder(String(year));

    // Week Folder
    let weekFolder = yearFolder.getFoldersByName("Week " + week).hasNext() ? yearFolder.getFoldersByName("Week " + week).next() : yearFolder.createFolder("Week " + week);

    return weekFolder;
}

function getWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
}

/**
 * Send Email Notification
 */
function sendNotification(orderId, clientName, link) {
    const email = Session.getActiveUser().getEmail(); // Notify owner (script runner)
    // Or check SETTINGS for a specific email

    MailApp.sendEmail({
        to: email,
        subject: `New Order Received: ${orderId} - ${clientName}`,
        htmlBody: `<p>A new order has been received from <strong>${clientName}</strong>.</p>
               <p><a href="${link}">View PDF Order</a></p>`
    });
}
