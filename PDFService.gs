/**
 * PDFService.gs
 * Handles PDF generation via Staging Sheet (ORDER_DATA)
 */

/**
 * Generate PDF for an order
 * Uses LockService to ensure single-threaded access to ORDER_DATA
 */
function createOrderPdf(orderId, client, items, date) {
    const lock = LockService.getScriptLock();
    try {
        // Wait up to 30 seconds for other PDF jobs to finish
        lock.waitLock(30000);
    } catch (e) {
        throw new Error("System is busy generating another PDF. Please try again in a few seconds.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
        // 0. Prepare Staging Data (ORDER_DATA)
        populateOrderDataStaging(items);
        SpreadsheetApp.flush(); // Ensure data is calculated

        // 1. Prepare Template
        const templateSheet = getSheet(SHEET_NAMES.ORDERS_EXPORT);
        const tempSheet = templateSheet.copyTo(ss);
        tempSheet.setName("Temp_" + orderId);

        const config = getAppConfig();

        // 2. Populate Headers (Named Ranges) on the TEMP sheet
        // These are typically specific to the invoice (Client, Date)
        // The Line Items are handled via formulas pointing to ORDER_DATA

        setNamedRangeValue(tempSheet, "EXP_CLT_NAME", client['Name'] || "");
        setNamedRangeValue(tempSheet, "EXP_ORD_ID", orderId);
        setNamedRangeValue(tempSheet, "EXP_DATE", Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy"));

        const contactKey = Object.keys(client).find(k => k.toLowerCase() === config.CFG_COL_CONTACT.toLowerCase()) || "Contact Name";
        const repKey = Object.keys(client).find(k => k.toLowerCase() === config.CFG_COL_REP.toLowerCase()) || "Sales Rep";
        setNamedRangeValue(tempSheet, "EXP_CONTACT", client[contactKey] || "");
        setNamedRangeValue(tempSheet, "EXP_REP", client[repKey] || "");

        SpreadsheetApp.flush();

        // 3. Export
        const folder = getOrCreateFolder(date);
        const fileName = `Order_${String(client['Name']).replace(/[^a-zA-Z0-9]/g, '_')}_${orderId}.pdf`;
        const pdfFile = exportSheetToPdf(tempSheet, folder, fileName);

        return pdfFile.getUrl();

    } catch (e) {
        Logger.log("PDF Creation failed: " + e.message);
        throw e;
    } finally {
        // Cleanup
        const tempSheet = ss.getSheetByName("Temp_" + orderId);
        if (tempSheet) ss.deleteSheet(tempSheet);

        lock.releaseLock();
    }
}

/**
 * Write Order Quantities to ORDER_DATA
 */
function populateOrderDataStaging(items) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.ORDER_DATA);

    // Ensure sheet exists and has products
    // We assume setupOrderDataSheet() is run periodically, but we can check row count
    // If empty, run setup
    if (!sheet || sheet.getLastRow() < 2) {
        sheet = setupOrderDataSheet();
    }

    // 1. Get current Mapping (SKU -> Row Index)
    // We assume SKU is in Col A (1)
    const lastRow = sheet.getLastRow();
    const skuVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const skuMap = {};
    skuVals.forEach((row, i) => {
        // key = sku, value = row index (0-based relative to data start, so +2 for sheet row)
        skuMap[String(row[0]).trim()] = i + 2;
    });

    // 2. Clear old Quantities (Cols F=6, G=7)
    // We clear the range to ensure clean slate
    sheet.getRange(2, 6, lastRow - 1, 2).clearContent();

    // 3. Prepare Writes
    // We can't batch easily unless we build the whole array, which is cleaner.
    // Let's build a sparse update or read-modify-write?
    // Batch writing the whole column is fastest.

    const qtyUnitArr = new Array(lastRow - 1).fill().map(u => [""]);
    const qtyCaseArr = new Array(lastRow - 1).fill().map(u => [""]);

    // Fill arrays
    items.forEach(item => {
        const sku = String(item.sku).trim();
        const rowIndex = skuMap[sku]; // Sheet Row Index

        if (rowIndex) {
            const arrayIndex = rowIndex - 2; // Array index
            const val = item.quantity;

            // Type check
            const unitType = (item.type || item.unit || 'unit').toLowerCase();

            if (unitType === 'case' || unitType === 'mc') {
                qtyCaseArr[arrayIndex][0] = val;
            } else {
                qtyUnitArr[arrayIndex][0] = val;
            }
        } else {
            Logger.log(`Warning: SKU ${sku} in order but not in ORDER_DATA catalog.`);
        }
    });

    // 4. Write
    sheet.getRange(2, 6, qtyUnitArr.length, 1).setValues(qtyUnitArr);
    sheet.getRange(2, 7, qtyCaseArr.length, 1).setValues(qtyCaseArr);
}


/**
 * Helper: Set value to a named range specific to a sheet
 */
function setNamedRangeValue(sheet, rangeName, value) {
    try {
        const range = sheet.getRange(rangeName);
        try { range.clearDataValidation(); } catch (valError) { }
        range.setValue(value);
    } catch (e) {
        Logger.log(`Named Range ${rangeName} not found: ${e.toString()}`);
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
    const year = date.getFullYear();
    const week = getWeekNumber(date);

    const rootName = "Order_System_Exports";
    let rootFolder = DriveApp.getFoldersByName(rootName).hasNext() ? DriveApp.getFoldersByName(rootName).next() : DriveApp.createFolder(rootName);

    let yearFolder = rootFolder.getFoldersByName(String(year)).hasNext() ? rootFolder.getFoldersByName(String(year)).next() : rootFolder.createFolder(String(year));

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
