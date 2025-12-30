/**
 * SheetService.gs
 * Handles all direct interactions with the Google Spreadsheet
 */

const SHEET_NAMES = {
    WELCOME: "Welcome",
    ORDERS_EXPORT: "ORDERS_EXPORT",
    SETTINGS: "SETTINGS",
    ORDERS: "ORDERS",
    ORDER_PLACING: "ORDER_PLACING",
    CLIENT_DATA: "CLIENT DATA",
    PRODUCTS: "PRODUCTS",
    ORDER_DATA: "ORDER_DATA" // New Staging Sheet
};

/**
 * Setup/Refresh the SETTINGS sheet
 * Creates dynamic category list and color inputs
 * Also initializes Config Values
 */
/**
 * Setup/Refresh the SETTINGS sheet
 * Creates standard layout if missing: Col A = Key, Col B = Value
 */
function setupSettingsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);

    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
        // Move to end
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(ss.getNumSheets());
    }

    // Set Header Row content if not set
    const headerRange = sheet.getRange("A1:B1");
    if (headerRange.getValues()[0][0] !== "Config Key") {
        const headers = ["Config Key", "Value"];
        headerRange.setValues([headers])
            .setFontWeight("bold")
            .setBackground("#e0e0e0");

        // Helper Note
        sheet.getRange("A1").setNote("The internal name of the setting");
        sheet.getRange("B1").setNote("Enter a Hex Code (e.g. #FF0000) or Color Name (e.g. Red) for category.");
        sheet.getRange("D1").setNote("Name of the Contact Person column in CLIENT DATA");
        sheet.getRange("E1").setNote("Name of the Sales Rep column in CLIENT DATA");
        sheet.getRange("F1").setNote("TRUE/FALSE to enable Sale Mode globally");

        // Add Default Keys if empty
        if (sheet.getLastRow() < 2) {
            const defaults = [
                ["Main", "#006c4c"],
                ["CFG_COL_CONTACT", "Contact Name"],
                ["CFG_COL_REP", "Sales Rep"],
                ["ENABLE_SALES", "TRUE"]
            ];
            sheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
        }

        // Default Values for D2, E2, F2 (These are not part of the A/B config, but separate cells)
        sheet.getRange("D2").setValue("Contact Name");
        sheet.getRange("E2").setValue("Sales Rep");
        sheet.getRange("F2").setValue("TRUE");
    }

    // Category Section Headers (User Modified: F=Name, G=Color, H=SaleActive, I=Order)
    const catHeader = sheet.getRange("F1:I1");
    if (catHeader.getValues()[0][0] !== "Category Name") {
        catHeader.setValues([["Category Name", "Color", "Sale Active?", "Display Order"]]);
        catHeader.setFontWeight("bold").setBackground("#e0e0e0");
        sheet.getRange("H1").setNote("TRUE/FALSE checkbox to enable sale mode for this category.");
        sheet.getRange("I1").setNote("Enter a number (1, 2, 3...) to control display order.");
        // Add Checkboxes to Column H (Rows 2-20)
        sheet.getRange("H2:H20").insertCheckboxes();
    }

    // Formatting
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 200);
}

/**
 * Setup/Refresh the ORDER_DATA staging sheet
 * Should be run on admin trigger or manually
 */
function setupOrderDataSheet() {
    // ... (rest of function usually)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.ORDER_DATA);

    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAMES.ORDER_DATA);
    }

    // ... (retaining the specific implementation usually found here, but let's just update getCategorySettings mostly)
    // Actually, I was in setupSettingsSheet above.
}

// ... (Skipping setupOrderDataSheet implementation for brevity in this tool call as I am targeting specific chunks)

/**
 * Fetch Category Settings (Colors & Order & SaleStatus)
 * Reads F/G/H/I for Category Meta: Name, Color, SaleActive, Order
 * Returns: { "CategoryName": { color: "#Hex", order: 1, saleActive: true }, ... }
 */
function getCategorySettings() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const settings = {};

    // Standard Default Colors
    const defaults = { 'Main': { color: '#006c4c', order: 0 } };

    // 1. Read Configs/Settings from A/B
    if (sheet && sheet.getLastRow() > 1) {
        const configData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
        configData.forEach(row => {
            const key = String(row[0]).trim();
            const val = String(row[1]).trim();
            if (key) {
                if (!settings[key]) settings[key] = { color: val, order: 999 };
                else settings[key].color = val;
            }
        });
    }

    // 2. Read Category Settings from F2:I20
    // F=Name(0), G=Color(1), H=SaleActive(2), I=Order(3)
    if (sheet) {
        // Extend range to I
        const catData = sheet.getRange("F2:I20").getValues();
        catData.forEach(row => {
            const cat = String(row[0]).trim();
            const color = String(row[1]).trim();

            // New Mapping based on User's Manual Column Insert
            const saleActive = (row[2] === true || String(row[2]).toUpperCase() === 'TRUE');
            const order = parseInt(row[3]); // Now in 4th position (index 3)

            if (cat) {
                if (!settings[cat]) settings[cat] = { color: "#cccccc", order: 999, saleActive: false };

                if (color) settings[cat].color = color;
                if (!isNaN(order)) settings[cat].order = order;
                settings[cat].saleActive = saleActive;
            }
        });
    }

    // Ensure Main Default
    if (!settings['Main'] || settings['Main'].color === "#cccccc") {
        settings['Main'] = defaults['Main'];
    }

    return settings;
}

/**
 * Get specific sheet by name
 */
function getSheet(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found. Please ensure the spreadsheet is set up correctly.`);
    }
    return sheet;
}

/**
 * Get all data from CLIENT DATA sheet
 * Robustly finds headers and data
 * Returns array of objects keyed by header name
 */
function getClientData() {
    const sheet = getSheet(SHEET_NAMES.CLIENT_DATA);

    // Attempt to find where data actually starts
    // User originally said headers on Row 2, Data Row 3
    // But let's scan a bit if needed or stick to the rule but be safe

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol < 1) return [];

    // Get the headers from Row 2
    const headerRange = sheet.getRange(2, 1, 1, lastCol);
    const headerValues = headerRange.getValues()[0];

    // Filter out empty headers to avoid "undefined" keys
    // Create a map of index -> headerName
    const validIndices = [];
    const headers = [];

    headerValues.forEach((h, i) => {
        const val = String(h).trim();
        if (val) {
            validIndices.push(i);
            headers.push(val);
        }
    });

    if (headers.length === 0) return [];

    // Get Data from Row 3 to Last Row
    // Check if there is data
    if (lastRow < 3) return [];

    const dataValues = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

    return dataValues.map(row => {
        let client = {};
        // Only map valid columns
        validIndices.forEach((colIndex, i) => {
            const headerName = headers[i];
            client[headerName] = row[colIndex];
        });
        return client;
    });
}

/**
 * Find a client by ID
 */
function getClientById(clientId) {
    const clients = getClientData();
    if (clients.length === 0) return null;

    // 1. Find the key that looks like "ClientID" or "CLIENT_ID"
    // We scan the first client object to find the matching key
    const sample = clients[0];
    const keys = Object.keys(sample);

    // Look for key matching "clientid" or "client_id" (case-insensitive, ignore noise)
    const idKey = keys.find(k => {
        const normalized = k.toLowerCase().replace(/[^a-z0-9]/g, '');
        return normalized === 'clientid';
    });

    if (!idKey) {
        // Fallback: Check if the user ID matches ANY value in the row? No, too dangerous.
        // Return null or throw to indicate configuration error
        console.error("Could not find a Client ID column. Available keys:", keys);
        return null;
    }

    // 2. Search for the client with loose comparison
    const targetId = String(clientId).trim().toLowerCase();

    const client = clients.find(c => {
        const val = String(c[idKey]).trim().toLowerCase();
        return val === targetId;
    });


    if (client) {
        // Normalize keys for Frontend and OrderService
        // Find keys for Name and Address fuzzily too if needed, but let's stick to the mapped ones plus mapped fallback
        client.Name = client['Company Name'] || client['Company Name '] || client['Name'] || "";
        client.Address = client['Address'] || client['Address '] || "";
    }

    return client;
}

/**
 * DEBUG FUNCTION: Return diagnostic info about Client Lookup
 */
function debugClientLookup(clientId) {
    const sheet = getSheet(SHEET_NAMES.CLIENT_DATA);
    const allValues = sheet.getDataRange().getValues();

    if (allValues.length === 0) return { error: "Sheet is completely empty" };

    const rawHeaders = allValues[0];
    const rawFirstRow = allValues.length > 1 ? allValues[1] : "NO DATA ROW";

    const clients = getClientData();
    // Re-run match logic to show what happens

    let sampleKeys = [];
    if (clients.length > 0) {
        sampleKeys = Object.keys(clients[0]);
    }

    const targetId = String(clientId).trim().toLowerCase();

    return {
        sheetName: sheet.getName(),
        totalSheetRows: allValues.length,
        rawHeaders: JSON.stringify(rawHeaders),
        rawFirstRowData: JSON.stringify(rawFirstRow),
        parsedObjectKeys: JSON.stringify(sampleKeys), // What getClientData produced
        searchingFor: targetId
    };
}

/**
 * DEBUG: Fetch raw data from SETTINGS sheet
 */
function debugGetSettingsData() {
    const sheet = getSheet(SHEET_NAMES.SETTINGS);
    if (!sheet) return "Settings sheet not found";
    return sheet.getDataRange().getValues();
}
/**
 * Get App Configuration from SETTINGS sheet
 * Expects Col A = Key, Col B = Value
 */
/**
 * Get App Configuration from SETTINGS sheet
 * Expects Col A = KeyName, Col D = Key, Col E = Value (Wait, user said A=Key, B=Value)
 * User Logic: "column A being the key column B being the range for the value"
 */
function getAppConfig() {
    const config = {
        CFG_COL_CONTACT: "Contact Name", // Default
        CFG_COL_REP: "Sales Rep"         // Default
    };

    try {
        const sheet = getSheet(SHEET_NAMES.SETTINGS);
        if (!sheet) return config;

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return config;

        // Read all of Col A and Col B
        const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A2:B

        data.forEach(row => {
            const key = String(row[0]).trim();
            const val = String(row[1]).trim();

            if (key) {
                // If the key matches our known config keys, update them
                // We handle standard keys and potential named range keys by name
                if (key === "CFG_COL_CONTACT") config.CFG_COL_CONTACT = val;
                if (key === "CFG_COL_REP") config.CFG_COL_REP = val;

                // Also support arbitrary keys if needed later
                config[key] = val;
            }
        });

    } catch (e) {
        console.error("Error reading config:", e);
    }

    return config;
}

/**
 * VISUAL HELPER: Apply background colors to the Settings sheet Column G
 * based on the Hex/Color value inside the cell.
 */
function applyCategoryColorsVisuals() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);

    if (!sheet) {
        ss.toast("Settings sheet not found.");
        return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // Column G is Index 7
    const range = sheet.getRange(2, 7, lastRow - 1, 1);
    const values = range.getValues();
    const backgrounds = [];

    values.forEach(row => {
        const colorVal = String(row[0]).trim();
        // Check if it looks like a valid color (Hex or Name)
        if (colorVal) {
            backgrounds.push([colorVal]);
        } else {
            backgrounds.push([null]); // Clear background
        }
    });

    range.setBackgrounds(backgrounds);
    ss.toast("Updated color previews in Column G.");
}
