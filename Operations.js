/**
 * Operations.gs
 * Core data management and administrative operations
 */

/**
 * Fetch App Configuration from Settings sheet
 */
function getAppConfig() {
    const config = { CFG_COL_CONTACT: "Contact Name", CFG_COL_REP: "Sales Rep" };
    try {
        const sheet = getSheet(SHEET_NAMES.SETTINGS);
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return config;
        const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        data.forEach(row => {
            const key = String(row[0]).trim();
            const val = String(row[1]).trim();
            if (key) config[key] = val;
        });

        // Directive v1.8.49: Fetch Admin Key from Named Range
        try {
            const adminRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("ADMIN_LOGIN");
            if (adminRange) {
                config.ADMIN_KEY = String(adminRange.getValue()).trim();
            } else {
                config.ADMIN_KEY = config.ADMIN_KEY || "ADMIN123"; // Fallback if range missing
            }
        } catch (e) { }

    } catch (e) {
        console.error("Error reading config:", e);
    }
    return config;
}

/**
 * Update Configuration Setting
 */
function updateConfigSetting(key, value) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (!sheet) {
        setupSettingsSheet();
        sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    let rowToUpdate = -1;
    for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === key) {
            rowToUpdate = i + 2;
            break;
        }
    }

    if (rowToUpdate !== -1) {
        sheet.getRange(rowToUpdate, 2).setValue(value);
    } else {
        sheet.appendRow([key, value]);
    }
    SpreadsheetApp.flush();
    return { success: true, key: key, value: value };
}

/**
 * Dashboard Action Router
 */
function onSelectionChange(e) {
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== SHEET_NAMES.DASHBOARD) return;

    if (range.getNumRows() === 1 && range.getColumn() === 2) {
        const row = range.getRow();
        if (row < 6 || row > 15) return;

        const actions = {
            6: "showOrderFormDialog",
            8: "showAddProductSidebar",
            11: "generateSelectedOrderPdf",
            12: "cleanupProductSheet",
            13: "styleProductHeaders",
            14: "refreshDailyOperationsDashboard"
        };

        const functionName = actions[row];
        if (functionName && typeof this[functionName] === 'function') {
            SpreadsheetApp.getActiveSpreadsheet().toast("Processing action: " + actionNameFromRow(row) + "...", "Order System", 3);
            sheet.getRange("A1").activate(); // Reset selection
            this[functionName]();
        }
    }
}

/**
 * Helper to get clean name for Toast
 */
function actionNameFromRow(row) {
    const names = {
        6: "Launch Web App",
        8: "Add Product",
        11: "Generate PDF",
        12: "Cleanup",
        13: "Refresh Visuals",
        14: "Update Daily Ops"
    };
    return names[row] || "Action";
}

/**
 * Save Client Information Update Request
 * Writes to CLIENT_INFO_UPDATES sheet for admin review
 */
function saveClientInfoUpdate(updateData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.CLIENT_INFO_UPDATES);

    // Create sheet if it doesn't exist
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAMES.CLIENT_INFO_UPDATES);
        sheet.appendRow([
            'Timestamp',
            'Original Client ID',
            'New Client ID',
            'New Client Name',
            'New Address',
            'Status'
        ]);
        sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    // Append the update request
    sheet.appendRow([
        new Date(),
        updateData.originalClientId || '',
        updateData.newClientId || '',
        updateData.newClientName || '',
        updateData.newAddress || '',
        'Pending Review'
    ]);

    return { success: true, message: 'Update request submitted for review.' };
}

/**
 * Client Data Operations
 */
function getClientById(clientId) {
    const clients = getClientData();
    if (clients.length === 0) return null;
    const targetId = String(clientId).trim().toLowerCase();
    const idKey = Object.keys(clients[0]).find(k => superNormalize(k) === 'clientid');
    if (!idKey) return null;

    const client = clients.find(c => String(c[idKey]).trim().toLowerCase() === targetId);
    if (client) {
        client.Name = client['Company Name'] || client['Company Name '] || client['Name'] || "";
        client.Address = client['Address'] || client['Address '] || "";

        // DEBUG: Log all client keys
        const clientKeys = Object.keys(client);
        console.log('[getClientById] Client keys:', JSON.stringify(clientKeys));

        // Read section permissions - try to find section columns dynamically
        client.allowedSections = [];

        clientKeys.forEach(key => {
            const upperKey = String(key).toUpperCase();
            if (upperKey.includes('SECTION')) {
                const val = client[key];
                console.log(`[getClientById] Section key "${key}" = ${val} (type: ${typeof val})`);

                if (val === true || String(val).toUpperCase() === 'TRUE') {
                    // Extract A, B, C, D from key name
                    if (upperKey.includes('_A') || upperKey.endsWith('A')) client.allowedSections.push('A');
                    else if (upperKey.includes('_B') || upperKey.endsWith('B')) client.allowedSections.push('B');
                    else if (upperKey.includes('_C') || upperKey.endsWith('C')) client.allowedSections.push('C');
                    else if (upperKey.includes('_D') || upperKey.endsWith('D')) client.allowedSections.push('D');
                }
            }
        });

        console.log('[getClientById] Allowed sections after check:', JSON.stringify(client.allowedSections));

        // If no sections specified, allow all (backward compatibility)
        if (client.allowedSections.length === 0) {
            console.log('[getClientById] No sections found, defaulting to all');
            client.allowedSections = ['A', 'B', 'C', 'D'];
        }
    }
    return client;
}

function getClientData() {
    const sheet = getSheet(SHEET_NAMES.CLIENT_DATA);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 3 || lastCol < 1) return [];

    // Read BOTH header rows - row 1 has SECTION columns, row 2 has other columns
    const headerRow1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerRow2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

    // Smart merge: Use row 1 ONLY for SECTION columns, row 2 for everything else
    const validIndices = [];
    const headers = [];
    for (let i = 0; i < lastCol; i++) {
        const h1 = String(headerRow1[i] || '').trim();
        const h2 = String(headerRow2[i] || '').trim();

        // Use row 1 header only if it contains "SECTION", otherwise use row 2
        let header;
        if (h1.toUpperCase().includes('SECTION')) {
            header = h1; // Use SECTION_A, SECTION_B, etc. from row 1
        } else {
            header = h2 || h1; // Use row 2 (CLIENT_ID, Company Name, etc.), fallback to row 1
        }

        if (header) {
            validIndices.push(i);
            headers.push(header);
        }
    }

    console.log('[getClientData] Merged headers:', JSON.stringify(headers.slice(0, 10)) + '...');

    if (headers.length === 0) return [];
    const dataValues = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

    return dataValues.map(row => {
        let client = {};
        validIndices.forEach((colIndex, i) => { client[headers[i]] = row[colIndex]; });
        return client;
    });
}

/**
 * Fetch Client Types from the CLIENT_TYPES named range
 * Returns an array of type strings for the dropdown
 */
function getClientTypes() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const range = ss.getRangeByName("CLIENT_TYPES");
        if (!range) return [];

        const values = range.getValues();
        const types = [];
        values.forEach(row => {
            const val = String(row[0] || "").trim();
            if (val) types.push(val);
        });
        return types;
    } catch (e) {
        console.error("[getClientTypes] Error:", e.message);
        return [];
    }
}

/**
 * Fetch Section Names from named ranges SECTION_A, SECTION_B, SECTION_C, SECTION_D
 * Returns an array of { key: 'A', name: 'Tobacco' } objects
 */
function getSectionNames() {
    const sections = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const keys = ['A', 'B', 'C', 'D'];

    keys.forEach(key => {
        try {
            const range = ss.getRangeByName('SECTION_' + key);
            if (range) {
                const name = String(range.getValue() || '').trim();
                sections.push({ key: key, name: name || ('Section ' + key) });
            } else {
                sections.push({ key: key, name: 'Section ' + key });
            }
        } catch (e) {
            sections.push({ key: key, name: 'Section ' + key });
        }
    });

    return sections;
}

/**
 * Add a new client to the CLIENT DATA sheet
 * @param {Object} clientData - { clientId, companyName, type, phone, manager, address }
 * @returns {Object} - { success, message }
 */
function addNewClient(clientData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.CLIENT_DATA);
    if (!sheet) throw new Error("CLIENT DATA sheet not found.");

    const lastCol = sheet.getLastColumn();
    const headerRow1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerRow2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

    // Build a column index map from the header row (row 2) for data fields
    // Uses superNormalize for resilient matching regardless of column order or casing
    const colMap = {};
    const colMapRaw = {};
    headerRow2.forEach((h, i) => {
        const key = String(h || "").trim();
        const norm = superNormalize(key);
        if (norm) {
            colMap[norm] = i;
            colMapRaw[key.toLowerCase()] = i;
        }
    });

    // Build section column map from row 1 (SECTION_A, SECTION_B, etc.)
    const sectionColMap = {};
    headerRow1.forEach((h, i) => {
        const key = String(h || "").trim().toUpperCase();
        if (key.includes('SECTION')) {
            sectionColMap[key] = i;
        }
    });

    console.log("[addNewClient] Data column map:", JSON.stringify(Object.keys(colMap)));
    console.log("[addNewClient] Section column map:", JSON.stringify(sectionColMap));

    // Validate required fields
    const clientId = String(clientData.clientId || "").trim();
    const companyName = String(clientData.companyName || "").trim();
    if (!clientId) return { success: false, message: "Client ID is required." };
    if (!companyName) return { success: false, message: "Company Name is required." };

    // Check for duplicate Client ID
    const existingClients = getClientData();
    const idKey = Object.keys(existingClients[0] || {}).find(k => superNormalize(k) === 'clientid');
    if (idKey) {
        const duplicate = existingClients.find(c => String(c[idKey]).trim().toLowerCase() === clientId.toLowerCase());
        if (duplicate) return { success: false, message: "Client ID '" + clientId + "' already exists." };
    }

    // Build the new row (all empty, then fill mapped columns)
    const newRow = new Array(lastCol).fill("");

    // Map data fields to columns using normalized + raw matching for resilience
    const setCol = (names, value) => {
        // Try exact raw lowercase match first, then superNormalize match
        for (const name of names) {
            if (colMapRaw[name] !== undefined) {
                newRow[colMapRaw[name]] = value;
                return true;
            }
        }
        for (const name of names) {
            const norm = superNormalize(name);
            if (colMap[norm] !== undefined) {
                newRow[colMap[norm]] = value;
                return true;
            }
        }
        return false;
    };

    setCol(['client_id', 'clientid', 'client id', 'id'], clientId);
    setCol(['company name', 'companyname', 'name', 'company'], companyName);
    setCol(['type', 'client type', 'clienttype', 'account type'], clientData.type || "");
    setCol(['phone', 'phone number', 'telephone', 'tel'], clientData.phone || "");
    setCol(['manager', 'contact', 'contact name', 'manager name'], clientData.manager || "");
    setCol(['address', 'street address', 'full address'], clientData.address || "");

    // Set section values (TRUE/FALSE) from checkbox selections
    const sections = clientData.sections || {};
    ['A', 'B', 'C', 'D'].forEach(key => {
        const colKey = 'SECTION_' + key;
        if (sectionColMap[colKey] !== undefined) {
            newRow[sectionColMap[colKey]] = sections[key] === true;
        }
    });

    // Append after the last row
    const lastRow = sheet.getLastRow();
    const newRowNum = lastRow + 1;
    sheet.getRange(newRowNum, 1, 1, lastCol).setValues([newRow]);

    // Apply checkbox data validation to the section columns in the new row
    const checkboxRule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .setAllowInvalid(false)
        .build();

    ['A', 'B', 'C', 'D'].forEach(key => {
        const colKey = 'SECTION_' + key;
        if (sectionColMap[colKey] !== undefined) {
            const colNum = sectionColMap[colKey] + 1; // 1-indexed
            sheet.getRange(newRowNum, colNum).setDataValidation(checkboxRule);
        }
    });

    SpreadsheetApp.flush();
    console.log("[addNewClient] Added client: " + clientId + " - " + companyName + " with sections: " + JSON.stringify(sections));

    return { success: true, message: "Client '" + companyName + "' added successfully!" };
}

/**
 * Load color definitions from the COLOUR_FORMAT_DEFINITIONS named range.
 * Expects a table with 3 columns: Color Name | Hex Value | Text Color
 * Returns an object keyed by lowercase color name.
 */
function loadColorDefinitions_() {
    const defs = {};
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const range = ss.getRangeByName('COLOUR_FORMAT_DEFINITIONS');
        if (!range) return defs;

        const data = range.getValues();
        for (let r = 0; r < data.length; r++) {
            const name = String(data[r][0] || '').trim().toLowerCase();
            const hex = String(data[r][1] || '').trim();
            const textColor = String(data[r][2] || '').trim();
            if (name && hex) {
                defs[name] = { hex: hex, textColor: textColor || '#ffffff' };
            }
        }
    } catch (e) {
        console.warn('[loadColorDefinitions_] Could not load COLOUR_FORMAT_DEFINITIONS:', e.message);
    }
    return defs;
}

/**
 * Resolve a color value — accepts hex (#FFA500) or a named color (Orange).
 * Returns { hex, textColor } if a named color is matched, or { hex, textColor: null } for raw hex.
 */
function resolveColor_(value, colorDefs) {
    if (!value) return { hex: value, textColor: null };
    const v = String(value).trim();
    if (v.startsWith('#')) return { hex: v, textColor: null };   // Already hex, no auto text color
    const lookup = colorDefs[v.toLowerCase()];
    if (lookup) return { hex: lookup.hex, textColor: lookup.textColor };  // Named color found
    return { hex: v, textColor: null };                          // Unknown — pass through as-is
}

/**
 * Fetch App Styles (Primary/Secondary colours from named ranges)
 * Reads: PRIMARY_COLOUR, PRIMARY_COLOUR_TEXT, SECONDARY_COLOUR, SECONDARY_COLOUR_TEXT
 * Accepts hex values (#FFA500) or named colors (Orange, Teal, Navy, etc.)
 * When a named color is used, automatically applies its defined text color unless overridden.
 */
function getAppStyles() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const colorDefs = loadColorDefinitions_();
    const styles = {
        primaryColor: '#FFA500',      // Default orange (matches current brand)
        primaryTextColor: '#ffffff',
        secondaryColor: '#625b71',    // MD3 default secondary
        secondaryTextColor: '#ffffff'
    };

    // Read raw values and keep range references for font color fallback
    const rawValues = {};
    const rangeRefs = {};
    const rangeNames = ['PRIMARY_COLOUR', 'PRIMARY_COLOUR_TEXT', 'SECONDARY_COLOUR', 'SECONDARY_COLOUR_TEXT'];
    rangeNames.forEach(rangeName => {
        try {
            const range = ss.getRangeByName(rangeName);
            if (range) {
                const val = String(range.getValue() || '').trim();
                if (val) {
                    rawValues[rangeName] = val;
                    rangeRefs[rangeName] = range;
                }
            }
        } catch (e) {
            console.warn('[getAppStyles] Named range ' + rangeName + ' not found, using default.');
        }
    });

    // Resolve PRIMARY_COLOUR (may be a name like "Orange" or hex "#FFA500")
    if (rawValues['PRIMARY_COLOUR']) {
        const resolved = resolveColor_(rawValues['PRIMARY_COLOUR'], colorDefs);
        styles.primaryColor = resolved.hex;
    }

    // PRIMARY_COLOUR_TEXT logic
    if (rawValues['PRIMARY_COLOUR_TEXT'] && rawValues['PRIMARY_COLOUR_TEXT'] !== rawValues['PRIMARY_COLOUR']) {
        styles.primaryTextColor = rawValues['PRIMARY_COLOUR_TEXT'];
    } else {
        // Fallback: If text color is missing OR is identical to background, read from cell font color
        try {
            if (rangeRefs['PRIMARY_COLOUR']) {
                styles.primaryTextColor = rangeRefs['PRIMARY_COLOUR'].getFontColor();
            }
        } catch (e) {
            console.warn('[getAppStyles] Error fetching font color for PRIMARY_COLOUR:', e.message);
        }
    }

    // Resolve SECONDARY_COLOUR
    if (rawValues['SECONDARY_COLOUR']) {
        const resolved = resolveColor_(rawValues['SECONDARY_COLOUR'], colorDefs);
        styles.secondaryColor = resolved.hex;
    }

    // SECONDARY_COLOUR_TEXT logic
    if (rawValues['SECONDARY_COLOUR_TEXT'] && rawValues['SECONDARY_COLOUR_TEXT'] !== rawValues['SECONDARY_COLOUR']) {
        styles.secondaryTextColor = rawValues['SECONDARY_COLOUR_TEXT'];
    } else {
        // Fallback: If text color is missing OR is identical to background, read from cell font color
        try {
            if (rangeRefs['SECONDARY_COLOUR']) {
                styles.secondaryTextColor = rangeRefs['SECONDARY_COLOUR'].getFontColor();
            }
        } catch (e) {
            console.warn('[getAppStyles] Error fetching font color for SECONDARY_COLOUR:', e.message);
        }
    }

    console.log('[getAppStyles] Loaded:', JSON.stringify(styles));
    return styles;
}

/**
 * Ensure text color has adequate contrast against background.
 * If contrast is insufficient, returns white or black based on background luminance.
 * @param {string} bgHex - Background color hex (e.g. '#0000FF')
 * @param {string} textHex - Text color hex (e.g. '#000000')
 * @returns {string} - Validated text color hex
 */
function ensureContrast_(bgHex, textHex) {
    if (!bgHex || !textHex) return textHex || '#ffffff';

    const toLum = (hex) => {
        const h = String(hex).replace('#', '');
        if (h.length !== 6) return -1;
        const r = parseInt(h.substr(0, 2), 16);
        const g = parseInt(h.substr(2, 2), 16);
        const b = parseInt(h.substr(4, 2), 16);
        return ((r * 299) + (g * 587) + (b * 114)) / 1000;
    };

    const bgLum = toLum(bgHex);
    const txtLum = toLum(textHex);
    if (bgLum < 0) return textHex; // Can't parse bg, leave as-is

    // Check if contrast is sufficient (difference > 100 on 0-255 scale)
    if (txtLum >= 0 && Math.abs(bgLum - txtLum) > 100) {
        return textHex; // Contrast is fine
    }

    // Insufficient contrast — pick white or black based on background
    return (bgLum >= 128) ? '#000000' : '#ffffff';
}

/**
 * Fetch Category Settings (Colors & Order & SaleStatus)
 */
function getCategorySettings() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const settings = {};
    if (!sheet) return settings;

    const data = sheet.getDataRange().getValues();
    const lastRow = data.length;
    if (lastRow < 1) return settings;

    let headerRowIdx = -1;
    let catColIdx = -1;
    let bestCatScore = -1;

    for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < data[r].length; c++) {
            const s = String(data[r][c]).toLowerCase().trim();
            let score = -1;
            if (s === "category name" || s === "cat name") score = 10;
            else if (s === "category" || s === "cat") score = 5;

            if (score > bestCatScore) {
                headerRowIdx = r;
                catColIdx = c;
                bestCatScore = score;
            }
        }
        if (bestCatScore === 10) break;
    }

    if (headerRowIdx === -1) return settings;

    const headers = data[headerRowIdx];
    let colorIdx = -1, saleIdx = -1, orderIdx = -1, textColIdx = -1;
    let orderCandidates = [];

    headers.forEach((h, i) => {
        const head = String(h).trim().toLowerCase();
        // IMPORTANT: Check 'text colour' BEFORE generic 'colour' to prevent false match
        if (head.includes('text colo') || head.includes('font colo') || head === 'text colour' || head === 'text color') textColIdx = i;
        else if (head.includes('color') || head.includes('colour') || head === 'hex') colorIdx = i;
        else if (head.includes('sale active') || head.includes('sale status') || head.includes('sale mode')) saleIdx = i;
        else if (head.includes('order') || head.includes('sort') || head.includes('display')) orderCandidates.push(i);
    });

    // Find SECTION column for client access filtering
    let sectionIdx = -1;
    headers.forEach((h, i) => {
        const head = String(h).trim().toLowerCase();
        if (head === 'section' || head === 'group' || head === 'access') sectionIdx = i;
    });

    if (orderCandidates.length > 0) {
        let maxScore = -999;
        orderCandidates.forEach(idx => {
            let score = 0;
            for (let r = headerRowIdx + 1; r < Math.min(headerRowIdx + 11, lastRow); r++) {
                const val = data[r][idx];
                if (typeof val === 'number' && !isNaN(val)) score += 15;
                else if (typeof val === 'boolean' || val === true || val === false) score -= 40;
                else if (!isNaN(parseInt(val))) score += 5;
            }
            if (headers[idx].toLowerCase().includes('order') || headers[idx].toLowerCase().includes('sort')) score += 50;
            if (headers[idx].toLowerCase().includes('display order')) score += 75;
            if (score > maxScore) { maxScore = score; orderIdx = idx; }
        });
    }

    const range = sheet.getRange(headerRowIdx + 1, 1, lastRow - headerRowIdx, sheet.getLastColumn());
    const dataSlice = range.getValues();
    const backgrounds = range.getBackgrounds();
    // PERFORMANCE (Going GAS): Removed getFontColors() — text colour now read from data column

    for (let r = 0; r < dataSlice.length; r++) {
        const rawCatName = String(dataSlice[r][catColIdx]).trim();
        // STOP if name is empty - prevents picking up stray text below the table
        if (!rawCatName) break;
        if (rawCatName.toLowerCase().includes("category")) continue;

        const catKey = superNormalize(rawCatName);
        if (!catKey) continue;

        let order = 999;
        if (orderIdx > -1) {
            const rawVal = dataSlice[r][orderIdx];
            if (typeof rawVal === 'number') order = rawVal;
            else if (rawVal && !isNaN(parseInt(rawVal))) {
                const parsed = parseInt(String(rawVal).replace(/[^0-9]/g, ''));
                if (!isNaN(parsed)) order = parsed;
            }
        }

        const catColor = colorIdx > -1 ? (String(dataSlice[r][colorIdx] || "").trim() || backgrounds[r][colorIdx]) : "#cccccc";
        // Read text colour from dedicated column, auto-contrast fallback
        let catText = "";
        if (textColIdx > -1) {
            catText = String(dataSlice[r][textColIdx] || "").trim();
        }
        if (!catText && catColor && catColor.startsWith('#')) {
            catText = getContrastYIQ(catColor);
        }
        const catSection = sectionIdx > -1 ? String(dataSlice[r][sectionIdx]).trim().toUpperCase() : "";

        settings[catKey] = {
            name: rawCatName, // Store original name
            color: catColor,
            textColor: catText,
            order: order,
            section: catSection, // A, B, C, D for client filtering
            saleActive: saleIdx > -1 ? (dataSlice[r][saleIdx] === true || String(dataSlice[r][saleIdx]).toUpperCase() === 'TRUE') : false
        };
    }

    return settings;
}

/**
 * Fetch Default Variations from Settings
 * Looks for specific range or keywords in SETTINGS sheet
 */
function getVariationDefaults() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const defaults = { var1: "", var2: "", var3: "", var4: "", values: [] };
    if (!sheet) return defaults;

    const data = sheet.getDataRange().getValues();

    // Look for "VARIATION DEFAULTS" section
    const startRow = data.findIndex(r => String(r[0]).toUpperCase().includes("VARIATION DEFAULTS"));
    if (startRow === -1) return defaults;

    // Assume header is at startRow, data follows
    // [VARIATION DEFAULTS]
    // [Var 1 Name] | [Var 2 Name] | [Var 3 Name] | [Var 4 Name]
    // [Value 1] | [Value 2] | [Value 3] | [Value 4]

    const headerRow = data[startRow + 1];
    if (headerRow) {
        defaults.var1 = String(headerRow[0] || "");
        defaults.var2 = String(headerRow[1] || "");
        defaults.var3 = String(headerRow[2] || "");
        defaults.var4 = String(headerRow[3] || "");
    }

    return defaults;
}

/**
 * Fetch Variation Groups from the "Variation Groups and Values" table in SETTINGS.
 * Table structure: Group Name | Variation Number | Group Data (Comma Separated List)
 * Returns: [{ groupName, variationNumber, values: [...] }, ...]
 */
/**
 * Fetch Variation Groups from the "VARIATION_GROUPS_AND_VALUES" named range.
 * Range contains: [Title, Header Row, Data Rows...]
 * Returns: [{ groupName, variationNumber, values: [...] }, ...]
 */
function getVariationGroups() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName('VARIATION_GROUPS_AND_VALUES');
    const groups = [];
    if (!range) {
        console.warn('[getVariationGroups] Named range VARIATION_GROUPS_AND_VALUES not found.');
        return groups;
    }

    const data = range.getValues();
    const startRow = range.getRow();
    const startCol = range.getColumn();

    // Skip the title (index 0) and header (index 1)
    for (let i = 2; i < data.length; i++) {
        const groupName = String(data[i][0] || '').trim();
        if (!groupName) continue; // Skip empty rows within the range

        const varNum = parseInt(String(data[i][1] || '1')) || 1;
        const rawValues = String(data[i][2] || '').trim();

        const values = rawValues
            .split(',')
            .map(v => v.trim())
            .filter(v => v.length > 0);

        groups.push({
            groupName: groupName,
            variationNumber: varNum,
            values: values,
            _absoluteRow: startRow + i,  // 1-indexed sheet row
            _absoluteCol: startCol       // 1-indexed sheet column
        });
    }

    console.log('[getVariationGroups] Found ' + groups.length + ' groups in named range.');
    return groups;
}

/**
 * Add a new value to an existing Variation Group in SETTINGS.
 */
function addValueToVariationGroup(groupName, newValue) {
    const groups = getVariationGroups();
    const group = groups.find(g => g.groupName.toLowerCase() === groupName.toLowerCase());

    if (!group) throw new Error('Group "' + groupName + '" not found in VARIATION_GROUPS_AND_VALUES.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName('VARIATION_GROUPS_AND_VALUES');
    if (!range) throw new Error('Range not found during update.');
    const sheet = range.getSheet();

    const existingValues = group.values;
    if (existingValues.includes(newValue)) return { success: true };

    const updated = existingValues.length > 0 ? existingValues.join(', ') + ', ' + newValue : newValue;
    sheet.getRange(group._absoluteRow, group._absoluteCol + 2).setValue(updated);

    return { success: true };
}

/**
 * Create a new Variation Group in the persistent table.
 */
function createNewVariationGroup(varNum, groupName, valuesString) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName('VARIATION_GROUPS_AND_VALUES');
    if (!range) throw new Error('VARIATION_GROUPS_AND_VALUES named range not found.');

    const sheet = range.getSheet();
    const data = range.getValues();
    const startRow = range.getRow();
    const startCol = range.getColumn();

    // Find the first empty row within or after the range
    let insertRow = -1;
    for (let i = 2; i < data.length; i++) {
        if (!String(data[i][0] || '').trim()) {
            insertRow = startRow + i;
            break;
        }
    }

    if (insertRow === -1) {
        // Append at the bottom of the named range's row sequence
        insertRow = startRow + data.length;
    }

    sheet.getRange(insertRow, startCol, 1, 3).setValues([[groupName, varNum, valuesString]]);
    return { success: true, groupName: groupName };
}

/**
 * ==========================================
 * Header Protection & Recovery System
 * ==========================================
 */

/** Sheets to protect - each entry defines how many header rows to back up */
const PROTECTED_SHEETS = {
    'CLIENT DATA': { headerRows: 2 },  // Row 1: SECTION_*, Row 2: field headers
    'PRODUCTS': { headerRows: 1 },
    'ORDERS': { headerRows: 1 },
    'SETTINGS': { headerRows: 1 }
};

/**
 * Backup all critical sheet headers to Script Properties.
 * Call once to set the "golden" snapshot,	or re-run to update after intentional changes.
 */
function backupSheetHeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const backup = {};
    const timestamp = new Date().toISOString();

    Object.keys(PROTECTED_SHEETS).forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            console.warn(`[backupSheetHeaders] Sheet "${sheetName}" not found, skipping.`);
            return;
        }
        const config = PROTECTED_SHEETS[sheetName];
        const lastCol = sheet.getLastColumn();
        if (lastCol < 1) return;

        const headerData = [];
        for (let r = 1; r <= config.headerRows; r++) {
            const row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
            headerData.push(row.map(v => String(v || "").trim()));
        }

        backup[sheetName] = {
            headers: headerData,
            colCount: lastCol,
            headerRows: config.headerRows,
            timestamp: timestamp
        };
    });

    props.setProperty('HEADER_BACKUP', JSON.stringify(backup));
    props.setProperty('HEADER_BACKUP_TIMESTAMP', timestamp);

    const count = Object.keys(backup).length;
    SpreadsheetApp.getActiveSpreadsheet().toast(
        `Backed up headers for ${count} sheets at ${timestamp}`,
        'Header Backup', 5
    );

    console.log('[backupSheetHeaders] Saved backup for:', Object.keys(backup).join(', '));
    return { success: true, message: `Backed up ${count} sheets.`, timestamp: timestamp };
}

/**
 * Compare current sheet headers against the backup.
 * Shows a detailed report as a dialog.
 */
function compareSheetHeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const backupJson = props.getProperty('HEADER_BACKUP');

    if (!backupJson) {
        SpreadsheetApp.getUi().alert(
            'No Backup Found',
            'No header backup exists yet. Run "Backup Sheet Headers" first to create a golden snapshot.',
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        return { success: false, message: 'No backup found.' };
    }

    const backup = JSON.parse(backupJson);
    const backupTimestamp = props.getProperty('HEADER_BACKUP_TIMESTAMP') || 'Unknown';
    let report = `HEADER COMPARISON REPORT\nBackup from: ${backupTimestamp}\n${'='.repeat(50)}\n\n`;
    let hasChanges = false;

    Object.keys(PROTECTED_SHEETS).forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        const savedData = backup[sheetName];

        if (!sheet && savedData) {
            report += `⚠️ SHEET "${sheetName}": MISSING (was backed up)\n\n`;
            hasChanges = true;
            return;
        }
        if (!savedData) {
            report += `ℹ️ SHEET "${sheetName}": No backup exists for this sheet\n\n`;
            return;
        }

        const config = PROTECTED_SHEETS[sheetName];
        const lastCol = sheet.getLastColumn();
        const maxCol = Math.max(lastCol, savedData.colCount);

        let sheetReport = '';
        let sheetHasChanges = false;

        for (let r = 0; r < config.headerRows; r++) {
            const currentRow = lastCol > 0
                ? sheet.getRange(r + 1, 1, 1, maxCol).getValues()[0].map(v => String(v || "").trim())
                : [];
            const savedRow = savedData.headers[r] || [];

            // Pad shorter arrays
            while (currentRow.length < maxCol) currentRow.push('');
            while (savedRow.length < maxCol) savedRow.push('');

            for (let c = 0; c < maxCol; c++) {
                const saved = savedRow[c] || '';
                const current = currentRow[c] || '';

                if (saved !== current) {
                    sheetHasChanges = true;
                    const colLetter = columnToLetter(c + 1);
                    if (saved && !current) {
                        sheetReport += `  🔴 Row ${r + 1}, Col ${colLetter}: REMOVED "${saved}"\n`;
                    } else if (!saved && current) {
                        sheetReport += `  🟢 Row ${r + 1}, Col ${colLetter}: ADDED "${current}"\n`;
                    } else {
                        sheetReport += `  🟡 Row ${r + 1}, Col ${colLetter}: CHANGED "${saved}" → "${current}"\n`;
                    }
                }
            }

            // Check if column was moved (exists in both but different position)
            savedRow.forEach((savedHeader, savedIdx) => {
                if (!savedHeader) return;
                const currentIdx = currentRow.indexOf(savedHeader);
                if (currentIdx !== -1 && currentIdx !== savedIdx && savedHeader === currentRow[currentIdx]) {
                    sheetReport += `  ↔️ Row ${r + 1}: "${savedHeader}" moved from Col ${columnToLetter(savedIdx + 1)} → Col ${columnToLetter(currentIdx + 1)}\n`;
                    sheetHasChanges = true;
                }
            });
        }

        if (sheetHasChanges) {
            report += `⚠️ SHEET "${sheetName}": Changes detected\n${sheetReport}\n`;
            hasChanges = true;
        } else {
            report += `✅ SHEET "${sheetName}": No changes\n\n`;
        }
    });

    if (!hasChanges) {
        report += '\n🎉 All headers match the backup. No changes detected.';
    } else {
        report += '\n⚠️ Some headers have changed. Use "Reset Sheet Headers" to restore from backup.';
    }

    // Show report as a scrollable dialog
    const htmlReport = HtmlService.createHtmlOutput(
        `<pre style="font-family: Consolas, 'Courier New', monospace; font-size: 13px; white-space: pre-wrap; padding: 16px;">${report}</pre>`
    ).setWidth(600).setHeight(500);

    SpreadsheetApp.getUi().showModalDialog(htmlReport, 'Header Comparison Report');

    console.log('[compareSheetHeaders] Report generated. Has changes:', hasChanges);
    return { success: true, hasChanges: hasChanges, report: report };
}

/**
 * Reset sheet headers back to the backed-up golden snapshot.
 * Prompts for confirmation before making changes.
 */
function resetSheetHeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const props = PropertiesService.getScriptProperties();
    const backupJson = props.getProperty('HEADER_BACKUP');

    if (!backupJson) {
        ui.alert(
            'No Backup Found',
            'No header backup exists yet. Run "Backup Sheet Headers" first.',
            ui.ButtonSet.OK
        );
        return { success: false, message: 'No backup found.' };
    }

    // Confirm with the user
    const confirm = ui.alert(
        'Reset Sheet Headers',
        'This will restore all header rows to their backed-up state.\n\n' +
        'Only headers will be changed — your data rows will NOT be affected.\n\n' +
        'Do you want to proceed?',
        ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) {
        return { success: false, message: 'Cancelled by user.' };
    }

    const backup = JSON.parse(backupJson);
    let restoredCount = 0;

    Object.keys(backup).forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            console.warn(`[resetSheetHeaders] Sheet "${sheetName}" not found, skipping.`);
            return;
        }

        const savedData = backup[sheetName];
        const colCount = savedData.colCount;

        // Ensure enough columns exist
        const currentCols = sheet.getMaxColumns();
        if (currentCols < colCount) {
            sheet.insertColumnsAfter(currentCols, colCount - currentCols);
        }

        // Write header rows back
        for (let r = 0; r < savedData.headerRows; r++) {
            const headerValues = savedData.headers[r];
            sheet.getRange(r + 1, 1, 1, colCount).setValues([headerValues]);
        }

        restoredCount++;
        console.log(`[resetSheetHeaders] Restored headers for "${sheetName}" (${colCount} cols)`);
    });

    SpreadsheetApp.flush();
    ss.toast(`Restored headers for ${restoredCount} sheets from backup.`, 'Headers Reset', 5);

    return { success: true, message: `Restored ${restoredCount} sheets.` };
}

/**
 * ==========================================
 * Initialization & Deployment System
 * ==========================================
 */

/**
 * Runs once on first open of a new spreadsheet copy.
 * Automatically backs up headers and stamps version info.
 * Called from onOpen() — uses a flag in Script Properties to run only once.
 */
function initializeOnFirstOpen_() {
    try {
        const props = PropertiesService.getScriptProperties();
        const initialized = props.getProperty('SYSTEM_INITIALIZED');

        if (!initialized) {
            // First-time initialization
            console.log('[initializeOnFirstOpen] Running first-time setup...');

            // 1. Auto-backup headers as the golden snapshot
            backupSheetHeaders();

            // 2. Stamp version and spreadsheet info
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            props.setProperty('SYSTEM_INITIALIZED', 'true');
            props.setProperty('INSTALLED_VERSION', CURRENT_VERSION);
            props.setProperty('SPREADSHEET_ID', ss.getId());
            props.setProperty('INITIALIZED_AT', new Date().toISOString());

            ss.toast('System initialized! Headers backed up as golden snapshot.', 'First-Time Setup', 5);
            console.log('[initializeOnFirstOpen] Initialization complete.');
        }
    } catch (e) {
        // Silently fail — onOpen should never break the menu
        console.error('[initializeOnFirstOpen] Error:', e.message);
    }
}

/**
 * Register the current spreadsheet as the MASTER copy.
 * The master's ID is stored so copies can reference it for version checks.
 */
function registerAsMaster() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const scriptId = ScriptApp.getScriptId();

    props.setProperty('IS_MASTER', 'true');
    props.setProperty('MASTER_ID', ss.getId());
    props.setProperty('MASTER_SCRIPT_ID', scriptId);
    props.setProperty('MASTER_VERSION', CURRENT_VERSION);
    props.setProperty('MASTER_NAME', ss.getName());

    // Initialize the copy registry if it doesn't exist
    if (!props.getProperty('COPY_REGISTRY')) {
        props.setProperty('COPY_REGISTRY', JSON.stringify([]));
    }

    ss.toast(
        'Registered as MASTER.\n' +
        'Version: ' + CURRENT_VERSION + '\n' +
        'Script ID: ' + scriptId,
        'Master Registered', 5
    );

    return { success: true, masterId: ss.getId(), scriptId: scriptId, version: CURRENT_VERSION };
}

/**
 * Create a clean copy of this spreadsheet for a colleague.
 * Preserves: headers, settings, product catalog, formatting, named ranges
 * Clears: client data, orders, daily operations, client info updates
 * Removes: linked sheet URLs from settings
 */
function createCleanCopy() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();

    // Prompt for colleague name
    const response = ui.prompt(
        'Create New Copy',
        'Enter the colleague\'s name or company for the copy title:',
        ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;
    const collegeName = response.getResponseText().trim();
    if (!collegeName) {
        ui.alert('Please provide a name for the copy.');
        return;
    }

    ss.toast('Creating clean copy... This may take a moment.', 'Please Wait', 10);

    try {
        // 1. Make the copy via Drive
        const copyName = ss.getName() + ' — ' + collegeName;
        const file = DriveApp.getFileById(ss.getId());
        const copyFile = file.makeCopy(copyName);
        const copyId = copyFile.getId();
        const copySs = SpreadsheetApp.openById(copyId);

        // 2. Clear admin/client data (keep headers)
        const sheetsToClear = [
            { name: SHEET_NAMES.CLIENT_DATA, headerRows: 2 },
            { name: SHEET_NAMES.ORDERS, headerRows: 1 },
            { name: SHEET_NAMES.DAILY_OPERATIONS, headerRows: 1 },
            { name: SHEET_NAMES.CLIENT_INFO_UPDATES, headerRows: 1 }
        ];

        sheetsToClear.forEach(cfg => {
            const sheet = copySs.getSheetByName(cfg.name);
            if (sheet) {
                const lastRow = sheet.getLastRow();
                if (lastRow > cfg.headerRows) {
                    sheet.deleteRows(cfg.headerRows + 1, lastRow - cfg.headerRows);
                }
                console.log(`[createCleanCopy] Cleared data from "${cfg.name}"`);
            }
        });

        // 3. Clear linked URLs from SETTINGS (keep keys, remove URL values)
        const settingsSheet = copySs.getSheetByName(SHEET_NAMES.SETTINGS);
        if (settingsSheet) {
            const settingsData = settingsSheet.getDataRange().getValues();
            for (let r = 1; r < settingsData.length; r++) {
                const key = String(settingsData[r][0] || '').toLowerCase();
                // Clear URL-type settings
                if (key.includes('url') || key.includes('folder') || key.includes('link') ||
                    key.includes('drive') || key.includes('export')) {
                    settingsSheet.getRange(r + 1, 2).setValue('');
                }
            }
            console.log('[createCleanCopy] Cleared linked URLs from SETTINGS');
        }

        // 4. Stamp the copy with version and master info
        const copyProps = PropertiesService.getScriptProperties();
        // Note: Script Properties are per-script, not per-spreadsheet.
        // For copies, we'll store meta info in a hidden sheet instead.
        let metaSheet = copySs.getSheetByName('_SYSTEM_META');
        if (!metaSheet) {
            metaSheet = copySs.insertSheet('_SYSTEM_META');
            metaSheet.hideSheet();
        }
        metaSheet.clear();
        metaSheet.getRange(1, 1, 6, 2).setValues([
            ['MASTER_ID', ss.getId()],
            ['MASTER_SCRIPT_ID', ScriptApp.getScriptId()],
            ['CREATED_FROM_VERSION', CURRENT_VERSION],
            ['CREATED_AT', new Date().toISOString()],
            ['COPY_OWNER', collegeName],
            ['MASTER_NAME', ss.getName()]
        ]);

        // 5. Register the copy in the master's registry
        let registry = [];
        try {
            registry = JSON.parse(props.getProperty('COPY_REGISTRY') || '[]');
        } catch (e) { registry = []; }

        registry.push({
            id: copyId,
            name: copyName,
            owner: collegeName,
            createdAt: new Date().toISOString(),
            createdFromVersion: CURRENT_VERSION
        });

        props.setProperty('COPY_REGISTRY', JSON.stringify(registry));

        // 6. Share the copy URL
        const copyUrl = copySs.getUrl();

        const resultHtml = HtmlService.createHtmlOutput(
            '<div style="font-family: Roboto, sans-serif; padding: 16px;">' +
            '<h3 style="color: #006c4c;">✅ Copy Created Successfully!</h3>' +
            '<p><strong>Name:</strong> ' + copyName + '</p>' +
            '<p><strong>Version:</strong> ' + CURRENT_VERSION + '</p>' +
            '<p><strong>Registered in master:</strong> Yes</p>' +
            '<p style="margin-top: 16px;"><a href="' + copyUrl + '" target="_blank" ' +
            'style="background: #006c4c; color: white; padding: 10px 20px; text-decoration: none; border-radius: 8px;">' +
            '🔗 Open Copy</a></p>' +
            '<p style="font-size: 12px; color: #999; margin-top: 16px;">' +
            'Copy ID: ' + copyId + '</p>' +
            '</div>'
        ).setWidth(450).setHeight(300);

        ui.showModalDialog(resultHtml, 'New Copy Created');

        console.log('[createCleanCopy] Created copy:', copyId, 'for', collegeName);
        return { success: true, copyId: copyId, copyUrl: copyUrl };

    } catch (e) {
        console.error('[createCleanCopy] Error:', e.message);
        ui.alert('Error creating copy: ' + e.message);
        return { success: false, message: e.message };
    }
}

/**
 * Check if this spreadsheet's code is up to date with the master.
 * Works from both master and copies.
 */
function checkForUpdates() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const props = PropertiesService.getScriptProperties();

    const isMaster = props.getProperty('IS_MASTER') === 'true';

    if (isMaster) {
        // Master view: Show registry of all copies and their versions
        const registry = JSON.parse(props.getProperty('COPY_REGISTRY') || '[]');
        const masterVersion = CURRENT_VERSION;

        let report = '<div style="font-family: Roboto, sans-serif; padding: 16px;">';
        report += '<h3 style="color: #006c4c;">📡 Master Version: ' + masterVersion + '</h3>';

        if (registry.length === 0) {
            report += '<p>No copies registered yet. Use "Create New Copy for Colleague" to make one.</p>';
        } else {
            report += '<table style="width:100%; border-collapse:collapse; font-size:13px;">';
            report += '<tr style="background:#f5f5f5;"><th style="padding:8px; text-align:left;">Copy</th>' +
                '<th style="padding:8px;">Owner</th>' +
                '<th style="padding:8px;">Created Version</th>' +
                '<th style="padding:8px;">Status</th></tr>';

            registry.forEach(copy => {
                const isOutdated = copy.createdFromVersion !== masterVersion;
                const status = isOutdated
                    ? '<span style="color:#B00020;">⚠️ Outdated</span>'
                    : '<span style="color:#006c4c;">✅ Current</span>';
                report += '<tr style="border-bottom:1px solid #eee;">' +
                    '<td style="padding:8px;">' + (copy.name || copy.id) + '</td>' +
                    '<td style="padding:8px; text-align:center;">' + (copy.owner || '-') + '</td>' +
                    '<td style="padding:8px; text-align:center;">' + copy.createdFromVersion + '</td>' +
                    '<td style="padding:8px; text-align:center;">' + status + '</td></tr>';
            });
            report += '</table>';
            report += '<p style="font-size:12px; color:#666; margin-top:16px;">' +
                'To update copies, push the latest code with <code>clasp push</code> to each copy\'s script project.</p>';
        }
        report += '</div>';

        const html = HtmlService.createHtmlOutput(report).setWidth(600).setHeight(400);
        ui.showModalDialog(html, 'Deployment Dashboard');

    } else {
        // Copy view: Check against meta sheet for master info
        let metaSheet = ss.getSheetByName('_SYSTEM_META');
        let masterInfo = {};

        if (metaSheet) {
            const metaData = metaSheet.getDataRange().getValues();
            metaData.forEach(row => {
                masterInfo[String(row[0]).trim()] = String(row[1]).trim();
            });
        }

        const installedVersion = masterInfo['CREATED_FROM_VERSION'] || props.getProperty('INSTALLED_VERSION') || 'Unknown';
        const masterName = masterInfo['MASTER_NAME'] || 'Unknown';
        const currentCodeVersion = CURRENT_VERSION;
        const isUpToDate = currentCodeVersion === installedVersion || currentCodeVersion > installedVersion;

        let report = '<div style="font-family: Roboto, sans-serif; padding: 16px;">';
        report += '<h3>Version Status</h3>';
        report += '<p><strong>Running Code Version:</strong> ' + currentCodeVersion + '</p>';
        report += '<p><strong>Created from Master Version:</strong> ' + installedVersion + '</p>';
        report += '<p><strong>Master:</strong> ' + masterName + '</p>';

        if (isUpToDate) {
            report += '<p style="color:#006c4c; font-weight:bold; margin-top:16px;">✅ This copy is up to date.</p>';
        } else {
            report += '<p style="color:#B00020; font-weight:bold; margin-top:16px;">⚠️ This copy may be outdated.</p>';
            report += '<p style="font-size:13px;">Contact the master spreadsheet owner to push the latest code update.</p>';
        }
        report += '</div>';

        const html = HtmlService.createHtmlOutput(report).setWidth(450).setHeight(300);
        ui.showModalDialog(html, 'Update Check');
    }
}

/**
 * Utility: Update the master version stamp.
 * Call this after pushing a new version to update the registry.
 */
function stampMasterVersion() {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('MASTER_VERSION', CURRENT_VERSION);

    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Master version stamped: ' + CURRENT_VERSION,
        'Version Update', 3
    );
}

/**
 * Pull code updates from the master script.
 * Works on COPY spreadsheets — reads master script ID from _SYSTEM_META,
 * fetches the master's script files via the Apps Script API,
 * then overwrites this script's files with the master's code.
 *
 * PREREQUISITE: Apps Script API must be enabled in the GCP project.
 * Go to: https://script.google.com/home/usersettings and enable it.
 */
function pullUpdatesFromMaster() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // 1. Get master script ID from _SYSTEM_META sheet
    const metaSheet = ss.getSheetByName('_SYSTEM_META');
    if (!metaSheet) {
        ui.alert(
            'Not a Copy',
            'This spreadsheet does not have a _SYSTEM_META sheet.\n' +
            'This function only works on copies created via "Create New Copy for Colleague".',
            ui.ButtonSet.OK
        );
        return;
    }

    const metaData = metaSheet.getDataRange().getValues();
    const meta = {};
    metaData.forEach(row => {
        meta[String(row[0]).trim()] = String(row[1]).trim();
    });

    const masterScriptId = meta['MASTER_SCRIPT_ID'];
    if (!masterScriptId) {
        ui.alert(
            'Missing Master Info',
            'No master script ID found in _SYSTEM_META.\n' +
            'This copy may have been created before the update feature was added.',
            ui.ButtonSet.OK
        );
        return;
    }

    // 2. Confirm with user
    const confirm = ui.alert(
        'Pull Updates from Master',
        'This will replace ALL code in this spreadsheet with the latest code from the master.\n\n' +
        'Master: ' + (meta['MASTER_NAME'] || 'Unknown') + '\n' +
        'Current version: ' + CURRENT_VERSION + '\n\n' +
        'Your spreadsheet DATA (clients, orders, settings) will NOT be affected.\n' +
        'Only the script code will be updated.\n\n' +
        'Proceed?',
        ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    ss.toast('Fetching latest code from master...', 'Updating', 15);

    try {
        const token = ScriptApp.getOAuthToken();
        const myScriptId = ScriptApp.getScriptId();

        // 3. Fetch master's script content
        const masterUrl = 'https://script.googleapis.com/v1/projects/' + masterScriptId + '/content';
        const masterResponse = UrlFetchApp.fetch(masterUrl, {
            method: 'GET',
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true
        });

        if (masterResponse.getResponseCode() !== 200) {
            const errBody = masterResponse.getContentText();
            console.error('[pullUpdatesFromMaster] Failed to fetch master:', errBody);

            // Check for common errors
            if (masterResponse.getResponseCode() === 403) {
                ui.alert(
                    'API Access Denied',
                    'The Apps Script API needs to be enabled.\n\n' +
                    'Steps:\n' +
                    '1. Go to https://script.google.com/home/usersettings\n' +
                    '2. Turn ON "Google Apps Script API"\n' +
                    '3. Try again.\n\n' +
                    'Also ensure you have view access to the master spreadsheet.',
                    ui.ButtonSet.OK
                );
            } else {
                ui.alert('Error fetching master code: ' + errBody);
            }
            return;
        }

        const masterContent = JSON.parse(masterResponse.getContentText());
        const masterFiles = masterContent.files;

        if (!masterFiles || masterFiles.length === 0) {
            ui.alert('Error: No files found in master script.');
            return;
        }

        console.log('[pullUpdatesFromMaster] Fetched ' + masterFiles.length + ' files from master');

        // 4. Push master's files to this script
        const updateUrl = 'https://script.googleapis.com/v1/projects/' + myScriptId + '/content';
        const updateResponse = UrlFetchApp.fetch(updateUrl, {
            method: 'PUT',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + token },
            payload: JSON.stringify({ files: masterFiles }),
            muteHttpExceptions: true
        });

        if (updateResponse.getResponseCode() !== 200) {
            const errBody = updateResponse.getContentText();
            console.error('[pullUpdatesFromMaster] Failed to update:', errBody);
            ui.alert('Error updating code: ' + errBody);
            return;
        }

        // 5. Update the meta sheet with new version info
        // Find which version we just pulled (parse from the master's Config.js)
        let pulledVersion = 'Unknown';
        masterFiles.forEach(f => {
            if (f.name === 'Config') {
                const versionMatch = f.source.match(/CURRENT_VERSION\s*=\s*["']([^"']+)["']/);
                if (versionMatch) pulledVersion = versionMatch[1];
            }
        });

        // Update meta sheet
        const metaValues = metaSheet.getDataRange().getValues();
        for (let r = 0; r < metaValues.length; r++) {
            if (String(metaValues[r][0]).trim() === 'CREATED_FROM_VERSION') {
                metaSheet.getRange(r + 1, 2).setValue(pulledVersion);
            }
        }
        // Add update timestamp
        const lastRow = metaSheet.getLastRow();
        metaSheet.getRange(lastRow + 1, 1, 1, 2).setValues([['LAST_UPDATED', new Date().toISOString()]]);

        console.log('[pullUpdatesFromMaster] Successfully updated to version:', pulledVersion);

        // 6. Show success and instruct to reload
        const successHtml = HtmlService.createHtmlOutput(
            '<div style="font-family: Roboto, sans-serif; padding: 20px; text-align: center;">' +
            '<h2 style="color: #006c4c;">✅ Code Updated Successfully!</h2>' +
            '<p style="font-size: 16px;">Updated to version: <strong>' + pulledVersion + '</strong></p>' +
            '<p style="font-size: 14px; color: #666;">Files synced: ' + masterFiles.length + '</p>' +
            '<hr style="margin: 20px 0;">' +
            '<p style="color: #B00020; font-weight: bold; font-size: 15px;">⚠️ You must RELOAD this spreadsheet for changes to take effect.</p>' +
            '<p style="font-size: 13px;">Close this tab and reopen the spreadsheet, or press Ctrl+Shift+R.</p>' +
            '</div>'
        ).setWidth(450).setHeight(280);

        ui.showModalDialog(successHtml, 'Update Complete');

    } catch (e) {
        console.error('[pullUpdatesFromMaster] Error:', e.message, e.stack);
        ui.alert('Update failed: ' + e.message);
    }
}

/**
 * ==========================================
 * Factory Reset System
 * ==========================================
 * Reads the "SHEET HEADER BACKUPS" sheet to understand what needs to be preserved.
 * 
 * Expected format — each row in "SHEET HEADER BACKUPS":
 *   Column A : Type → "Sheet" | "Range" | "Table"
 *   Column B : Name → Sheet name, Named range name, or "SheetName::TableHeader" for tables
 *   Column C+: Headers/values to preserve (for Sheet/Table types, these are the column headers)
 *
 * For "Sheet": Clears all data rows, restores header row(s) from the backup.
 * For "Range": Clears the named range value to blank.
 * For "Table": Finds the table within its sheet, clears data but keeps headers.
 */

const HEADER_BACKUP_SHEET = 'SHEET HEADER BACKUPS';

/**
 * Read the header backup definitions from the SHEET HEADER BACKUPS sheet.
 * Returns an array of { type, name, headers } objects.
 */
function readHeaderBackupDefinitions_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(HEADER_BACKUP_SHEET);
    if (!sheet) throw new Error('Sheet "' + HEADER_BACKUP_SHEET + '" not found. Cannot perform reset.');

    const data = sheet.getDataRange().getValues();
    const definitions = [];

    for (let r = 0; r < data.length; r++) {
        const type = String(data[r][0] || '').trim().toLowerCase();
        if (!type || type === 'type') continue; // Skip empty rows and header row

        const name = String(data[r][1] || '').trim();
        if (!name) continue;

        // Collect all non-empty values from column C onwards as headers
        const headers = [];
        for (let c = 2; c < data[r].length; c++) {
            const val = String(data[r][c] || '').trim();
            headers.push(val); // Keep even empty values to preserve column positions
        }

        // Trim trailing empty values
        while (headers.length > 0 && headers[headers.length - 1] === '') {
            headers.pop();
        }

        definitions.push({ type: type, name: name, headers: headers });
    }

    return definitions;
}

/**
 * Factory Reset — wipes the spreadsheet back to a clean state.
 * Preserves sheet structure, headers, and named range keys.
 * Requires DUAL confirmation before executing.
 */
function factoryResetSpreadsheet() {
    const ui = SpreadsheetApp.getUi();

    // ═══════════════════════════════════════════
    // DUAL CONFIRMATION — Dialog 1
    // ═══════════════════════════════════════════
    const confirm1 = ui.alert(
        '⚠️ FACTORY RESET — Step 1 of 2',
        'This will DELETE ALL DATA from the spreadsheet and reset it to a clean state.\n\n' +
        '• All orders will be deleted\n' +
        '• All products will be deleted\n' +
        '• All client data will be deleted\n' +
        '• All configuration values will be cleared\n' +
        '• Sheet headers and structure will be preserved\n\n' +
        'This action CANNOT be undone.\n\n' +
        'Do you want to proceed?',
        ui.ButtonSet.YES_NO
    );
    if (confirm1 !== ui.Button.YES) {
        ui.alert('Factory reset cancelled.');
        return;
    }

    // ═══════════════════════════════════════════
    // DUAL CONFIRMATION — Dialog 2
    // ═══════════════════════════════════════════
    const confirm2 = ui.alert(
        '🚨 FINAL WARNING — Step 2 of 2',
        'You are about to PERMANENTLY DELETE all data.\n\n' +
        'Type confirmation: Are you ABSOLUTELY SURE?\n\n' +
        'Click YES to execute the factory reset.\n' +
        'Click NO to cancel.',
        ui.ButtonSet.YES_NO
    );
    if (confirm2 !== ui.Button.YES) {
        ui.alert('Factory reset cancelled.');
        return;
    }

    // ═══════════════════════════════════════════
    // EXECUTE RESET
    // ═══════════════════════════════════════════
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const definitions = readHeaderBackupDefinitions_();
        const log = [];

        definitions.forEach(def => {
            try {
                switch (def.type) {

                    case 'sheet': {
                        const sheet = ss.getSheetByName(def.name);
                        if (!sheet) {
                            log.push('⚠️ Sheet "' + def.name + '" not found — skipped.');
                            break;
                        }

                        const lastRow = sheet.getMaxRows();
                        const headerRowCount = 1; // Headers go in row 1

                        // Clear everything below headers
                        if (lastRow > headerRowCount) {
                            sheet.getRange(headerRowCount + 1, 1, lastRow - headerRowCount, sheet.getMaxColumns()).clearContent().clearFormat();
                        }

                        // Restore headers from backup
                        if (def.headers.length > 0) {
                            sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
                        }

                        log.push('✅ Sheet "' + def.name + '" — cleared data, restored ' + def.headers.length + ' headers.');
                        break;
                    }

                    case 'range': {
                        try {
                            const range = ss.getRangeByName(def.name);
                            if (range) {
                                range.clearContent();
                                log.push('✅ Range "' + def.name + '" — cleared.');
                            } else {
                                log.push('⚠️ Range "' + def.name + '" not found — skipped.');
                            }
                        } catch (e) {
                            log.push('⚠️ Range "' + def.name + '" error: ' + e.message);
                        }
                        break;
                    }

                    case 'table': {
                        // Table format: name = "SheetName::TableHeaderValue" or just a sheet name
                        // Headers in columns C+ are the table column headers
                        const parts = def.name.split('::');
                        const sheetName = parts[0].trim();
                        const tableAnchor = parts.length > 1 ? parts[1].trim() : '';

                        const sheet = ss.getSheetByName(sheetName);
                        if (!sheet) {
                            log.push('⚠️ Table sheet "' + sheetName + '" not found — skipped.');
                            break;
                        }

                        if (tableAnchor && def.headers.length > 0) {
                            // Find the table anchor row (the header text that starts the table)
                            const data = sheet.getDataRange().getValues();
                            let tableStartRow = -1;
                            let tableStartCol = -1;

                            for (let r = 0; r < data.length; r++) {
                                for (let c = 0; c < data[r].length; c++) {
                                    if (String(data[r][c]).trim().toLowerCase() === tableAnchor.toLowerCase()) {
                                        tableStartRow = r + 1; // 1-indexed
                                        tableStartCol = c + 1;
                                        break;
                                    }
                                }
                                if (tableStartRow > -1) break;
                            }

                            if (tableStartRow > -1) {
                                // Write table headers
                                sheet.getRange(tableStartRow, tableStartCol, 1, def.headers.length)
                                    .setValues([def.headers]).setFontWeight('bold');

                                // Clear data rows below (up to 100 rows to be safe)
                                const clearRows = Math.min(100, sheet.getMaxRows() - tableStartRow);
                                if (clearRows > 0) {
                                    sheet.getRange(tableStartRow + 1, tableStartCol, clearRows, def.headers.length)
                                        .clearContent().clearFormat();
                                }

                                log.push('✅ Table "' + def.name + '" — cleared data, restored headers.');
                            } else {
                                log.push('⚠️ Table anchor "' + tableAnchor + '" not found in "' + sheetName + '" — skipped.');
                            }
                        } else {
                            log.push('⚠️ Table "' + def.name + '" — invalid format, expected "SheetName::AnchorHeader".');
                        }
                        break;
                    }

                    default:
                        log.push('⚠️ Unknown type "' + def.type + '" for "' + def.name + '" — skipped.');
                }
            } catch (e) {
                log.push('❌ Error processing "' + def.name + '": ' + e.message);
            }
        });

        // Clear Script Properties (except master/copy registry keys)
        const preserveKeys = ['MASTER_SCRIPT_ID', 'MASTER_SS_ID', 'COPY_REGISTRY', 'IS_MASTER'];
        const props = PropertiesService.getScriptProperties();
        const allProps = props.getProperties();
        Object.keys(allProps).forEach(key => {
            if (!preserveKeys.includes(key)) {
                props.deleteProperty(key);
            }
        });
        log.push('✅ Script Properties cleared (master/copy keys preserved).');

        // Show results
        const resultHtml = HtmlService.createHtmlOutput(
            '<div style="font-family:Roboto,sans-serif;padding:16px;">' +
            '<h3 style="color:#E53935;">🏭 Factory Reset Complete</h3>' +
            '<p style="color:#666;">The spreadsheet has been reset to a clean state.</p>' +
            '<div style="background:#f5f5f5;border-radius:8px;padding:12px;margin-top:12px;' +
            'max-height:300px;overflow-y:auto;font-size:13px;line-height:1.6;">' +
            log.join('<br>') +
            '</div>' +
            '<p style="margin-top:16px;font-size:12px;color:#999;">Version ' + CURRENT_VERSION + '</p>' +
            '</div>'
        ).setWidth(500).setHeight(450);

        ui.showModalDialog(resultHtml, '🏭 Factory Reset Results');

    } catch (e) {
        console.error('[factoryResetSpreadsheet] Error:', e.message, e.stack);
        ui.alert('Factory reset failed: ' + e.message);
    }
}
