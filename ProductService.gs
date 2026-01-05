/**
 * ProductService.gs
 * Handles fetching product catalog
 */

/**
 * Fetch all products from PRODUCTS sheet
 * Structure: [Ref # | SKU | Product Name | Variation Name | Price | Units/Case | Order Amount | Subtotal]
 * We only need the catalog info: SKU, Names, Prices, Units per Case
 */
function getProductCatalog() {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Assuming row 1 is headers and data starts at row 2
    if (lastRow < 2) return [];

    // Result Array
    const catalog = [];

    // 1. Get Headers
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // 2. Map Headers to Indices
    const map = {
        status: -1, sku: -1, ref: -1, name: -1, variation: -1, price: -1, category: -1,
        quantity: -1, image: -1, backgroundColor: -1, description: -1,
        salePrice: -1, onSale: -1 // Keep these just in case, though not in user list
    };

    headers.forEach((h, i) => {
        const header = String(h).trim().toLowerCase();

        if (header === 'status') map.status = i;
        else if (header === 'sku') map.sku = i;
        else if (header === 'reference character' || header === 'ref') map.ref = i;
        else if (header === 'product name') map.name = i;

        // Robust Regex Matching for Variations
        else if (/var.*2|2nd.*var|option|size|strength|dosage|mg/i.test(header) && !header.includes('image')) map.variation2 = i;
        else if (/var.*1/i.test(header) || header === 'variation' || header === 'var' || header === 'variation name') map.variation = i;

        else if (header === 'price') map.price = i;
        else if (['sale', 'sale price'].includes(header)) map.salePrice = i;
        else if (['category', 'cat'].includes(header)) map.category = i;
        else if (['quantity', 'units per case', 'units', 'case'].includes(header)) map.quantity = i;
        else if (/image/i.test(header)) map.image = i;
        else if (['colour', 'color'].includes(header)) map.backgroundColor = i;
        else if (/desc/i.test(header)) map.description = i;

        // Legacy support
        else if (header.includes('on sale')) map.onSale = i;
    });

    // 3. Get Data & Font Colors
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = range.getValues();
    const fontColors = range.getFontColors();

    return data.map((row, rIndex) => {
        // Optional Status Check
        const status = map.status > -1 ? String(row[map.status]).toLowerCase() : "active";
        if (status === 'inactive' || status === 'archived') return null;

        return {
            ref: map.ref > -1 ? row[map.ref] : "",
            sku: map.sku > -1 ? row[map.sku] : "",
            category: map.category > -1 ? row[map.category] : "Uncategorized",
            name: map.name > -1 ? row[map.name] : "",
            variation: map.variation > -1 ? row[map.variation] : "",
            variation2: map.variation2 > -1 ? row[map.variation2] : "", // New
            price: map.price > -1 ? row[map.price] : 0,
            unitsPerCase: map.quantity > -1 ? row[map.quantity] : 1,
            salePrice: map.salePrice > -1 ? row[map.salePrice] : 0,
            onSale: map.onSale > -1 ? Boolean(row[map.onSale]) : (map.salePrice > -1 && Number(row[map.salePrice]) > 0),
            description: map.description > -1 ? row[map.description] : "",
            image: map.image > -1 ? row[map.image] : "",
            backgroundColor: map.backgroundColor > -1 ? row[map.backgroundColor] : "",
            textColor: map.backgroundColor > -1 ? fontColors[rIndex][map.backgroundColor] : ""
        };
    }).filter(p => p && p.sku && p.name);
}

// ... (existing getExistingSkus) ...

/**
 * Add New Products Batch
 * @param {Array} newItems - Array of product objects
 */
function addProductBatch(newItems) {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();

    // Mapping logic
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[String(h).trim().toLowerCase()] = i);

    // Prepare rows
    const rowsToAdd = newItems.map(item => {
        const row = new Array(headers.length).fill("");

        // Helper to set value
        const setVal = (searchKeys, val) => {
            if (!Array.isArray(searchKeys)) searchKeys = [searchKeys];
            for (const key of searchKeys) {
                const lowerKey = key.toLowerCase();
                if (headerMap[lowerKey] !== undefined) {
                    row[headerMap[lowerKey]] = val;
                    return;
                }
            }
        };

        // Standard Mapping
        setVal(['sku'], item.sku);
        setVal(['product name', 'name'], item.name);
        setVal(['category', 'cat'], item.category);
        setVal(['variation 1', 'variation', 'var'], item.variation);
        setVal(['variation 2', 'var 2'], item.variation2); // New
        setVal(['unit price', 'price'], item.price);
        setVal(['units per case', 'units/case', 'quantity', 'units'], item.unitsPerCase || 1);

        // New Fields & Reference
        if (item.ref) setVal(['reference character', 'ref', 'reference'], item.ref);
        if (item.description) setVal(['description', 'desc'], item.description);
        if (item.image) setVal(['image url', 'img', 'image'], item.image);
        if (item.backgroundColor) setVal(['colour', 'color'], item.backgroundColor); // New
        if (item.salePrice) setVal(['sale', 'sale price'], item.salePrice); // New

        return row;
    });

    if (rowsToAdd.length > 0) {
        sheet.getRange(lastRow + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
    }
    return { success: true, count: rowsToAdd.length };
}

/**
 * Group products for easier UI rendering (Optional helper)
 * Validates availability or logic if needed
 */
function getGroupedProducts() {
    const products = getProductCatalog();
    const grouped = {};

    // Group by Product Name
    products.forEach(p => {
        if (!grouped[p.name]) {
            grouped[p.name] = [];
        }
        grouped[p.name].push(p);
    });

    return grouped;
}

/**
 * Get Unique Base Products for Sidebar Dropdown
 * Returns: [{name, category, description, image}, ...]
 */
function getExistingBaseProducts() {
    const products = getProductCatalog();
    const map = new Map();

    products.forEach(p => {
        // Use exact name as key
        const baseName = p.name;
        if (!map.has(baseName)) {
            map.set(baseName, {
                name: baseName,
                category: p.category,
                description: p.description,
                image: p.image
            });
        }
    });

    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name));
}

function getDebugHeaders() {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.map(String);
}

/**
 * Archive Products (Soft Delete)
 * Moves selected SKUs to DELETED_PRODUCTS sheet
 */
function archiveProducts(skusToArchive) {
    if (!skusToArchive || skusToArchive.length === 0) return { success: false, message: "No SKUs provided." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);

    // Ensure Archive Sheet Exists
    let archiveSheet = ss.getSheetByName(SHEET_NAMES.DELETED_PRODUCTS);
    if (!archiveSheet) {
        archiveSheet = ss.insertSheet(SHEET_NAMES.DELETED_PRODUCTS);
        // Copy headers from Products
        const headers = prodSheet.getRange(1, 1, 1, prodSheet.getLastColumn()).getValues();
        archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    }

    const data = prodSheet.getDataRange().getValues();
    const headers = data[0];
    const skuIndex = headers.findIndex(h => String(h).toLowerCase() === 'sku');

    if (skuIndex === -1) return { success: false, message: "SKU Column not found." };

    // Find rows to move
    // We iterate backwards to delete safely
    const rowsToMove = [];
    const rowsToDelete = [];

    for (let i = data.length - 1; i > 0; i--) { // Skip header
        const val = String(data[i][skuIndex]);
        if (skusToArchive.includes(val)) {
            rowsToMove.push(data[i]);
            rowsToDelete.push(i + 1); // 1-based index
        }
    }

    if (rowsToMove.length === 0) return { success: false, message: "No matching products found." };

    // 1. Append to Archive
    // Reverse rowsToMove to maintain approximate order if desired, but not strictly necessary for archive
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove.reverse());

    // 2. Delete from Products
    // Delete one by one? Or batch delete? Batch is hard with non-contiguous.
    // Since we iterated backwards, we can delete safely one by one.
    rowsToDelete.forEach(rowIdx => {
        prodSheet.deleteRow(rowIdx);
    });

    return { success: true, count: rowsToMove.length };
}
