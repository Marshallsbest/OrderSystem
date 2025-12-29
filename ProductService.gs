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
        ref: -1, sku: -1, category: -1, name: -1, variation: -1, price: -1, unitsPerCase: -1,
        salePrice: -1, onSale: -1, description: -1, image: -1
    };

    headers.forEach((h, i) => {
        const header = String(h).trim().toLowerCase();
        // Loose matching
        if (header.includes('ref')) map.ref = i;
        else if (header.includes('sku')) map.sku = i;
        else if (header.includes('category')) map.category = i;
        else if (header === 'name' || header.includes('product name')) map.name = i;
        else if (header.includes('variation')) map.variation = i;
        else if (header === 'price' || header.includes('unit price') && !header.includes('sale')) map.price = i;
        else if (header.includes('units') && header.includes('case')) map.unitsPerCase = i;

        // New Fields
        else if (header.includes('sale price')) map.salePrice = i;
        else if (header.includes('on sale')) map.onSale = i;
        else if (header.includes('description')) map.description = i;
        else if (header.includes('image')) map.image = i;
    });

    // Fallback? If specific critical columns are missing, maybe we default to standard indices or just fail gracefully.
    // Let's assume standard indices if dynamic fails for SKU or Name
    if (map.sku === -1) map.sku = 1;
    if (map.name === -1) map.name = 3;
    if (map.category === -1) map.category = 2; // Default if not found

    // 3. Get Data
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    return data.map(row => ({
        ref: map.ref > -1 ? row[map.ref] : "",
        sku: map.sku > -1 ? row[map.sku] : "",
        category: map.category > -1 ? row[map.category] : "Uncategorized",
        name: map.name > -1 ? row[map.name] : "",
        variation: map.variation > -1 ? row[map.variation] : "",
        price: map.price > -1 ? row[map.price] : 0,
        unitsPerCase: map.unitsPerCase > -1 ? row[map.unitsPerCase] : "",
        // New Fields
        salePrice: map.salePrice > -1 ? row[map.salePrice] : 0,
        onSale: map.onSale > -1 ? Boolean(row[map.onSale]) : false,
        description: map.description > -1 ? row[map.description] : "",
        image: map.image > -1 ? row[map.image] : ""
    })).filter(p => p.sku && p.name);
}

/**
 * Get Existing SKUs for validation
 */
function getExistingSkus() {
    const products = getProductCatalog();
    return products.map(p => p.sku);
}

/**
 * Add New Products Batch
 * @param {Array} newItems - Array of product objects
 */
function addProductBatch(newItems) {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();

    // We need to map object keys back to column order
    // This assumes a standard column order if we are appending
    // Let's rely on finding headers first to be safe
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[String(h).trim().toLowerCase()] = i);

    // Prepare rows
    const rowsToAdd = newItems.map(item => {
        const row = new Array(headers.length).fill("");

        // Helper to set value if header exists (Fuzzy Match & Multi-Key)
        const setVal = (searchKeys, val) => {
            if (!Array.isArray(searchKeys)) searchKeys = [searchKeys];

            // 1. Try exact matches from our headerMap
            for (const key of searchKeys) {
                const lowerKey = key.toLowerCase();
                if (headerMap[lowerKey] !== undefined) {
                    row[headerMap[lowerKey]] = val;
                    return;
                }
            }

            // 2. Try partial includes if strict match fails (fallback)
            for (const key of searchKeys) {
                const lowerKey = key.toLowerCase();
                const foundKey = Object.keys(headerMap).find(k => k.includes(lowerKey));
                if (foundKey) {
                    row[headerMap[foundKey]] = val;
                    return;
                }
            }
        };

        // Standard Mapping
        setVal(['sku'], item.sku);
        setVal(['product name', 'name'], item.name);
        setVal(['category', 'cat'], item.category);
        setVal(['variation', 'var', 'variation name'], item.variation);
        setVal(['unit price', 'price'], item.price);
        setVal(['units per case', 'units/case', 'case', 'units', 'pc'], item.unitsPerCase || 1);

        // New Fields & Reference
        if (item.ref) setVal(['ref', 'reference', 'ref #'], item.ref);
        if (item.description) setVal(['description', 'desc'], item.description);
        if (item.image) setVal(['image url', 'img', 'image'], item.image);

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
