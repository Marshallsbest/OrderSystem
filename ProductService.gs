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

    // Assuming row 1 is headers and data starts at row 2
    if (lastRow < 2) return [];

    // Read columns A to F (up to Units/Case)
    // Adjust range if Order Amount is needed (but that's usually for manual entry)
    const range = sheet.getRange(2, 1, lastRow - 1, 6);
    const values = range.getValues();

    // Map array to objects
    // Col Indices (0-based): 
    // 0: Ref, 1: SKU, 2: ProdName, 3: VarName, 4: Price, 5: Units/Case
    return values.map(row => ({
        ref: row[0],
        sku: row[1],
        name: row[2],
        variation: row[3],
        price: row[4],
        unitsPerCase: row[5]
    })).filter(p => p.sku && p.name); // Filter empty rows
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
