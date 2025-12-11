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

    // Read columns A to G (up to Units/Case)
    // Adjusted for new "Category" column: [Ref, SKU, Category, Name, Variation, Price, Units/Case]
    const range = sheet.getRange(2, 1, lastRow - 1, 7);
    const values = range.getValues();

    // Map array to objects
    // Col Indices (0-based): 
    // 0: Ref, 1: SKU, 2: Category, 3: Name, 4: Variation, 5: Price, 6: Units/Case
    return values.map(row => ({
        ref: row[0],
        sku: row[1],
        category: row[2], // New field
        name: row[3],
        variation: row[4],
        price: row[5],
        unitsPerCase: row[6]
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
