/**
 * ProductService.gs
 * Handles fetching product catalog
 * Version: v1.8.31
 */

/**
 * Fetch all products from PRODUCTS sheet
 * Structure: [Ref # | SKU | Product Name | Variation Name | Price | Units/Case | Order Amount | Subtotal]
 * We only need the catalog info: SKU, Names, Prices, Units per Case
 */
/**
 * Fetch Quantities Summed by Named Ranges for Export
 * Replaces complex frontend summing with single server call
 */

/**
 * Contrast Calculation (YIQ)
 * Determines if black or white text is better for a background color
 */
function getContrastYIQ(hexcolor) {
    if (!hexcolor || hexcolor.length < 4) return "#ffffff";
    if (hexcolor.charAt(0) === '#') hexcolor = hexcolor.substring(1);
    if (hexcolor.length === 3) hexcolor = hexcolor.split('').map(c => c + c).join('');
    const r = parseInt(hexcolor.substr(0, 2), 16);
    const g = parseInt(hexcolor.substr(2, 2), 16);
    const b = parseInt(hexcolor.substr(4, 2), 16);
    const yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
    return (yiq >= 128) ? '#000000' : '#ffffff';
}

/**
 * Safe Price Parsing
 * Removes '$' and handles string parsing
 */
function parsePrice(val) {
    if (typeof val === 'number') return val;
    let s = String(val || "").replace(/[$,]/g, '').trim();
    if (s === "") return 0;
    return parseFloat(s) || 0;
}
/**
 * Fetch all products from PRODUCTS sheet
 * Uses CacheService for repeat load performance (5 min TTL)
 */
function getProductCatalog() {
    // Check cache first
    const cache = CacheService.getScriptCache();
    const cached = cache.get('PRODUCT_CATALOG');
    if (cached) {
        try {
            const parsed = JSON.parse(cached);
            if (parsed && parsed.length > 0) {
                console.log('[getProductCatalog] Cache HIT — ' + parsed.length + ' products');
                return parsed;
            }
        } catch (e) {
            console.warn('[getProductCatalog] Cache parse error, rebuilding.');
        }
    }

    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const headerMap = getProductHeaderMap();
    const map = headerMap.indices;
    const lastCol = headerMap.rawHeaders.length;

    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = range.getValues();
    // PERFORMANCE: Removed getFontColors() — text color now read from data column or auto-calculated

    let lastParent = null; // Reset for this run

    const result = data.map((row, rIndex) => {
        const nodeType = map.node > -1 ? String(row[map.node]).trim().toLowerCase() : "";
        const rowName = map.name > -1 ? String(row[map.name]).trim() : "";

        const isParent = nodeType === 'parent';
        // A row is a variation if its name is empty OR it's explicitly marked as variation/child
        const isVariation = nodeType === 'variation' || nodeType === 'child' || nodeType === 'v' || (rowName === "" && lastParent !== null);

        // Individual Product Sale Active - Support multiples formats (Checkbox, YES, 1, "true")
        const sVal = map.onSale > -1 ? row[map.onSale] : null;
        const sValStr = String(sVal || "").trim().toLowerCase();
        const productSaleChecked = (sVal === true || sVal === 1 || sVal === '1' || (sValStr !== "" && sValStr !== "false" && sValStr !== "no" && sValStr !== "0" && sValStr !== "off"));

        const raw = {
            isParent: isParent,
            ref: map.ref > -1 ? String(row[map.ref]).trim() : "",
            sku: map.sku > -1 ? String(row[map.sku]).trim() : "",
            category: map.category > -1 ? String(row[map.category]).trim() : "",
            name: rowName,
            variation: map.variation > -1 ? String(row[map.variation]).trim() : "",
            variation2: map.variation2 > -1 ? String(row[map.variation2]).trim() : "",
            variation3: map.variation3 > -1 ? String(row[map.variation3]).trim() : "",
            variation4: map.variation4 > -1 ? String(row[map.variation4]).trim() : "",
            variation4: map.variation4 > -1 ? String(row[map.variation4]).trim() : "",
            price: map.price > -1 ? parsePrice(row[map.price]) : 0,
            unitsPerCase: map.unitsPerCase > -1 ? row[map.unitsPerCase] : "",
            salePrice: map.salePrice > -1 ? parsePrice(row[map.salePrice]) : 0,
            onSale: productSaleChecked, // Primary Source of Truth
            description: map.description > -1 ? String(row[map.description]).trim() : "",
            image: map.image > -1 ? String(row[map.image]).trim() : "",
            backgroundColor: map.backgroundColor > -1 ? String(row[map.backgroundColor] || "").trim() : "",
            textColor: (function () {
                // PERFORMANCE: Only resolve for parent rows. Children inherit via createProductModel.
                if (!isParent && lastParent) return ""; // Will inherit from parent in model
                // Priority 1: Dedicated "Text Colour" column (fastest — no formatting API call)
                if (map.textColor > -1) {
                    const tc = String(row[map.textColor] || "").trim();
                    if (tc) return tc;
                }
                // Priority 2: Auto-contrast from background color
                if (map.backgroundColor > -1) {
                    const bg = String(row[map.backgroundColor] || "").trim();
                    if (bg && bg.startsWith('#')) return getContrastYIQ(bg);
                }
                return "";
            })(),
            brand: map.brand > -1 ? String(row[map.brand]).trim() : "",
            zoneVariation: map.zoneVariation > -1 ? String(row[map.zoneVariation]).trim() : "",
            commissionRate: map.commissionRate > -1 ? (row[map.commissionRate] || 1.5) : 1.5,  // Default $1.50
            saleCommission: map.saleCommission > -1 ? (row[map.saleCommission] || 1.0) : 1.0,  // Default $1.00
            node: nodeType,
            orderQty: map.orderQty > -1 ? parseFloat(row[map.orderQty]) || 0 : 0,
            pdfRangeName: map.pdfRangeName > -1 ? String(row[map.pdfRangeName]).trim() : "",

            // Inventory Mapping - Defaults to Column 0 (A) if header not found
            inventory: map.inventory > -1 ? row[map.inventory] : (row[0] || "")
        };

        if (isParent) {
            lastParent = createProductModel(raw, null, rIndex);
            return lastParent;
        }

        // Status Filtering
        const status = map.status > -1 ? String(row[map.status]).toLowerCase() : "active";
        if (status === 'inactive' || status === 'archived') return null;

        // Determine if we should inherit. Only clear lastParent when we hit a new parent row.
        // Don't clear it just because the child lacks explicit marking.
        // (Removed: if (!isVariation) { lastParent = null; })

        // Skip blatant header rows
        const skuVal = String(raw.sku || "").toLowerCase().trim();
        if (skuVal === "sku" || raw.name.toLowerCase().trim() === "product name") return null;

        const model = createProductModel(raw, lastParent, rIndex);

        // DEBUG: Log first 5 products to see what's happening
        if (rIndex < 10) {
            console.log(`DEBUG: Row ${rIndex + 2} | SKU: ${raw.sku} | sVal: ${sVal} | onSale: ${model.onSale}`);
        }

        return model;
    }).filter(p => p && (p.sku || p.isParent) && p.name);

    // Cache the result (max 100KB per key, split if needed)
    try {
        const jsonStr = JSON.stringify(result);
        if (jsonStr.length <= 100000) {
            cache.put('PRODUCT_CATALOG', jsonStr, 300); // 5 min TTL
            console.log('[getProductCatalog] Cached ' + result.length + ' products (' + jsonStr.length + ' bytes)');
        } else {
            console.warn('[getProductCatalog] Catalog too large to cache (' + jsonStr.length + ' bytes). Consider chunking.');
        }
    } catch (e) {
        console.warn('[getProductCatalog] Cache write error:', e.message);
    }

    return result;
}

/**
 * Invalidate the product catalog cache.
 * Call this after any product modification (add, edit, delete).
 */
function invalidateProductCache() {
    try {
        CacheService.getScriptCache().remove('PRODUCT_CATALOG');
        console.log('[invalidateProductCache] Cache cleared.');
    } catch (e) {
        console.warn('[invalidateProductCache] Error:', e.message);
    }
}

/**
 * Dynamic Header Mapper
 */
function getProductHeaderMap() {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const internalKeys = [
        // 1. HIGH SPECIFICITY / COMPOSITE KEYS (Check these first)
        { key: 'totalCommission', aliases: ['total commission', 'sum commission', 'calculated commission', 'comm total'] },
        { key: 'totalPcsOrdered', aliases: ['total pcs ordered', 'total ordered', 'sum qty', 'pieces total'] },
        { key: 'saleCommission', aliases: ['sale commission', 'on sale commission', 'promo commission', 'commission sale'] },
        { key: 'commissionRate', aliases: ['commission rate', 'commission', 'comm', 'normal commission', 'base commission'] },
        { key: 'salePrice', aliases: ['sale price', 'offer price', 'discount price', 'promo price', 'sp', 'promo'] },
        { key: 'onSale', aliases: ['on sale', 'active sale', 'sale status', 'sale active', 'promo active', 'on-sale', 'on sale?', 'sale active?', 'sale', 'sale?', 'sales', 'promo'] },
        { key: 'price', aliases: ['unit price', 'price', 'regular price', 'rp'] },
        { key: 'parentName', aliases: ['parent name', 'parent', 'group name'] },
        { key: 'zoneVariation', aliases: ['zone variation name', 'zone variation', 'zone'] },
        { key: 'pdfRangeName', aliases: ['pdf range name', 'range name', 'range code', 'rn'] },

        // 2. STRUCTURAL / CORE KEYS
        { key: 'node', aliases: ['product node', 'node', 'type', 'node type', 'p/c', 'status type', 'classification'] },
        { key: 'sku', aliases: ['sku', 'item code', 'product code'] },
        { key: 'ref', aliases: ['reference character', 'ref', 'reference', 'ref code'] },
        { key: 'category', aliases: ['category', 'cat', 'department'] },
        { key: 'name', aliases: ['product name', 'name', 'base name', 'item name'] },
        // 3. LOGISTICS & QUANTITY (High Priority for Grid Rendering)
        { key: 'unitsPerCase', aliases: ['units per case', 'units/case', 'case count', 'case size', 'pk size', 'units', 'box size'] },
        { key: 'orderQty', aliases: ['order qty', 'ordered', 'q ordered', 'current order', 'qty'] },
        { key: 'inventory', aliases: ['inventory', 'stock', 'availability', 'stock level', 'quantity in hand'] },

        // 4. VARIATIONS
        { key: 'variation', aliases: ['variation 1', 'var1', 'var 1', 'flavor', 'strain', 'flavour', 'breed'] },
        { key: 'variation2', aliases: ['variation 2', 'var2', 'var 2', 'strength', 'dosage', 'potency'] },
        { key: 'variation3', aliases: ['variation 3', 'var3', 'var 3', 'format', 'pack', 'size/weight'] },
        { key: 'variation4', aliases: ['variation 4', 'var4', 'var 4', 'multiplier', 'comm units'] },
        { key: 'backgroundColor', aliases: ['colour', 'color', 'hex', 'background color'] },
        { key: 'textColor', aliases: ['text color', 'text colour', 'font color', 'font colour', 'txt color'] },
        { key: 'image', aliases: ['image url', 'img', 'image', 'picture'] },
        { key: 'description', aliases: ['description', 'desc', 'product info'] }
    ];

    const indices = {};
    const labels = {};
    const assignedIndices = new Set();

    // Pass 1: Exact Matches
    headers.forEach((h, i) => {
        const head = String(h).trim().toLowerCase();
        if (!head) return;
        for (const config of internalKeys) {
            if (config.aliases.includes(head)) {
                if (indices[config.key] === undefined) {
                    indices[config.key] = i;
                    labels[config.key] = String(h).trim();
                    assignedIndices.add(i);
                }
                break;
            }
        }
    });

    // Pass 2: Loose Matches (for columns not yet assigned)
    headers.forEach((h, i) => {
        if (assignedIndices.has(i)) return;
        const head = String(h).trim().toLowerCase();
        if (!head) return;
        for (const config of internalKeys) {
            if (config.aliases.some(alias => head.includes(alias))) {
                if (indices[config.key] === undefined) {
                    indices[config.key] = i;
                    labels[config.key] = String(h).trim();
                    assignedIndices.add(i);
                }
                break;
            }
        }
    });

    return { indices, labels, rawHeaders: headers };
}

// ... (existing getExistingSkus) ...

function addProductBatch(newItems) {
    invalidateProductCache();
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();
    const headerMap = getProductHeaderMap();
    const indices = headerMap.indices;
    const headers = headerMap.rawHeaders;

    const rowsToAdd = [];

    // NEW: We only add a Parent row if we are creating a NEW product group
    // We check this by seeing if the parent name already exists or if specifically requested
    const first = newItems[0];
    const catalog = getProductCatalog();
    const parentExists = catalog.some(p => p.name.toLowerCase() === first.name.toLowerCase() && p.isParent);

    if (newItems.length > 0 && !parentExists) {
        // Add empty row above parent for visual separation
        rowsToAdd.push(new Array(headers.length).fill(""));

        const parentRow = new Array(headers.length).fill("");

        const setP = (key, val) => {
            if (indices[key] !== undefined) parentRow[indices[key]] = val;
        };

        setP('node', 'Parent');
        setP('brand', first.brand);
        setP('name', first.name);
        setP('parentName', "");
        setP('category', first.category);
        setP('description', first.description || "");
        setP('image', first.image);

        // FIX: Ensure baseSku is added to parent (user request: "base sku was not added to the parent row")
        // Mapping 'sku' column on parent to the baseSku for reference
        setP('sku', first.baseSku);

        // FIX: Ensure color is set. 'backgroundColor' key usually maps to the 'Color' column.
        // We use first.backgroundColor (mapped from 'baseColor' in frontend)
        setP('backgroundColor', first.backgroundColor);
        // Write Text Colour to dedicated column (auto-contrast if blank)
        const parentTextColor = first.textColor || (first.backgroundColor ? getContrastYIQ(first.backgroundColor) : '');
        setP('textColor', parentTextColor);

        setP('ref', first.ref);
        setP('zoneVariation', first.zoneVariation || "");
        setP('commissionRate', first.commissionRate || 1.5);
        setP('saleCommission', first.saleCommission || 1.0);

        // FIX: Variation HEADER Names in Parent Row
        // The frontend sends these as var1Name, var2Name etc in the first item payload?
        setP('variation', first.var1Name || "Variation 1");
        setP('variation2', first.var2Name || "Variation 2");
        setP('variation3', first.var3Name || "Format");
        setP('variation4', first.var4Name || "Units");

        setP('price', "");
        setP('salePrice', "");

        // FIX: Ensure Column A (Inventory) is handled. 
        setP('inventory', "");

        rowsToAdd.push(parentRow);
    }

    newItems.forEach((item, index) => {
        const row = new Array(headers.length).fill("");
        const set = (key, val) => {
            if (indices[key] !== undefined) row[indices[key]] = val;
        };

        // AUTO-SKU: Use provided SKU or generate one
        const finalSku = item.sku || `${item.ref}-${item.baseSku || 'SKU'}-${index + 1}`;

        set('node', 'Child');
        set('brand', item.brand);
        set('sku', finalSku);

        // ADJUSTMENT: If parentName column is missing, keep name in the primary column
        if (indices.parentName !== undefined && indices.parentName > -1) {
            set('name', ""); // Clean look: parentName handles grouping
            set('parentName', item.name);
        } else {
            set('name', item.name); // Essential linking: name is the only key
        }

        set('category', item.category);
        set('variation', item.variation);
        set('variation2', item.variation2);
        set('variation3', item.variation3);
        set('variation4', item.variation4);
        set('price', item.price);
        set('unitsPerCase', item.unitsPerCase || 1);
        set('ref', item.ref);
        set('description', item.description);
        set('image', item.image);
        set('backgroundColor', ""); // Empty Color for Childs
        set('zoneVariation', item.zoneVariation);
        set('commissionRate', item.commissionRate || 1.5);  // Default $1.50
        set('saleCommission', item.saleCommission || 1.0);   // Default $1.00 on sale
        set('salePrice', item.salePrice);
        set('onSale', item.onSale);

        // FIX: "All child products should not as 'Instock' in column A"
        set('inventory', item.inventory || "0");

        rowsToAdd.push(row);
    });

    if (rowsToAdd.length > 0) {
        const startRow = lastRow + 1;
        const targetRange = sheet.getRange(startRow, 1, rowsToAdd.length, headers.length);
        targetRange.setValues(rowsToAdd);

        // --- INJECT FORMULAS ---
        // If !parentExists, we have: [Empty (startRow), Parent (startRow+1), Child1 (startRow+2)...]
        const groupStart = !parentExists ? startRow + 2 : startRow;
        const groupEnd = startRow + rowsToAdd.length - 1;

        if (!parentExists) {
            const pRow = startRow + 1; // Parent Row is now the second row in the batch

            // Parent Total Pcs: =SUM(R[start]:R[end])
            if (indices.totalPcsOrdered !== undefined && indices.orderQty !== undefined) {
                const qtyCol = String.fromCharCode(65 + indices.orderQty);
                sheet.getRange(pRow, indices.totalPcsOrdered + 1).setFormula(`=SUM(${qtyCol}${groupStart}:${qtyCol}${groupEnd})`);
            }
            // Parent Total Commission: =IF(M866,SUM(T867:T872),SUM(S867:S872))
            if (indices.totalCommission !== undefined && indices.onSale !== undefined && indices.commissionRate !== undefined && indices.saleCommission !== undefined) {
                const saleCol = String.fromCharCode(65 + indices.onSale);
                const commCol = String.fromCharCode(65 + indices.commissionRate);
                const saleCommCol = String.fromCharCode(65 + indices.saleCommission);
                const totCommCell = sheet.getRange(pRow, indices.totalCommission + 1);
                totCommCell.setFormula(`=IF(${saleCol}${pRow}, SUM(${saleCommCol}${groupStart}:${saleCommCol}${groupEnd}), SUM(${commCol}${groupStart}:${commCol}${groupEnd}))`);
            }
        }

        // Child Formulas
        const childRows = !parentExists ? newItems.map((_, i) => startRow + 2 + i) : newItems.map((_, i) => startRow + i);
        childRows.forEach((cRow, i) => {
            const parentRow = !parentExists ? startRow + 1 : -1;
            if (parentRow > 0) {
                // Child Order Qty Link: =ORDER_PLACING!C4 (C4, C5, C6...)
                if (indices.orderQty !== undefined) {
                    sheet.getRange(cRow, indices.orderQty + 1).setFormula(`=ORDER_PLACING!C${4 + i}`);
                }
            }
        });

        // APPLY STYLING
        if (!parentExists) {
            const parentColor = first.backgroundColor || "#666666";
            const parentRange = sheet.getRange(startRow + 1, 1, 1, headers.length);
            const textColor = getContrastYIQ(parentColor);
            parentRange.setBackground(parentColor);
            parentRange.setFontColor(textColor);
            parentRange.setFontWeight("bold");
        }

        // AUTO-GENERATE ORDER GRID
        if (!parentExists) {
            try {
                generateOrderPlacerForm(first.name, true);
            } catch (e) {
                console.error("Auto-Grid Error: " + e.message);
            }
        }
    }


    return { success: true, count: rowsToAdd.length };
}

/**
 * Updates an entire product group (Parent + Variations)
 * @param {string} originalBaseName - The current name in the sheet (to find rows)
 * @param {Object} baseInfo - New base data {name, category, description, image, color}
 * @param {Array} variations - Array of variation objects {sku, variation, variation2, price, ...}
 */
function updateProductGroup(originalBaseName, baseInfo, variations) {
    invalidateProductCache();
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) return { success: false, message: "Sheet is empty" };

    const map = mapProductColumns(sheet); // Use helper for column indices
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = range.getValues();
    const updates = []; // Store updates as {r, c, val} to verify before writing? 
    // Actually, bulk write is hard with sparse updates. 
    // We will modify 'data' array in place and write back ONLY if we can do full range, 
    // or use individual setValues for safety if edits are sparse? 
    // Batch write is better for performance. We'll modify 'data' and write back.

    // Index Variations by SKU for quick lookup
    const varMap = {};
    variations.forEach(v => varMap[v.sku] = v);
    const skusFound = new Set();
    const targetName = String(originalBaseName).trim().toLowerCase();
    const headerData = getProductHeaderMap();
    const hMap = headerData.indices;

    // 1. Update Existing Rows
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const nameIdx = hMap.name;
        if (nameIdx === undefined) continue;

        const rowName = String(row[nameIdx]).trim().toLowerCase();
        const skuIdx = hMap.sku;
        const rowSku = (skuIdx !== undefined) ? String(row[skuIdx]).trim() : "";

        if (rowName === targetName) {
            const set = (key, val) => {
                if (hMap[key] !== undefined) row[hMap[key]] = val;
            };

            set('brand', baseInfo.brand);
            set('category', baseInfo.category);
            set('description', baseInfo.description || "Description");
            set('image', baseInfo.image);
            set('backgroundColor', baseInfo.color);
            set('zoneVariation', baseInfo.zoneVariation);
            set('commissionRate', baseInfo.commissionRate);
            set('saleCommission', baseInfo.saleCommission);

            // Row Specific Logic
            const nodeIdx = hMap.node;
            const isParent = nodeIdx !== undefined && String(row[nodeIdx]).trim().toLowerCase() === 'parent';
            if (isParent) {
                set('name', baseInfo.name);
                set('parentName', "");
                set('sku', ""); // Parents have no SKU
                set('variation', baseInfo.var1Name || "");
                set('variation2', baseInfo.var2Name || "");
                set('variation3', baseInfo.var3Name || "");
                set('variation4', baseInfo.var4Name || "");
            } else {
                set('name', ""); // Child name is empty
                set('parentName', baseInfo.name);

                if (rowSku && varMap[rowSku]) {
                    const v = varMap[rowSku];
                    skusFound.add(rowSku);
                    set('variation', v.variation);
                    set('variation2', v.variation2);
                    set('variation3', v.variation3);
                    set('variation4', v.variation4);
                    set('price', v.price);
                    set('salePrice', v.salePrice);
                    set('onSale', v.onSale);
                    set('unitsPerCase', v.unitsPerCase);
                }
            }
        }
    }

    // Write Updates Back
    range.setValues(data);

    // Apply Style to Parent Row(s)
    const nodeIdx = hMap.node;
    if (nodeIdx !== undefined) {
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const nameIdx = hMap.name;
            const rowName = String(row[nameIdx]).trim().toLowerCase();
            const nodeVal = String(row[nodeIdx]).trim().toLowerCase();

            if (rowName === baseInfo.name.toLowerCase() && nodeVal === 'parent') {
                const sheetRow = 2 + i;
                const parentRange = sheet.getRange(sheetRow, 1, 1, lastCol);
                const parentColor = baseInfo.color || "#666666";
                const textColor = getContrastYIQ(parentColor);
                parentRange.setBackground(parentColor);
                parentRange.setFontColor(textColor);
                parentRange.setFontWeight("bold");
            }
        }
    }

    // Identify New Items (Variations in request but not in sheet)
    const newItems = variations.filter(v => !skusFound.has(v.sku));
    let addedCount = 0;
    if (newItems.length > 0) {
        // Prepare new items with Base Info attached
        const itemsToAdd = newItems.map(v => ({
            ...v,
            name: baseInfo.name, // Ensure they get the NEW name
            category: baseInfo.category,
            description: baseInfo.description,
            image: baseInfo.image,
            backgroundColor: baseInfo.color,
            zoneVariation: baseInfo.zoneVariation,
            commissionRate: baseInfo.commissionRate
        }));

        addProductBatch(itemsToAdd);
        addedCount = itemsToAdd.length;
    }
    // REFRESH ORDER GRID (APPEND NEW ONE)
    try {
        generateOrderPlacerForm(baseInfo.name, true);
    } catch (e) {
        console.error("Grid Update Error: " + e.message);
    }


    return { success: true, updated: skusFound.size, added: addedCount };
}

function getGroupedProducts() {
    const products = getProductCatalog();
    const grouped = {};
    products.forEach(p => {
        if (!grouped[p.name]) grouped[p.name] = [];
        grouped[p.name].push(p);
    });
    return grouped;
}

function getExistingBaseProducts() {
    const products = getProductCatalog();
    const map = new Map();
    products.forEach(p => {
        const baseName = p.name;
        if (!map.has(baseName)) {
            map.set(baseName, {
                name: baseName,
                category: p.category,
                description: p.description,
                image: p.image,
                zoneVariation: p.zoneVariation,
                commissionRate: p.commissionRate
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

function archiveProducts(skusToArchive) {
    invalidateProductCache();
    if (!skusToArchive || skusToArchive.length === 0) return { success: false, message: "No SKUs provided." };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);

    let archiveSheet = ss.getSheetByName(SHEET_NAMES.DELETED_PRODUCTS);
    if (!archiveSheet) {
        archiveSheet = ss.insertSheet(SHEET_NAMES.DELETED_PRODUCTS);
        const headers = prodSheet.getRange(1, 1, 1, prodSheet.getLastColumn()).getValues();
        archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    }

    const data = prodSheet.getDataRange().getValues();
    const headers = data[0];
    const skuIndex = headers.findIndex(h => String(h).toLowerCase() === 'sku');
    if (skuIndex === -1) return { success: false, message: "SKU Column not found." };

    const rowsToMove = [];
    const rowsToDelete = [];
    for (let i = data.length - 1; i > 0; i--) {
        const val = String(data[i][skuIndex]);
        if (skusToArchive.includes(val)) {
            rowsToMove.push(data[i]);
            rowsToDelete.push(i + 1);
        }
    }

    if (rowsToMove.length === 0) return { success: false, message: "No matching products found." };
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove.reverse());
    rowsToDelete.forEach(rowIdx => prodSheet.deleteRow(rowIdx));


    return { success: true, count: rowsToMove.length };
}

function mapProductColumns(sheet) {
    return getProductHeaderMap().indices;
}

function cleanupProductSheet() {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const map = mapProductColumns(sheet);
    if (map.node === -1) return;

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const rowsToDelete = [];
    const updates = [];
    let lastKey = null;
    let parentsInGroup = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const sheetRowIdx = i + 2;
        const rawNode = String(row[map.node]).trim();
        const lowNode = rawNode.toLowerCase();
        let pName = map.parentName > -1 ? String(row[map.parentName]).trim() : "";
        let name = String(row[map.name]).trim();
        let currentKey = pName || name;

        if (lowNode.includes('paren') && lowNode !== 'parent') {
            updates.push({ r: sheetRowIdx, c: map.node + 1, val: 'Parent' });
            row[map.node] = 'Parent';
        }

        // Normalize Breed / Variation Data
        if (map.variation > -1) {
            const val = String(row[map.variation]).trim();
            const lowVal = val.toLowerCase();
            let newVal = null;

            if (lowVal === 'ind' || lowVal === 'indica.') newVal = 'Indica';
            else if (lowVal === 'sat' || lowVal === 'sativa.') newVal = 'Sativa';
            else if (lowVal === 'hyb' || lowVal === 'hybrid.') newVal = 'Hybrid';

            if (newVal && newVal !== val) {
                updates.push({ r: sheetRowIdx, c: map.variation + 1, val: newVal });
                row[map.variation] = newVal;
            }
        }

        const isParent = String(row[map.node]).trim().toLowerCase() === 'parent';

        if (currentKey !== lastKey) {
            if (parentsInGroup.length > 1) {
                for (let k = 1; k < parentsInGroup.length; k++) rowsToDelete.push(parentsInGroup[k] + 2);
            }
            parentsInGroup = [];
            lastKey = currentKey;
        }
        if (isParent) parentsInGroup.push(i);
    }

    if (parentsInGroup.length > 1) {
        for (let k = 1; k < parentsInGroup.length; k++) rowsToDelete.push(parentsInGroup[k] + 2);
    }

    // Going GAS: Batch cell updates instead of individual setValue() calls
    if (updates.length > 0) {
        const rangeList = updates.map(u => sheet.getRange(u.r, u.c));
        updates.forEach((u, i) => rangeList[i].setValue(u.val));
        SpreadsheetApp.flush();
    }
    rowsToDelete.sort((a, b) => b - a);
    [...new Set(rowsToDelete)].forEach(r => sheet.deleteRow(r));


    SpreadsheetApp.getActiveSpreadsheet().toast(`Cleanup: Fixed ${updates.length} typos, Deleted ${[...new Set(rowsToDelete)].length} duplicate parents.`);
}

function generateParentRows() {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    const lastRow = sheet.getLastRow();
    const hData = getProductHeaderMap();
    const map = hData.indices;
    const lastCol = hData.rawHeaders.length;

    if (map.node === -1 || map.name === -1) { SpreadsheetApp.getUi().alert("Missing Node or Name column."); return; }

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let lastKey = null;
    let groupStartIndex = 0;
    let groupHasParent = false;
    let rowsAdded = 0;

    for (let i = 0; i <= data.length; i++) {
        const row = (i < data.length) ? data[i] : null;
        let currentKey = null;
        let isParentNode = false;

        if (row) {
            const pName = map.parentName > -1 ? String(row[map.parentName]).trim() : "";
            const name = String(row[map.name]).trim();
            currentKey = pName || name;
            isParentNode = (String(row[map.node]).trim().toLowerCase() === 'parent');
        }

        if (currentKey !== lastKey || i === data.length) {
            if (lastKey !== null && !groupHasParent) {
                const insertAtRow = groupStartIndex + 2 + (rowsAdded - 0); // Corrected offset
                sheet.insertRowBefore(insertAtRow);
                const child = data[groupStartIndex];

                const metaFields = {
                    node: "Parent",
                    brand: map.brand > -1 ? child[map.brand] : "",
                    name: lastKey,
                    parentName: hData.labels.parentName || "Parent Name",
                    category: map.category > -1 ? child[map.category] : "",
                    description: hData.labels.description || "Description",
                    image: map.image > -1 ? child[map.image] : "",
                    backgroundColor: map.backgroundColor > -1 ? child[map.backgroundColor] : "",
                    zoneVariation: map.zoneVariation > -1 ? child[map.zoneVariation] : "",
                    commissionRate: map.commissionRate > -1 ? child[map.commissionRate] : "",
                    ref: map.ref > -1 ? child[map.ref] : "",
                    variation: hData.labels.variation || "Var 1",
                    variation2: hData.labels.variation2 || "Var 2",
                    variation3: hData.labels.variation3 || "Var 3",
                    price: hData.labels.price || "Price",
                    salePrice: hData.labels.salePrice || "Sale Price"
                };

                // Going GAS: Build entire row in memory, write once
                const newRow = new Array(lastCol).fill("");
                for (const [key, val] of Object.entries(metaFields)) {
                    if (map[key] !== undefined && map[key] > -1) {
                        newRow[map[key]] = val;
                    }
                }
                sheet.getRange(insertAtRow, 1, 1, lastCol).setValues([newRow]);

                // Style Parent Row
                const pColor = metaFields.backgroundColor || "#666666";
                const pRange = sheet.getRange(insertAtRow, 1, 1, lastCol);
                pRange.setBackground(pColor);
                pRange.setFontColor(getContrastYIQ(pColor));
                pRange.setFontWeight("bold");

                rowsAdded++;
            }
            lastKey = currentKey;
            groupStartIndex = i;
            groupHasParent = isParentNode;
        } else {
            if (isParentNode) groupHasParent = true;
        }
    }


    SpreadsheetApp.getActiveSpreadsheet().toast(`Generated ${rowsAdded} Parent rows.`);
}

/**
 * Fetch Category Data from SETTINGS
 */
function getCategoryData() {
    const settings = getCategorySettings();
    return Object.keys(settings)
        .filter(key => key !== 'main') // Exclude branding row
        .map(key => ({
            name: settings[key].name || key,
            color: settings[key].color,
            saleActive: settings[key].saleActive,
            order: settings[key].order
        })).sort((a, b) => (a.order || 999) - (b.order || 999));
}

/**
 * Save a new Category to SETTINGS
 */
function saveNewCategory(catData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (!sheet) return { success: false, error: "Settings sheet not found" };

    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let catIdx = -1, colorIdx = -1, saleIdx = -1, orderIdx = -1;

    headers.forEach((h, i) => {
        const head = String(h).trim().toLowerCase();
        if (['category', 'cat', 'category name'].includes(head)) catIdx = i;
        else if (['color', 'colour', 'hex'].includes(head)) colorIdx = i;
        else if (['sale active', 'sale status', 'allow sales', 'sale'].includes(head)) saleIdx = i;
        else if (['order', 'sort', 'sort order'].includes(head)) orderIdx = i;
    });

    if (catIdx === -1) return { success: false, error: "Category column not found in Settings" };

    const newRow = new Array(headers.length).fill("");
    newRow[catIdx] = catData.name;
    if (colorIdx > -1) newRow[colorIdx] = catData.color;
    if (saleIdx > -1) newRow[saleIdx] = catData.saleActive;
    if (orderIdx > -1) newRow[orderIdx] = catData.order || 999;

    sheet.appendRow(newRow);
    if (saleIdx > -1) sheet.getRange(sheet.getLastRow(), saleIdx + 1).insertCheckboxes();

    return { success: true };
}

/**
 * High-Performance Product Sheet Formatter (v1.3.72)
 */
function styleProductHeaders() {
    const sheet = getSheet(SHEET_NAMES.PRODUCTS);
    let lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const map = mapProductColumns(sheet);

    // Safety Check
    if (map.node === -1 || lastRow < 2) return;

    // Get Headers to identify special columns
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let saleCheckboxCol = -1;
    headers.forEach((h, idx) => {
        if (['sale', 'on sale', 'active sale'].includes(String(h).trim().toLowerCase())) saleCheckboxCol = idx;
    });

    // LABEL COLUMNS DETECTION
    // Includes: Variations, Breed, Flavor, Strength, Option, Unit, Case, Qty, Quantity
    // PLUS: SKU, Price, Sale Price, Category
    const labelCols = [];
    headers.forEach((h, i) => {
        const s = String(h).trim().toLowerCase();
        if (s.includes('var') || s.includes('strength') || s.includes('size') || s.includes('flavour') || s.includes('flavor') || s.includes('option') || s.includes('unit') || s.includes('case') || s.includes('qty') || s.includes('quantity') || s.includes('breed')) {
            labelCols.push(i);
        }
        else if (s === 'sku' || s.includes('price') || s.includes('sale') || s.includes('category') || s.includes('cat')) {
            labelCols.push(i);
        }
    });

    // --- PHASE 1: STRUCTURAL CLEANUP (Spacers) ---
    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const rowsToDelete = [];

    // 1a. Identify Deletions (Duplicate Spacers)
    for (let i = allData.length - 1; i >= 1; i--) {
        const row = allData[i];
        const isEmpty = row.every(c => String(c).trim() === "");
        const prevRow = (i - 1 >= 0) ? allData[i - 1] : null;
        const isPrevEmpty = prevRow ? prevRow.every(c => String(c).trim() === "") : false;
        if (isEmpty && isPrevEmpty) rowsToDelete.push(i + 1);
    }

    // Execute Deletions
    if (rowsToDelete.length > 0) {
        rowsToDelete.forEach(r => sheet.deleteRow(r));
        SpreadsheetApp.flush();
        lastRow = sheet.getLastRow();
    }

    // 1b. Identify Insertions (Missing Spacers)
    const currentData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    let rowsAdded = 0;

    for (let i = currentData.length - 1; i >= 1; i--) {
        const row = currentData[i];
        const node = String(row[map.node]).trim().toLowerCase();

        if (node === 'parent') {
            if (i > 1) {
                const prevRow = currentData[i - 1];
                const isPrevEmpty = prevRow.every(c => String(c).trim() === "");
                if (!isPrevEmpty) {
                    sheet.insertRowBefore(i + 1);
                    rowsAdded++;
                }
            }
        }
    }
    if (rowsAdded > 0) {
        SpreadsheetApp.flush();
        lastRow = sheet.getLastRow();
    }

    // --- PHASE 2: BATCH CONTENT & STYLING ---
    const fullRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = fullRange.getValues();
    const bgColors = fullRange.getBackgrounds();
    const fontWeights = fullRange.getFontWeights();
    // Going GAS: Removed getFontColors() — build in memory from data column or auto-contrast
    const fontColors = values.map(row => row.map(() => "black"));

    let currentParentName = "";
    const parentRowIndices = [];

    for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const isEmpty = row.every(c => String(c).trim() === "");

        // CLEAN SPACER
        if (isEmpty) {
            for (let c = 0; c < row.length; c++) {
                bgColors[i][c] = null;
                fontColors[i][c] = "black";
                fontWeights[i][c] = "normal";
                values[i][c] = "";
            }
            continue;
        }

        const node = String(row[map.node]).trim().toLowerCase();
        const pName = map.parentName > -1 ? String(row[map.parentName]) : "";
        const name = String(row[map.name]);

        if (node === 'parent') {
            currentParentName = name || pName;
            parentRowIndices.push(i + 2);

            // STYLE PARENT
            let bgColor = "#e0e0e0";
            let textColor = "black";

            if (map.color > -1) {
                const cellVal = String(row[map.color]).trim();
                const cellBg = bgColors[i][map.color];
                if (cellVal.startsWith('#')) {
                    bgColor = cellVal;
                } else if (cellBg && cellBg !== '#ffffff') {
                    bgColor = cellBg;
                }
                // Text colour: read from data column first, then auto-contrast
                if (map.textColor > -1) {
                    const tc = String(row[map.textColor] || "").trim();
                    if (tc) textColor = tc;
                    else textColor = getContrastYIQ(bgColor);
                } else {
                    textColor = getContrastYIQ(bgColor);
                }
            }

            // Apply to Whole Row
            for (let c = 0; c < lastCol; c++) {
                bgColors[i][c] = bgColor;
                fontColors[i][c] = textColor;
                fontWeights[i][c] = "bold";
            }

            // FILL LABELS
            labelCols.forEach(colIdx => {
                // FORCE: If it's the Category Column, strict "Category" label.
                if (colIdx === map.category) {
                    values[i][colIdx] = "Category";
                }
                // FORCE: For other label columns (SKU, Price, Variations), overwrite with Header Label.
                // This ensures the parent row acts as a Header Row for these columns.
                else {
                    values[i][colIdx] = headers[colIdx];
                }
            });

        } else {
            // CHILD ROW
            // Inherit Parent Name
            if (map.parentName > -1 && currentParentName) {
                values[i][map.parentName] = currentParentName;
            }
            // Clear Product Name ONLY if we have a parentName column to handle the group identity
            if (map.name > -1 && map.parentName > -1) {
                values[i][map.name] = "";
            } else if (map.name > -1 && currentParentName) {
                // If no linking column, ensure the name is populated so grouping doesn't break
                values[i][map.name] = currentParentName;
            }
            // Reset Styles
            for (let c = 0; c < lastCol; c++) {
                bgColors[i][c] = null;
                fontWeights[i][c] = "normal";
            }
        }
    }

    // WRITE BACK (Batch)
    fullRange.setValues(values);
    fullRange.setBackgrounds(bgColors);
    fullRange.setFontColors(fontColors);
    fullRange.setFontWeights(fontWeights);


    // --- PHASE 3: CHECKBOXES ---
    if (saleCheckboxCol > -1) {
        // Remove ALL first
        const saleRange = sheet.getRange(2, saleCheckboxCol + 1, lastRow - 1, 1);
        saleRange.removeCheckboxes();
        saleRange.clearContent();

        // Add back only to Parents
        if (parentRowIndices.length > 0) {
            const ranges = parentRowIndices.map(r => `${sheet.getRange(r, saleCheckboxCol + 1).getA1Notation()}`);
            sheet.getRangeList(ranges).insertCheckboxes();
        }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast("Optimization Complete: Cleaned, Styled, Formatted.");
}
