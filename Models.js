/**
 * Models.gs
 * Defines reusable data models for the Order System
 */

/**
 * Product Model Factory
 * Creates a standardized Product object adhering to the Variable Product Model.
 * Handles metadata inheritance (Parent -> Child) and Attribute vs Value logic.
 * 
 * @param {Object} rawData - The raw values from the sheet row
 * @param {Object|null} parentModel - The parent product model if this is a child
 * @param {number} rowIndex - The 0-based index of the row in the sheet data
 * @returns {Object} A standardized product object
 */
function createProductModel(rawData, parentModel, rowIndex) {
    const isParent = String(rawData.node || "").toLowerCase() === 'parent';
    const sheetRow = rowIndex + 2;

    // Helper to parse numeric values safely
    const parseNumber = (val, def = 0) => {
        if (typeof val === 'number') return val;
        let s = String(val || "").replace(/[^0-9.]/g, '');
        let n = parseFloat(s);
        return isNaN(n) ? def : n;
    };

    // THE CORE MODEL
    const product = {
        // Identification
        isParent: isParent,
        id: sheetRow,
        groupId: isParent ? sheetRow : (parentModel ? parentModel.id : sheetRow),

        // Structural Data (Base Info)
        sku: rawData.sku || "",
        ref: rawData.ref || "",
        name: rawData.name || (parentModel ? parentModel.name : ""),
        category: rawData.category || (parentModel ? parentModel.category : "Uncategorized"),
        brand: rawData.brand || (parentModel ? parentModel.brand : ""),

        // Inventory / Availability
        // "0" disables product. Blank or anything else is available.
        inventory: rawData.inventory,
        isAvailable: String(rawData.inventory).trim() !== "0",

        // Attributes (Definitions on Parent, Values on Child)
        // Parent row variation columns store the "Attribute Name" (e.g., "Flavor")
        // Child row variation columns store the "Selection" (e.g., "Blueberry")
        variation: rawData.variation || "",
        variation2: rawData.variation2 || "",
        variation3: rawData.variation3 || "",
        variation4: rawData.variation4 || "",

        // Inherited Attribute Labels (Explicitly for UI Headers)
        headerVariation: parentModel ? parentModel.variation : (isParent ? (rawData.variation || "Flavor") : "Flavor"),
        headerVariation2: parentModel ? parentModel.variation2 : (isParent ? (rawData.variation2 || "Strength") : "Strength"),
        headerVariation3: parentModel ? parentModel.variation3 : (isParent ? (rawData.variation3 || "Format") : "Format"),
        headerVariation4: parentModel ? parentModel.variation4 : (isParent ? (rawData.variation4 || "Units") : "Units"),

        // Pricing & Logistics
        // If child price is 0 or empty, try inheriting from parent
        price: parseNumber(rawData.price || (parentModel ? parentModel.price : 0), 0),
        salePrice: parseNumber(rawData.salePrice || (parentModel ? parentModel.salePrice : 0), 0),
        onSale: (function () {
            const isTrue = (val) => {
                if (val === true || val === 1 || val === '1') return true;
                const s = String(val || "").trim().toLowerCase();
                return s === "true" || s === "yes" || s === "x" || s === "on";
            };

            // STRICT RULE: If this is a child/variation, it ONLY respects the Parent's checkbox.
            // If it's a standalone product or the Parent itself, it respects its own checkbox.
            if (parentModel) {
                // Return child's own value if explicitly set (truthy), otherwise inherit from parent
                const childSale = isTrue(rawData.onSale);
                return childSale || !!parentModel.onSale;
            }

            return isTrue(rawData.onSale);
        })(),

        // FORMAT vs UNITS LOGIC
        // variation3 (Format) defines the physical product tier (Single vs Case/Carton)
        // variation4 (Units) is the commission/piece multiplier
        hasCase: (function () {
            const v1 = String(rawData.variation || "").toLowerCase();
            const v2 = String(rawData.variation2 || "").toLowerCase();
            const v3 = String(rawData.variation3 || "").toLowerCase();
            const v4 = String(rawData.variation4 || "").toLowerCase();
            const units = parseInt(rawData.unitsPerCase) || 1;

            // Directive v1.8.19: Expanded bulk detection (pk, ct, box, box, numbers > 1)
            // FIXED v0.8.76: Added word boundaries \b to prevent partial matches (e.g. "Perfect" triggering "ct")
            const caseRegex = /\b(case|carton|box|multi|pack|bulk|master|pk|ct|disp)\b/i;
            const isNumericCase = (s) => {
                const n = parseInt(s);
                return !isNaN(n) && n > 1 && !s.includes("mg") && !s.includes("ml") && !s.includes("g");
            };

            return caseRegex.test(v1) || caseRegex.test(v2) || caseRegex.test(v3) || caseRegex.test(v4) ||
                isNumericCase(v1) || isNumericCase(v2) || isNumericCase(v3) || isNumericCase(v4) ||
                units > 1;
        })(),
        caseUnits: (parseInt(rawData.unitsPerCase) || 1),
        unitsMultiplier: parseNumber(rawData.variation4 || 1, 1), // Used for Commission and Piece counts
        unitsPerCase: (parseInt(rawData.unitsPerCase) || (parentModel ? parentModel.unitsPerCase : 1)),

        // Commission Data (Inherit from Parent if child value is empty/zero, with bulletproof defaults)
        commissionRate: (function () {
            const childVal = parseNumber(rawData.commissionRate);
            if (childVal > 0) return childVal;
            if (parentModel && parentModel.commissionRate > 0) return parentModel.commissionRate;
            return 1.5;  // Ultimate fallback: $1.50
        })(),
        saleCommission: (function () {
            const childVal = parseNumber(rawData.saleCommission);
            if (childVal > 0) return childVal;
            if (parentModel && parentModel.saleCommission > 0) return parentModel.saleCommission;
            return 1.0;  // Ultimate fallback: $1.00
        })(),

        // Content
        description: rawData.description || (parentModel ? parentModel.description : ""),
        image: rawData.image || (parentModel ? parentModel.image : ""),

        // Styling & Branding (Directives v1.8.17: Comprehensive White/Zero-Width Filtering)
        backgroundColor: (function () {
            const raw = String(rawData.backgroundColor || "").trim().toLowerCase();
            const isWhite = (raw === "#ffffff" || raw === "#fff" || raw === "white" || raw === "transparent" || !raw);
            return isWhite ? "" : raw;
        })(),
        textColor: String(rawData.textColor || "").trim(),
        // Grouping overrides: Always prefer Parent for Group Identity
        groupName: (function () {
            const pName = parentModel ? String(parentModel.name || "").trim() : "";
            if (pName) return pName;
            const rName = String(rawData.name || "").trim();
            if (isParent) return rName || "Unnamed Group";
            return rName || String(rawData.sku || "").trim() || "Unnamed Product";
        })(),
        groupColor: (function () {
            const pBg = String(parentModel ? parentModel.backgroundColor : "").trim().toLowerCase();
            const isPWhite = (pBg === "#ffffff" || pBg === "#fff" || pBg === "white" || !pBg);
            if (!isPWhite) return parentModel.backgroundColor;

            const rBg = String(rawData.backgroundColor || "").trim().toLowerCase();
            const isRWhite = (rBg === "#ffffff" || rBg === "#fff" || rBg === "white" || !rBg);
            if (isParent && !isRWhite) return rawData.backgroundColor;
            return "";
        })(),
        groupTextColor: (parentModel && parentModel.textColor) ? parentModel.textColor : (rawData.textColor || ""),

        // PDF & Export Ranges
        pdfRangeName: rawData.pdfRangeName || (parentModel ? parentModel.pdfRangeName : ""),
        singleRangeName: (rawData.pdfRangeName || (parentModel ? parentModel.pdfRangeName : (rawData.sku || "").split('-')[0])) + "_SINGLE",
        multiRangeName: (rawData.pdfRangeName || (parentModel ? parentModel.pdfRangeName : (rawData.sku || "").split('-')[0])) + "_MULTI",

        zoneVariation: rawData.zoneVariation || (parentModel ? parentModel.zoneVariation : ""),
        orderQty: rawData.orderQty || 0,

        // Metadata
        timestamp: new Date().getTime(),
        version: "1.7.03"
    };

    return product;
}

/**
 * Order Model Factory
 * Standardizes Order data for fulfillment and regulatory tracking.
 * Based on the WooCommerce Order structure.
 * 
 * @param {Object} rawData - The raw order payload (from form or spreadsheet)
 * @returns {Object} A standardized order object
 */
function createOrderModel(rawData) {
    const timestampDate = new Date();
    const formattedDate = Utilities.formatDate(timestampDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

    const model = {
        // Identification
        id: rawData.id || ("ORD-" + timestampDate.getTime()),
        number: rawData.number || String(rawData.id || timestampDate.getTime()),
        status: rawData.status || "pending",

        // Timestamps
        date_created: rawData.date_created || formattedDate,
        date_modified: rawData.date_modified || formattedDate,

        // Client/Customer Information
        customer_id: rawData.clientId || rawData.customer_id || 0,

        // Billing & Shipping (Regulatory & Delivery)
        billing: rawData.billing || {
            first_name: rawData.clientName || "",
            last_name: "",
            address_1: rawData.clientAddress || "",
            city: "",
            state: "",
            postcode: "",
            country: "US",
            email: "",
            phone: ""
        },
        shipping: rawData.shipping || {
            first_name: rawData.clientName || "",
            last_name: "",
            address_1: rawData.clientAddress || "",
            city: "",
            state: "",
            postcode: "",
            country: "US"
        },

        // Itemized List (The core order data)
        line_items: (rawData.items || rawData.line_items || []).map(item => {
            return {
                id: item.id || 0,
                sku: item.sku || "",
                name: item.name || "",
                product_id: item.product_id || 0,
                variation_id: item.variation_id || 0,
                quantity: parseInt(item.quantity) || 0,
                subtotal: String(item.subtotal || "0.00"),
                total: String(item.total || "0.00")
            };
        }),

        // Fulfillment Extras
        shipping_lines: rawData.shipping_lines || [],
        coupon_lines: rawData.coupon_lines || [],

        // Financial Summaries (Regulatory totals)
        discount_total: String(rawData.discount_total || "0.00"),
        shipping_total: String(rawData.shipping_total || "0.00"),
        total: String(rawData.total || "0.00"),

        // Custom Extension Point
        meta_data: rawData.meta_data || []
    };

    return model;
}
