/**
 * OrderFormPDFService.gs
 *
 * Generates ORDER_FORM_1-style PDFs via two paths:
 *
 *   PATH A — Sheet Export (ORDERS sheet trigger):
 *     Reads ORDER_FORM_1 template dynamically → writes aggregated quantities
 *     into the live sheet → exports the sheet as a native Google Sheets PDF
 *     (preserves all real formatting, merged cells, colors) → clears values.
 *
 *   PATH B — HTML PDF (web form auto-generate on submit):
 *     Same aggregation logic, but renders an HTML document that mirrors the
 *     ORDER_FORM_1 layout, converted to PDF without touching the sheet.
 *
 * Both paths share readOrderFormTemplate() and aggregateOrderByRef() so any
 * structural change to ORDER_FORM_1 is automatically reflected in output.
 *
 * Version: v0.9.19
 */

// ============================================================================
// SHARED: TEMPLATE READER
// ============================================================================

/**
 * Reads the live ORDER_FORM_1 sheet structure every time it is called.
 * Detects header positions (Singles, MC, Total, Ref) dynamically so the
 * output automatically adapts if columns are added, moved, or renamed.
 *
 * @returns {Object|null} Template descriptor:
 *   {
 *     sheet,                  // Sheet object
 *     headerRow,              // 0-based index of the header row
 *     refCol,                 // 0-based col index for the Ref/short-code column
 *     labelCol,               // 0-based col index for the product name column
 *     formatCol,              // 0-based col index for format (may be -1)
 *     priceCol,               // 0-based col index for Price/Unit (may be -1)
 *     mcPriceCol,             // 0-based col index for MC/case price (may be -1)
 *     singlesCol,             // 0-based col index for Singles quantity
 *     mcCol,                  // 0-based col index for MC quantity
 *     totalCol,               // 0-based col index for Total quantity
 *     lastCol,                // total columns in sheet
 *     rows: [                 // one entry per data row below the header
 *       {
 *         type,               // 'category' | 'product' | 'blank'
 *         sheetRowNum,        // 1-based sheet row number (for Range writes)
 *         ref,                // e.g. 'FRR' (product rows only)
 *         label,              // product name text
 *         format,             // format text (e.g. '8-pack')
 *         price,              // price text (e.g. '$36/pack')
 *         mcPrice,            // MC price text (e.g. '$350/MC (10)')
 *         bgColor,            // hex background (category rows)
 *       }
 *     ]
 *   }
 */
function readOrderFormTemplate(sheetName) {
    sheetName = sheetName || 'ORDER_FORM_1';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return null;

    const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const allBg = sheet.getRange(1, 1, lastRow, lastCol).getBackgrounds();

    // ── Find header row ────────────────────────────────────────────────────────
    // Scan up to row 10 for the row that contains "Singles", "MC" / "Multi",
    // and "Total" keywords.
    let headerRow = -1;
    let singlesCol = -1;
    let mcCol = -1;
    let totalCol = -1;

    for (let r = 0; r < Math.min(10, lastRow); r++) {
        const row = allValues[r];
        let hits = 0;
        for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || '').trim().toLowerCase();
            if (cell === 'singles' || cell === 'single') { singlesCol = c; hits++; }
            else if (cell === 'mc' || cell.startsWith('multi') || cell === 'case') { mcCol = c; hits++; }
            else if (cell === 'total') { totalCol = c; hits++; }
        }
        if (hits >= 2) { headerRow = r; break; }
    }

    if (headerRow === -1) throw new Error(
        'ORDER_FORM_1 header row not found. ' +
        'Ensure the sheet has columns labelled "Singles", "MC", and "Total".'
    );

    // ── Detect other column positions from header row ─────────────────────────
    const headerRowData = allValues[headerRow];
    let refCol = -1;
    let labelCol = -1;
    let formatCol = -1;
    let priceCol = -1;
    let mcPriceCol = -1;

    for (let c = 0; c < headerRowData.length; c++) {
        const cell = String(headerRowData[c] || '').trim().toLowerCase();
        if (c === singlesCol || c === mcCol || c === totalCol) continue;

        if (cell === 'ref' || cell === 'ref #' || cell === 'code') refCol = c;
        else if (cell.includes('product') || cell.includes('name') || cell.includes('item')) labelCol = c;
        else if (cell.includes('format') || cell.includes('pack') || cell.includes('size')) formatCol = c;
        else if (cell.includes('price') && !cell.includes('mc')) priceCol = c;
        else if (cell.includes('mc') && cell.includes('price')) mcPriceCol = c;
    }

    // Fallback positional guesses when header labels are absent/different
    // Typical layout: A=blank/inv, B=Ref, C=Name, D=Format, E=Price, F=MCPrice, G=Singles, H=MC, I=Total
    if (refCol === -1) refCol = 1;  // Column B
    if (labelCol === -1) labelCol = 2;  // Column C

    // For format and price: scan between labelCol+1 and singlesCol
    if (formatCol === -1 || priceCol === -1) {
        for (let c = labelCol + 1; c < singlesCol && c < headerRowData.length; c++) {
            if (c === formatCol || c === priceCol || c === mcPriceCol) continue;
            const cell = String(headerRowData[c] || '').trim().toLowerCase();
            if (cell.includes('price') || cell.includes('unit')) {
                if (priceCol === -1) priceCol = c;
                else if (mcPriceCol === -1) mcPriceCol = c;
            } else {
                if (formatCol === -1) formatCol = c;
            }
        }
    }

    // ── Parse data rows ────────────────────────────────────────────────────────
    const rows = [];

    // Helper: true if a hex color is a low-saturation grey (not a real category color)
    const _isGrey = (hex) => {
        try {
            const h = hex.replace('#', '');
            const r = parseInt(h.substr(0, 2), 16);
            const g = parseInt(h.substr(2, 2), 16);
            const b = parseInt(h.substr(4, 2), 16);
            const max = Math.max(r, g, b), min = Math.min(r, g, b);
            return max === 0 ? true : (max - min) / max < 0.15; // saturation < 15%
        } catch (e) { return true; }
    };

    for (let r = headerRow + 1; r < lastRow; r++) {
        const rowData = allValues[r];
        const sheetRowNum = r + 1; // 1-based

        // Determine row background — scan all cells, skip greys, prefer saturated colours
        let rowBg = '#ffffff';
        for (let c = 0; c < Math.min(lastCol, lastCol); c++) {
            const bg = String(allBg[r][c] || '').toLowerCase();
            if (!bg || bg === '#ffffff' || bg === '#000000') continue;
            if (_isGrey(bg)) continue; // skip near-grey cells (e.g. #f3f3f3 in ref column)
            rowBg = bg;
            break;
        }

        const refVal = String(rowData[refCol] || '').trim();
        const labelVal = String(rowData[labelCol] || '').trim();

        // Blank row
        const rowHasContent = rowData.some(c => String(c || '').trim() !== '');
        if (!rowHasContent) {
            rows.push({ type: 'blank', sheetRowNum });
            continue;
        }

        // Category header: coloured background AND no ref code
        if (rowBg !== '#ffffff' && !refVal) {
            // Use only the most prominent cell text (label col or first non-empty)
            const categoryLabel = labelVal ||
                rowData.map(c => String(c || '').trim()).filter(c => c)[0] || '';
            rows.push({ type: 'category', sheetRowNum, label: categoryLabel, bgColor: rowBg });

            // ── STOP after "Shipping" row — everything below is secondary table / totals
            if (categoryLabel.toLowerCase().includes('shipping')) break;
            continue;
        }

        // Product row (has a ref code)
        if (refVal) {
            const formatVal = formatCol > -1 ? String(rowData[formatCol] || '').trim() : '';
            const priceVal = priceCol > -1 ? String(rowData[priceCol] || '').trim() : '';
            const mcPriceVal = mcPriceCol > -1 ? String(rowData[mcPriceCol] || '').trim() : '';

            rows.push({
                type: 'product',
                sheetRowNum,
                ref: refVal.toUpperCase(),
                label: labelVal,
                format: formatVal,
                price: priceVal,
                mcPrice: mcPriceVal,
                bgColor: rowBg,
                singlesQty: 0,
                mcQty: 0,
                totalQty: 0,
            });
            continue;
        }

        // Catch-all: non-blank row with no ref (e.g. sub-header text, notes)
        const anyText = labelVal ||
            rowData.map(c => String(c || '').trim()).filter(c => c)[0] || '';
        if (anyText) {
            rows.push({ type: 'category', sheetRowNum, label: anyText, bgColor: rowBg });
        }
    }

    return {
        sheet,
        headerRow,
        refCol, labelCol, formatCol, priceCol, mcPriceCol,
        singlesCol, mcCol, totalCol,
        lastCol,
        rows,
    };
}

// ============================================================================
// SHARED: ORDER AGGREGATION
// ============================================================================

/**
 * Aggregates ordered items by ref code, splitting single-unit products
 * from multi-pack/case products using the product model's `hasCase` flag.
 *
 * @param {Array}  orderItems  [{sku, quantity, price, onSale}]
 * @param {Array}  catalog     Full product catalog from getProductCatalog()
 * @returns {Object}
 *   {
 *     byForm: {
 *       [formNumber]: {                         // e.g. "1", "2"
 *         byRef: { [refCode]: { singlesQty, mcQty, totalPrice, items[] } },
 *         itemizedDetail: [{ ref, sku, name, variation, quantity, isCase }]
 *       }
 *     }
 *   }
 */
function aggregateOrderByRef(orderItems, catalog) {
    // byForm[formNumber] = { byRef: {...}, itemizedDetail: [...] }
    const byForm = {};

    const _ensureForm = (num) => {
        if (!byForm[num]) byForm[num] = { byRef: {}, itemizedDetail: [] };
    };
    const _ensureRef = (form, ref) => {
        if (!form.byRef[ref]) form.byRef[ref] = { singlesQty: 0, mcQty: 0, totalPrice: 0, items: [] };
    };

    (orderItems || []).forEach(item => {
        const qty = parseInt(item.quantity) || 0;
        if (qty <= 0) return;

        const product = catalog.find(p =>
            String(p.sku || '').trim().toUpperCase() === String(item.sku || '').trim().toUpperCase()
        );

        const ref = product ? String(product.ref || '').trim().toUpperCase() : '?';
        const isCase = product ? !!product.hasCase : false;
        const formNum = product ? String(product.orderFormNumber || '1').trim() : '1';
        const unitPrice = item.price ||
            (product ? (product.onSale && product.salePrice > 0 ? product.salePrice : product.price) : 0);

        _ensureForm(formNum);
        _ensureRef(byForm[formNum], ref);

        const formData = byForm[formNum];
        if (isCase) {
            formData.byRef[ref].mcQty += qty;
        } else {
            formData.byRef[ref].singlesQty += qty;
        }
        formData.byRef[ref].totalPrice += unitPrice * qty;

        const variation = product
            ? [product.variation, product.variation2, product.variation3]
                .filter(v => v && String(v).trim())
                .join(' / ')
            : '';

        formData.byRef[ref].items.push({
            sku: item.sku,
            name: product ? product.name : item.sku,
            variation,
            quantity: qty,
            isCase,
            price: unitPrice,
        });

        formData.itemizedDetail.push({
            ref,
            sku: item.sku,
            name: product ? product.name : item.sku,
            variation,
            quantity: qty,
            isCase,
        });
    });

    return { byForm };
}

// ============================================================================
// PATH A — SHEET POPULATE → NATIVE PDF EXPORT (ORDERS sheet trigger)
// ============================================================================

/**
 * Entry point called from the ORDERS sheet menu.
 * Works on the currently selected row.
 *
 * Workflow:
 *   1. Parse order from selected row
 *   2. Read ORDER_FORM_1 template structure
 *   3. Aggregate order quantities by ref code
 *   4. Write Singles/MC/Total quantities into ORDER_FORM_1
 *   5. Export ORDER_FORM_1 as a native Sheets PDF
 *   6. Clear the written cells (restore blank template)
 *   7. Save PDF to Orders folder → show link
 */
function generateSelectedOrderFormPdf() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== SHEET_NAMES.ORDERS) {
        ss.toast('Please run this from the ORDERS sheet.');
        return;
    }

    const activeRow = sheet.getActiveCell().getRow();
    if (activeRow < 2) {
        ss.toast('Please select an order row first.');
        return;
    }

    const rowData = sheet.getRange(activeRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const orderId = String(rowData[ORDER_COL.INVOICE_NUMBER] || '').trim();
    const clientName = String(rowData[ORDER_COL.CLIENT] || '').trim();
    const orderDate = (rowData[ORDER_COL.TIME_STAMP] instanceof Date)
        ? rowData[ORDER_COL.TIME_STAMP]
        : new Date();

    if (!orderId) { ss.toast('No Invoice Number found in selected row.'); return; }

    // Parse product items from the row
    const items = [];
    for (let i = ORDER_COL.PRODUCTS_START; i < rowData.length; i++) {
        const cell = String(rowData[i] || '').trim();
        if (!cell) continue;
        const m = cell.match(/\[(\d+)\|@?([^\|]+)\|\$?([\d.]+)\|([TF])\]/);
        if (m) items.push({
            quantity: parseInt(m[1]) || 0,
            sku: m[2].trim(),
            price: parseFloat(m[3]) || 0,
            onSale: m[4] === 'T',
        });
    }

    if (items.length === 0) { ss.toast('No products found in this order.'); return; }

    ss.toast(`Building Order Form PDF for ${clientName}…`);

    try {
        const catalog = getProductCatalog();
        const { byForm } = aggregateOrderByRef(items, catalog);
        const formNums = Object.keys(byForm);

        if (formNums.length === 0) {
            ss.toast('No products matched a known Order Form.');
            return;
        }
        // Read sales rep from CFG_SALES_REP named range (same source as PDFService)
        let cfgSalesRep = '';
        try {
            const repRange = ss.getRangeByName('CFG_SALES_REP');
            if (repRange) cfgSalesRep = String(repRange.getValue() || '').trim();
        } catch (e) {
            Logger.log('Could not read CFG_SALES_REP: ' + e.message);
        }
        if (!cfgSalesRep) cfgSalesRep = getSettingValue('CFG_SALES_REP') || getSettingValue('Sales Rep') || '';

        const orderData = {
            id: orderId,
            clientName,
            clientAddress: String(rowData[ORDER_COL.ADDRESS] || '').trim(),
            clientComments: String(rowData[ORDER_COL.COMMENT] || '').trim(),
            salesRep: cfgSalesRep || String(rowData[ORDER_COL.SALES_REP] || '').trim(),
            date: orderDate,
            items,
        };

        const formattedDate = formatDateWithOrdinal(orderDate);
        const folder = getOrdersFolder();
        const pdfUrls = [];

        formNums.forEach(formNum => {
            const sheetName = 'ORDER_FORM_' + formNum;

            // Native Sheets PDF — includes itemized detail written into secondary table
            const page1Url = _populateSheetAndExport({
                id: orderId,
                clientName,
                clientAddress: orderData.clientAddress,
                clientComments: orderData.clientComments,
                salesRep: orderData.salesRep,
                date: orderDate,
                items,
                _formNum: formNum,
                _byForm: byForm,
                _sheetName: sheetName,
            });

            pdfUrls.push({ formNum, url: page1Url });
            Logger.log('Order Form PDF (native): ' + page1Url);
        });

        const links = pdfUrls.map(u =>
            `<p><a href="${u.url}" target="_blank"
                  style="color:#b040b0;font-weight:700;font-size:14px;">
              📋 ORDER_FORM_${u.formNum} — Order Form PDF
            </a></p>`
        ).join('');

        const html = HtmlService.createHtmlOutput(`
          <div style="font-family:Arial,sans-serif;padding:20px;">
            <p style="color:#b040b0;font-size:16px;font-weight:700;">✓ PDF created!</p>
            ${links}
          </div>
        `).setWidth(400).setHeight(140);

        SpreadsheetApp.getUi().showModalDialog(html, 'Order Form PDF Ready');

    } catch (e) {
        SpreadsheetApp.getUi().alert('Error generating Order Form PDF:\n' + e.message);
        Logger.log('OrderFormPDF (Sheet) Error: ' + e.toString());
    }
}

/**
 * Build ONLY the itemized detail page HTML (page 2).
 * Used by Path A so we can generate a separate companion detail PDF
 * while the main order form PDF comes from the native sheet export.
 */
function _buildDetailPageHtml(itemizedDetail, orderData, formattedDate) {
    const clientName = orderData.clientName || 'Unknown Client';
    const salesRep = orderData.salesRep || '';
    const displayRep = String(salesRep).split(/\s+/)[0] || salesRep;

    const ordered = (itemizedDetail || [])
        .filter(i => i.quantity > 0)
        .sort((a, b) => a.ref.localeCompare(b.ref) || (a.variation || a.name).localeCompare(b.variation || b.name));

    const half = Math.ceil(ordered.length / 2);
    const leftCol = ordered.slice(0, half);
    const rightCol = ordered.slice(half);
    const maxRows = Math.max(leftCol.length, rightCol.length);

    let rows = '';
    for (let i = 0; i < maxRows; i++) {
        const L = leftCol[i];
        const R = rightCol[i];
        const strip = i % 2 === 0 ? '#ffffff' : '#f5f5f5';
        rows += `
      <tr style="background:${strip};">
        <td style="text-align:center;font-weight:700;font-size:9px;color:#888;width:32px;border:1px solid #ddd;padding:3px 4px;">${L ? _esc(L.ref) : ''}</td>
        <td style="padding:3px 7px;border:1px solid #ddd;font-size:10px;">${L ? _esc(L.variation || L.name) : ''}</td>
        <td style="text-align:center;font-weight:700;width:44px;border:1px solid #ddd;padding:3px 4px;">${L ? L.quantity : ''}</td>
        <td style="width:10px;background:#ddd;border:none;"></td>
        <td style="text-align:center;font-weight:700;font-size:9px;color:#888;width:32px;border:1px solid #ddd;padding:3px 4px;">${R ? _esc(R.ref) : ''}</td>
        <td style="padding:3px 7px;border:1px solid #ddd;font-size:10px;">${R ? _esc(R.variation || R.name) : ''}</td>
        <td style="text-align:center;font-weight:700;width:44px;border:1px solid #ddd;padding:3px 4px;">${R ? R.quantity : ''}</td>
      </tr>`;
    }

    return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:Arial,Helvetica,sans-serif; font-size:11px; color:#1a1a1a; }
  .page { padding:20px 24px; max-width:780px; margin:0 auto; }
  .hdr { font-size:15px; font-weight:800; color:#cc66cc;
         border-bottom:2px solid #cc66cc; padding-bottom:5px; margin-bottom:6px; }
  .sub { font-size:10px; color:#555; margin-bottom:12px; }
  table { width:100%; border-collapse:collapse; font-size:10px; }
  th { font-weight:700; font-size:9px; padding:4px 7px;
       text-transform:uppercase; letter-spacing:0.3px; border:1px solid #b050b0; }
</style></head><body>
<div class="page">
  <div class="hdr">Order Detail — ${_esc(clientName)}</div>
  <div class="sub">Date: ${formattedDate} &nbsp;|&nbsp; Order #${_esc(String(orderData.id || 'N/A'))} &nbsp;|&nbsp; Rep: ${_esc(displayRep)}</div>
  <table>
    <thead>
      <tr>
        <th bgcolor="#cc66cc" style="text-align:center;width:32px;color:#fff;">Ref</th>
        <th bgcolor="#cc66cc" style="text-align:left;color:#fff;">Flavour / Strain</th>
        <th bgcolor="#cc66cc" style="text-align:center;width:46px;color:#fff;">Qty</th>
        <th bgcolor="#999" style="width:10px;border:none;"></th>
        <th bgcolor="#cc66cc" style="text-align:center;width:32px;color:#fff;">Ref</th>
        <th bgcolor="#cc66cc" style="text-align:left;color:#fff;">Flavour / Strain</th>
        <th bgcolor="#cc66cc" style="text-align:center;width:46px;color:#fff;">Qty</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>
</div>
</body></html>`;
}

/**
 * Internal: write quantities into ORDER_FORM_1, export as PDF, then clear.
 *
 * @param {Object} orderData  Standard order object with items array,
 *                            optionally with _formNum, _byForm, _sheetName for pre-computed aggregation.
 * @returns {string}          URL of the saved PDF file
 */
function _populateSheetAndExport(orderData) {
    const lock = LockService.getScriptLock();
    try { lock.waitLock(30000); } catch (e) { throw new Error('System busy — try again in a moment.'); }

    const cellsToReset = []; // Track every cell we write so we can clear cleanly

    try {
        // Accept pre-computed aggregation (from generateSelectedOrderFormPdf) or compute fresh
        const byRef = orderData._byForm
            ? (orderData._byForm[orderData._formNum] || {}).byRef || {}
            : (() => {
                const catalog = getProductCatalog();
                const { byForm } = aggregateOrderByRef(orderData.items, catalog);
                return (byForm[orderData._formNum || '1'] || {}).byRef || {};
            })();

        const sheetName = orderData._sheetName || getOrderFormSheetName(orderData._formNum || '1');
        const template = readOrderFormTemplate(sheetName);
        if (!template) throw new Error(sheetName + ' is empty.');

        const formSheet = template.sheet;

        // ── Write header info (Rep / Date / Store / Shipping) ─────────────────
        // Scan rows 1-5 for cells containing label keywords and write values
        // into the cell immediately to the right of each label.
        const headerSearchRange = formSheet.getRange(1, 1, Math.min(5, formSheet.getLastRow()), formSheet.getLastColumn());
        const headerVals = headerSearchRange.getValues();
        const orderDate = orderData.date instanceof Date ? orderData.date : new Date();
        const dateStr = formatDateWithOrdinal(orderDate);

        headerVals.forEach((row, r) => {
            row.forEach((cell, c) => {
                const s = String(cell || '').trim().toLowerCase();
                const destCol = c + 1; // Check the cell to the right
                if (destCol >= row.length) return;
                if (s === 'date:' || s === 'date') {
                    const rangeAddr = formSheet.getRange(r + 1, destCol + 1);
                    rangeAddr.setValue(dateStr);
                    cellsToReset.push(rangeAddr);
                } else if (s === 'store:' || s === 'store') {
                    const rangeAddr = formSheet.getRange(r + 1, destCol + 1);
                    rangeAddr.setValue(orderData.clientName || '');
                    cellsToReset.push(rangeAddr);
                } else if (s === 'rep:' || s === 'rep') {
                    const rangeAddr = formSheet.getRange(r + 1, destCol + 1);
                    rangeAddr.setValue(orderData.salesRep || '');
                    cellsToReset.push(rangeAddr);
                }
                // "Shipping:" left blank unless orderData.shipping provided
            });
        });

        // ── Write Singles / MC / Total quantities into product rows ───────────
        template.rows.forEach(row => {
            if (row.type !== 'product') return;

            const data = byRef[row.ref];
            if (!data) return; // Product not in this order — leave blank

            const singlesQty = data.singlesQty;
            const mcQty = data.mcQty;
            const totalPrice = data.totalPrice || 0; // dollar total

            if (template.singlesCol > -1 && singlesQty > 0) {
                const r = formSheet.getRange(row.sheetRowNum, template.singlesCol + 1);
                r.setValue(singlesQty);
                cellsToReset.push(r);
            }
            if (template.mcCol > -1 && mcQty > 0) {
                const r = formSheet.getRange(row.sheetRowNum, template.mcCol + 1);
                r.setValue(mcQty);
                cellsToReset.push(r);
            }
            if (template.totalCol > -1 && totalPrice > 0) {
                const r = formSheet.getRange(row.sheetRowNum, template.totalCol + 1);
                // Write as a raw number so the sheet's existing currency format applies
                r.setValue(totalPrice);
                cellsToReset.push(r);
            }
        });

        // ── Calculate and write Grand Total to the "Total:" cell ──────────────
        const grandTotal = Object.values(byRef).reduce((sum, d) => sum + (d.totalPrice || 0), 0);
        const lastTemplateRow = template.rows.length > 0
            ? template.rows[template.rows.length - 1].sheetRowNum
            : formSheet.getLastRow();

        if (grandTotal > 0) {
            const scanEnd = Math.min(lastTemplateRow + 12, formSheet.getMaxRows()); // +12 to cover spacer blank rows
            const lastSheetCol = formSheet.getLastColumn();

            for (let scanRow = lastTemplateRow; scanRow <= scanEnd; scanRow++) {
                const scanVals = formSheet.getRange(scanRow, 1, 1, lastSheetCol).getValues()[0];
                const totalLabelIdx = scanVals.findIndex(c => /^\s*total:?\s*$/i.test(String(c || '')));
                if (totalLabelIdx > -1) {
                    const writeCol = (totalLabelIdx + 2 <= lastSheetCol)
                        ? totalLabelIdx + 2
                        : lastSheetCol;
                    const totalCell = formSheet.getRange(scanRow, writeCol);
                    totalCell.setValue(grandTotal);
                    cellsToReset.push(totalCell);
                    Logger.log(`Grand total $${grandTotal} → row ${scanRow} col ${writeCol}`);
                    break;
                }
            }
        }

        // ── Write itemized detail into the secondary table (page 2 of sheet) ──
        const itemizedDetail = (orderData._byForm && orderData._formNum)
            ? ((orderData._byForm[orderData._formNum] || {}).itemizedDetail || [])
            : [];

        Logger.log(`Itemized detail count: ${itemizedDetail.length}`);

        if (itemizedDetail.length > 0) {
            const secScanStart = lastTemplateRow + 1;
            const secScanEnd = Math.min(secScanStart + 20, formSheet.getMaxRows()); // use getMaxRows to catch header rows in blank area
            const lastSheetCol = formSheet.getLastColumn();
            let secHeaderRow = -1, nameCol1 = -1, qtyCol1 = -1, nameCol2 = -1, qtyCol2 = -1;

            Logger.log(`Scanning for secondary table header: rows ${secScanStart}–${secScanEnd}`);

            for (let r = secScanStart; r <= secScanEnd; r++) {
                const vals = formSheet.getRange(r, 1, 1, lastSheetCol).getValues()[0];
                const hasNameHeader = vals.some(v => /flower|concentrate|flavour|strain/i.test(String(v || '')));
                const hasQtyHeader = vals.some(v => /quantity|qty/i.test(String(v || '')));
                Logger.log(`  row ${r}: hasNameHeader=${hasNameHeader} hasQtyHeader=${hasQtyHeader} vals=[${vals.join('|')}]`);
                if (hasNameHeader && hasQtyHeader) {
                    secHeaderRow = r;
                    vals.forEach((v, i) => {
                        const s = String(v || '').toLowerCase();
                        if (/flower|concentrate|flavour|strain/i.test(s)) {
                            if (nameCol1 === -1) nameCol1 = i + 1;
                            else if (nameCol2 === -1) nameCol2 = i + 1;
                        } else if (/quantity|qty/i.test(s)) {
                            if (qtyCol1 === -1) qtyCol1 = i + 1;
                            else if (qtyCol2 === -1) qtyCol2 = i + 1;
                        }
                    });
                    break;
                }
            }

            Logger.log(`Secondary table: headerRow=${secHeaderRow} nameCol1=${nameCol1} qtyCol1=${qtyCol1} nameCol2=${nameCol2} qtyCol2=${qtyCol2}`);

            // If name columns had unrecognised/blank headers, infer from Quantity columns:
            // the name column is immediately to the LEFT of each Quantity column.
            if (qtyCol1 > 1 && nameCol1 === -1) nameCol1 = qtyCol1 - 1;
            if (qtyCol2 > 1 && nameCol2 === -1) nameCol2 = qtyCol2 - 1;

            Logger.log(`After inference: nameCol1=${nameCol1} nameCol2=${nameCol2}`);

            if (secHeaderRow > 0 && qtyCol1 > 0) {
                // NOTE: we do NOT write to or clear the header cells — the user has
                // set "Flavour / Strain" / "Quantity" permanently in the template.
                // Only the DATA rows below the header are written and cleared.

                const ordered = itemizedDetail
                    .filter(i => (i.quantity || 0) > 0)
                    .sort((a, b) => (a.ref || '').localeCompare(b.ref || '') ||
                        (a.variation || a.name || '').localeCompare(b.variation || b.name || ''));
                const half = Math.ceil(ordered.length / 2);
                const leftItems = ordered.slice(0, half);   // fill left column first
                const rightItems = ordered.slice(half);      // overflow into right column
                const refCol1 = nameCol1 > 1 ? nameCol1 - 1 : -1;
                const refCol2 = nameCol2 > 1 ? nameCol2 - 1 : -1;
                const maxRows = formSheet.getMaxRows();

                leftItems.forEach((item, i) => {
                    const rowNum = secHeaderRow + 1 + i;
                    if (rowNum > maxRows) return;
                    if (refCol1 > 0) { const rc = formSheet.getRange(rowNum, refCol1); rc.setValue(item.ref); cellsToReset.push(rc); }
                    if (nameCol1 > 0) { const nc = formSheet.getRange(rowNum, nameCol1); nc.setValue(item.variation || item.name); cellsToReset.push(nc); }
                    if (qtyCol1 > 0) { const qc = formSheet.getRange(rowNum, qtyCol1); qc.setValue(item.quantity); cellsToReset.push(qc); }
                });

                rightItems.forEach((item, i) => {
                    const rowNum = secHeaderRow + 1 + i;
                    if (rowNum > maxRows) return;
                    if (refCol2 > 0) { const rc = formSheet.getRange(rowNum, refCol2); rc.setValue(item.ref); cellsToReset.push(rc); }
                    if (nameCol2 > 0) { const nc = formSheet.getRange(rowNum, nameCol2); nc.setValue(item.variation || item.name); cellsToReset.push(nc); }
                    if (qtyCol2 > 0) { const qc = formSheet.getRange(rowNum, qtyCol2); qc.setValue(item.quantity); cellsToReset.push(qc); }
                });

                Logger.log(`Itemized: ${ordered.length} items — ${leftItems.length} left / ${rightItems.length} right`);



            }
        }

        // ── Resolve tokens anywhere in the template ───────────────────────────
        // Product tokens: {BBS-Q}, {BBS-T}, {BBS-QS}, {BBS-QM}
        // Order tokens:   {CLIENT_NAME}, {CLIENT_ADDRESS}, {CLIENT_PHONE},
        //                 {CLIENT_EMAIL}, {COMMENTS}, {SALES_REP}, {ORDER_TOTAL}
        _resolveTemplateTokens(formSheet, byRef, orderData, cellsToReset);

        SpreadsheetApp.flush(); // Ensure all values are committed


        // ── Export as native Sheets PDF ────────────────────────────────────────
        // No r1/r2 range restriction — let the sheet's own print settings / page
        // breaks control what appears on each page (avoids dark extra page).
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const ssId = ss.getId();
        const sheetId = formSheet.getSheetId();

        // Build export URL — Sheets PDF export API
        const exportUrl = [
            'https://docs.google.com/spreadsheets/d/', ssId,
            '/export?',
            'format=pdf',
            '&size=letter',
            '&portrait=true',
            '&fitw=true',
            '&sheetnames=false',
            '&printtitle=false',
            '&pagenumbers=false',
            '&gridlines=false',
            '&fzr=false',
            '&gid=', sheetId,   // Only this sheet — no range restriction
        ].join('');

        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(exportUrl, {
            headers: { Authorization: 'Bearer ' + token },
            muteHttpExceptions: true,
        });

        if (response.getResponseCode() !== 200) {
            throw new Error('PDF export failed (HTTP ' + response.getResponseCode() + '). ' +
                'Ensure the script has Spreadsheets scope.');
        }

        const clientName = String(orderData.clientName || 'Unknown').trim();
        const salesRep = String(orderData.salesRep || '').trim();
        const repSuffix = salesRep ? ` - ${salesRep}` : '';
        const fileName = `${clientName} ${dateStr}${repSuffix}.pdf`;
        const pdfBlob = response.getBlob().setName(fileName);

        const folder = getOrdersFolder();
        const pdfFile = folder.createFile(pdfBlob);

        Logger.log('Order Form PDF (Sheet export) saved: ' + pdfFile.getUrl());
        return pdfFile.getUrl();

    } finally {
        // ── Restore the blank template ─────────────────────────────────────────
        // Regular data cells → clearContent()
        // Token cells ({REF-Q} etc.) → setValue(original) to restore the token
        cellsToReset.forEach(item => {
            try {
                if (item && item._token) {
                    item.range.setValue(item.original); // restore {REF-Q} text
                } else {
                    item.clearContent();                // clear written data
                }
            } catch (e) { /* ignore individual cell errors */ }
        });
        SpreadsheetApp.flush();
        lock.releaseLock();
    }

}

// ============================================================================
// TOKEN SUBSTITUTION  —  ORDER_FORM template placeholders
// ============================================================================

/**
 * Scans every cell in the ORDER_FORM sheet for placeholder tokens.
 *
 * PRODUCT TOKENS  (keyed by REF, place anywhere in the template):
 *   {REF-Q}          Total units ordered (singles + cases combined)
 *   {REF-QS}         Singles quantity only
 *   {REF-QM}         MC / case quantity only
 *   {REF-T}          Total dollar amount for that product group
 *
 * ORDER-LEVEL TOKENS:
 *   {CLIENT_NAME}    Client / store name
 *   {CLIENT_ADDRESS} Client delivery address
 *   {CLIENT_PHONE}   Client phone number
 *   {CLIENT_EMAIL}   Client email address
 *   {COMMENTS}       Order comments / special instructions
 *   {SALES_REP}      Sales rep name
 *   {ORDER_TOTAL}    Grand total dollar amount
 *   {ORDER_DATE}     Order date (formatted)
 *   {ORDER_ID}       Invoice / order number
 */
function _resolveTemplateTokens(sheet, byRef, orderData, cellsToReset) {
    if (!sheet || !byRef) return;

    const od = orderData || {};
    const grandTotal = Object.values(byRef).reduce((s, d) => s + (d.totalPrice || 0), 0);

    const orderTokens = {
        'CLIENT_NAME': String(od.clientName || ''),
        'CLIENT_ADDRESS': String(od.clientAddress || ''),
        'CLIENT_PHONE': String(od.clientPhone || od.phone || ''),
        'CLIENT_EMAIL': String(od.clientEmail || od.email || ''),
        'COMMENTS': String(od.clientComments || od.comments || od.notes || ''),
        'SALES_REP': String(od.salesRep || ''),
        'ORDER_TOTAL': grandTotal.toFixed(2),
        'ORDER_DATE': od.date instanceof Date
            ? formatDateWithOrdinal(od.date)
            : String(od.date || new Date().toLocaleDateString()),
        'ORDER_ID': String(od.id || od.orderId || ''),
    };

    const TOKEN_RE = /\{([A-Z0-9_]+)(?:-(Q|QS|QM|T))?\}/gi;

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const numRows = values.length;
    const numCols = values[0] ? values[0].length : 0;

    if (numRows === 0 || numCols === 0) return;

    // Track which cells were changed so we can write in one batch + reset later
    const changedCells = []; // { row, col, original, resolved }

    for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
            const raw = String(values[r][c] || '');
            if (!raw.includes('{')) continue;  // fast skip — no token in this cell

            const resolved = raw.replace(TOKEN_RE, (match, key, suffix) => {
                const upperKey = key.toUpperCase();

                // Order-level token (no suffix)
                if (!suffix && orderTokens.hasOwnProperty(upperKey)) {
                    return orderTokens[upperKey];
                }
                // Product-level token (has suffix)
                if (suffix) {
                    const data = byRef[upperKey];
                    if (!data) return '0';
                    switch (suffix.toUpperCase()) {
                        case 'Q': return String((data.singlesQty || 0) + (data.mcQty || 0));
                        case 'QS': return String(data.singlesQty || 0);
                        case 'QM': return String(data.mcQty || 0);
                        case 'T': return Number(data.totalPrice || 0).toFixed(2);
                    }
                }
                return match; // Unknown — leave as-is
            });

            if (resolved !== raw) {
                changedCells.push({ row: r + 1, col: c + 1, original: raw, resolved });
            }
        }
    }

    if (changedCells.length === 0) return;

    Logger.log(`[TokenSubstitution] Resolving ${changedCells.length} token cell(s).`);

    changedCells.forEach(({ row, col, original, resolved }) => {
        const cell = sheet.getRange(row, col);
        cell.setValue(resolved);
        // Restore the original token text after PDF export so the template
        // stays reusable. We store the original value in the cell object's
        // clearContent replacement by re-setting the token string.
        // Trick: push a custom reset object alongside the regular Range ones.
        cellsToReset.push({ _token: true, range: cell, original });
        Logger.log(`  [Token] (R${row},C${col}) "${original}" → "${resolved}"`);
    });
}


// ============================================================================
// PATH B — HTML PDF (auto-generated on web form submission)
// ============================================================================


/**
 * Called by processOrder() in OrderService.js immediately after saving the
 * order row, replacing the old generateOrderPdf() call for ORDER_FORM_1 orders.
 *
 * Builds a two-page HTML document:
 *   Page 1 — ORDER_FORM_1 summary (mirrors template layout)
 *   Page 2 — Itemized detail grid (two-column Ref | Flavour/Strain | Qty)
 *
 * @param {Object} orderData  Standard order object with items array
 * @returns {string}          URL of the saved PDF file
 */
function generateOrderFormHtmlPdf(orderData) {
    // ── Now uses the same native-sheet export as the manual menu trigger ────────
    // (The function name is kept for backward-compat with OrderService.js)
    // NOTE: No outer lock needed — _populateSheetAndExport manages its own lock.

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Read sales rep from CFG_SALES_REP (same as generateSelectedOrderFormPdf)
    let cfgSalesRep = '';
    try {
        const repRange = ss.getRangeByName('CFG_SALES_REP');
        if (repRange) cfgSalesRep = String(repRange.getValue() || '').trim();
    } catch (e) { /* ignore */ }
    if (!cfgSalesRep) cfgSalesRep = getSettingValue('CFG_SALES_REP') || getSettingValue('Sales Rep') || '';

    const enrichedOrderData = Object.assign({}, orderData, {
        salesRep: cfgSalesRep || String(orderData.salesRep || '').trim(),
    });

    const catalog = getProductCatalog();
    const { byForm } = aggregateOrderByRef(enrichedOrderData.items || [], catalog);
    const formNums = Object.keys(byForm);
    let lastPdfUrl = '';

    formNums.forEach(formNum => {
        const sheetName = getOrderFormSheetName(formNum); // admin-configurable via SETTINGS
        const url = _populateSheetAndExport({
            id: enrichedOrderData.id,
            clientName: enrichedOrderData.clientName,
            clientAddress: enrichedOrderData.clientAddress,
            clientComments: enrichedOrderData.clientComments,
            salesRep: enrichedOrderData.salesRep,
            date: enrichedOrderData.date,
            items: enrichedOrderData.items,
            _formNum: formNum,
            _byForm: byForm,
            _sheetName: sheetName,
        });
        if (url) lastPdfUrl = url;
    });

    return lastPdfUrl;
}


/**
 * Build the full two-page HTML for the HTML-based PDF path.
 *
 * @param {Object} template       From readOrderFormTemplate() with quantities stamped
 * @param {Array}  itemizedDetail From aggregateOrderByRef()
 * @param {Object} orderData      Order metadata
 * @param {string} formattedDate  Pre-formatted date string
 * @returns {string}              Full HTML document
 */
function _buildOrderFormHtml(template, itemizedDetail, orderData, formattedDate) {
    const clientName = orderData.clientName || 'Unknown Client';
    const salesRep = orderData.salesRep || '';
    const displayRep = String(salesRep).split(/\s+/)[0] || salesRep;

    // ── PAGE 1: Summary table rows + running grand total ────────────────────
    let page1Rows = '';
    let grandTotal = 0;

    template.rows.forEach(row => {
        if (row.type === 'category') {
            const bg = ORDER_FORM_COLORS.categoryBg;
            const tc = ORDER_FORM_COLORS.categoryText;
            const bdr = ORDER_FORM_COLORS.accentBorder;
            // Use bgcolor attribute — Google's blob PDF converter ignores CSS background-color
            page1Rows += `
        <tr bgcolor="${bg}">
          <td colspan="8" bgcolor="${bg}"
              style="color:${tc};font-weight:700;font-size:11px;
                     padding:5px 8px;text-align:center;letter-spacing:0.3px;
                     border-top:2px solid ${bdr};border-bottom:2px solid ${bdr};">
            ${_esc(row.label)}
          </td>
        </tr>`;

        } else if (row.type === 'product') {
            const hasAny = (row.singlesQty > 0 || row.mcQty > 0);
            const sQty = row.singlesQty > 0 ? row.singlesQty : '';
            const mQty = row.mcQty > 0 ? row.mcQty : '';
            const totalAmt = row.totalPrice || 0;
            if (totalAmt > 0) grandTotal += totalAmt;
            const tDisplay = totalAmt > 0 ? '$' + totalAmt.toFixed(2) : '';

            const sStyle = row.singlesQty > 0 ? 'font-weight:700;' : '';
            const mStyle = row.mcQty > 0 ? 'font-weight:700;' : '';
            const tStyle = totalAmt > 0 ? `color:${ORDER_FORM_COLORS.totalText};font-weight:700;` : '';

            // MC price cell — bgcolor attribute for yellow (CSS background ignored by blob PDF)
            const mcPriceVal = row.mcPrice || '';
            const mcBgAttr = mcPriceVal ? ` bgcolor="${ORDER_FORM_COLORS.mcPriceBg}"` : '';
            const mcExtraStyle = mcPriceVal ? 'font-weight:700;' : '';

            page1Rows += `
        <tr>
          <td style="text-align:center;font-size:9px;color:#555;padding:3px 4px;border:1px solid #ccc;width:36px;">${_esc(row.ref)}</td>
          <td style="font-weight:700;padding:3px 6px;border:1px solid #ccc;">${_esc(row.label)}</td>
          <td style="font-style:italic;text-align:right;font-size:10px;color:#444;padding:3px 6px;border:1px solid #ccc;width:90px;">${_esc(row.format)}</td>
          <td style="text-align:right;font-size:10px;padding:3px 6px;border:1px solid #ccc;width:72px;">${_esc(row.price)}</td>
          <td${mcBgAttr} style="text-align:center;font-size:10px;padding:3px 6px;border:1px solid #ccc;width:110px;${mcExtraStyle}">${_esc(mcPriceVal)}</td>
          <td style="text-align:center;padding:3px 6px;border:1px solid #ccc;width:55px;${sStyle}">${sQty}</td>
          <td style="text-align:center;padding:3px 6px;border:1px solid #ccc;width:55px;${mStyle}">${mQty}</td>
          <td style="text-align:right;padding:3px 6px;border:1px solid #ccc;width:72px;${tStyle}">${tDisplay}</td>
        </tr>`;

        } else if (row.type === 'blank') {
            page1Rows += `<tr><td colspan="8" style="height:3px;border:none;background:#fff;"></td></tr>`;
        }
    });

    // Grand total row
    const grandTotalDisplay = grandTotal > 0 ? '$' + grandTotal.toFixed(2) : '';
    const grandTotalRow = `
        <tr style="border-top:2px solid #b050b0;">
          <td colspan="7" style="text-align:right;font-weight:700;font-size:11px;
                 padding:6px 8px;border:1px solid #ccc;background:#f8f8f8;
                 letter-spacing:0.3px;text-transform:uppercase;color:#333;">
            Order Total
          </td>
          <td style="text-align:right;font-weight:800;font-size:12px;
                 padding:6px 8px;border:1px solid #ccc;background:#e8f5e9;color:#1a6b2a;">
            ${grandTotalDisplay}
          </td>
        </tr>`;

    // ── PAGE 2: Itemized two-column detail ───────────────────────────────────
    const ordered = itemizedDetail
        .filter(i => i.quantity > 0)
        .sort((a, b) => a.ref.localeCompare(b.ref) || (a.variation || a.name).localeCompare(b.variation || b.name));

    const half = Math.ceil(ordered.length / 2);
    const leftCol = ordered.slice(0, half);
    const rightCol = ordered.slice(half);
    const maxRows = Math.max(leftCol.length, rightCol.length);

    let detailRows = '';
    for (let i = 0; i < maxRows; i++) {
        const L = leftCol[i];
        const R = rightCol[i];
        const strip = i % 2 === 0 ? '#ffffff' : '#f5f5f5';
        detailRows += `
      <tr style="background:${strip};">
        <td style="text-align:center;font-weight:700;font-size:8.5px;color:#888;width:32px;border:1px solid #ddd;padding:3px 4px;">${L ? _esc(L.ref) : ''}</td>
        <td style="padding:3px 7px;border:1px solid #ddd;font-size:10px;">${L ? _esc(L.variation || L.name) : ''}</td>
        <td style="text-align:center;font-weight:700;width:44px;border:1px solid #ddd;padding:3px 4px;">${L ? L.quantity : ''}</td>
        <td style="width:10px;background:#ddd;border:none;"></td>
        <td style="text-align:center;font-weight:700;font-size:8.5px;color:#888;width:32px;border:1px solid #ddd;padding:3px 4px;">${R ? _esc(R.ref) : ''}</td>
        <td style="padding:3px 7px;border:1px solid #ddd;font-size:10px;">${R ? _esc(R.variation || R.name) : ''}</td>
        <td style="text-align:center;font-weight:700;width:44px;border:1px solid #ddd;padding:3px 4px;">${R ? R.quantity : ''}</td>
      </tr>`;
    }

    // ── Full HTML document ────────────────────────────────────────────────────
    return `<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:Arial,Helvetica,sans-serif; font-size:11px; color:#1a1a1a; background:#fff; }
  .page { padding:18px 22px; max-width:800px; margin:0 auto; }
  .page-break { page-break-before:always; }

  /* ── Page 1 header ── */
  .p1-hdr  { display:flex; justify-content:space-between; align-items:baseline;
             border-bottom:2px solid #cc66cc; padding-bottom:7px; margin-bottom:9px; }
  .p1-name { font-size:17px; font-weight:800; }
  .p1-meta { font-size:10px; color:#555; text-align:right; line-height:1.7; }
  .p1-meta strong { color:#111; }

  /* Date / Store / Shipping bar — matches sheet top row */
  .p1-bar  { display:flex; gap:40px; font-size:11px; font-weight:700;
             margin-bottom:11px; }
  .p1-bar span { display:flex; gap:6px; align-items:center; }
  .p1-bar span em { font-weight:400; font-style:normal; min-width:80px;
                    border-bottom:1px solid #bbb; display:inline-block; }

  /* ── Order table ── */
  .ot { width:100%; border-collapse:collapse; font-size:10.5px; }
  .ot th { font-size:10px; font-weight:700; padding:5px 6px;
           background:${ORDER_FORM_COLORS.categoryBg}; color:${ORDER_FORM_COLORS.categoryText};
           border:1px solid ${ORDER_FORM_COLORS.accentBorder};
           text-transform:uppercase; letter-spacing:0.3px; }
  .ot th.la { text-align:left; }
  .ot th.ca { text-align:center; }
  .ot th.ra { text-align:right; }

  /* ── Grand total row ── */
  .grand-total { margin-top:0; }

  /* ── Page 2 ── */
  .p2-hdr { font-size:14px; font-weight:800; color:${ORDER_FORM_COLORS.categoryBg};
             border-bottom:2px solid ${ORDER_FORM_COLORS.accentBorder}; padding-bottom:5px; margin-bottom:6px; }
  .p2-sub { font-size:10px; color:#555; margin-bottom:11px; }
  .dt { width:100%; border-collapse:collapse; font-size:10px; }
  .dt th { background:${ORDER_FORM_COLORS.categoryBg}; color:${ORDER_FORM_COLORS.categoryText}; font-weight:700; font-size:9px;
           padding:4px 7px; text-transform:uppercase; letter-spacing:0.3px;
           border:1px solid ${ORDER_FORM_COLORS.accentBorder}; }
  .dt th.ca { text-align:center; }
</style>
</head>
<body>

<!-- ═══ PAGE 1: ORDER FORM SUMMARY ═══ -->
<div class="page">

  <div class="p1-hdr">
    <div class="p1-name">${_esc(clientName)}</div>
    <div class="p1-meta">
      Date: <strong>${formattedDate}</strong><br>
      Order: <strong>#${_esc(String(orderData.id || 'N/A'))}</strong><br>
      Rep: <strong>${_esc(displayRep)}</strong>
    </div>
  </div>

  <!-- Date / Store / Shipping row matching the template -->
  <div class="p1-bar">
    <span><strong>Date:</strong> <em>${formattedDate}</em></span>
    <span><strong>Store:</strong> <em>${_esc(clientName)}</em></span>
    <span><strong>Shipping:</strong> <em>${_esc(orderData.shipping || '')}</em></span>
  </div>

  <table class="ot">
    <thead>
      <tr>
        <th bgcolor="#cc66cc" class="ca" style="width:36px;border:1px solid #b050b0;color:#fff;"></th>
        <th bgcolor="#cc66cc" class="la" style="width:175px;border:1px solid #b050b0;color:#fff;"></th>
        <th bgcolor="#cc66cc" class="ra" style="width:90px;border:1px solid #b050b0;color:#fff;"></th>
        <th bgcolor="#cc66cc" colspan="2" class="ca" style="border:1px solid #b050b0;color:#fff;width:182px;">Price / Unit</th>
        <th bgcolor="#cc66cc" class="ca" style="width:55px;border:1px solid #b050b0;color:#fff;">Singles</th>
        <th bgcolor="#cc66cc" class="ca" style="width:55px;border:1px solid #b050b0;color:#fff;">MC</th>
        <th bgcolor="#cc66cc" class="ra" style="width:72px;border:1px solid #b050b0;color:#fff;">Total</th>
      </tr>
    </thead>
    <tbody>
      ${page1Rows}
    </tbody>
  </table>

  <!-- Total: sits below the table just like in the sheet template -->
  <div style="text-align:right;margin-top:6px;font-size:11px;font-weight:700;color:#1a1a1a;">
    Total:&nbsp;&nbsp;
    <span style="font-size:13px;font-weight:800;color:${ORDER_FORM_COLORS.totalText};min-width:80px;display:inline-block;text-align:right;">
      ${grandTotalDisplay}
    </span>
  </div>

</div>

<!-- ═══ PAGE 2: ITEMIZED DETAIL ═══ -->
<div class="page page-break">
  <div class="p2-hdr">Order Detail — ${_esc(clientName)}</div>
  <div class="p2-sub">
    Date: ${formattedDate} &nbsp;|&nbsp;
    Order #${_esc(String(orderData.id || 'N/A'))} &nbsp;|&nbsp;
    Rep: ${_esc(displayRep)}
  </div>

  <table class="dt">
    <thead>
      <tr>
        <th bgcolor="#cc66cc" class="ca" style="width:32px;color:#fff;">Ref</th>
        <th bgcolor="#cc66cc" style="text-align:left;color:#fff;">Flavour / Strain</th>
        <th bgcolor="#cc66cc" class="ca" style="width:46px;color:#fff;">Qty</th>
        <th bgcolor="#cc66cc" style="width:10px;border:none;"></th>
        <th bgcolor="#cc66cc" class="ca" style="width:32px;color:#fff;">Ref</th>
        <th bgcolor="#cc66cc" style="text-align:left;color:#fff;">Flavour / Strain</th>
        <th class="ca" style="width:46px;">Qty</th>
      </tr>
    </thead>
    <tbody>${detailRows}</tbody>
  </table>
</div>

</body></html>`;
}

// ============================================================================
// PRIVATE HELPERS
// ============================================================================

/** HTML-escape a string */
function _esc(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Return black or white depending on which has better contrast against bgHex.
 * Uses the same YIQ formula as getContrastYIQ() in ProductService.
 */
function _contrastColor(hex) {
    try {
        let h = String(hex || '').replace('#', '');
        if (h.length === 3) h = h.split('').map(c => c + c).join('');
        const r = parseInt(h.substr(0, 2), 16);
        const g = parseInt(h.substr(2, 2), 16);
        const b = parseInt(h.substr(4, 2), 16);
        return ((r * 299 + g * 587 + b * 114) / 1000) >= 128 ? '#000000' : '#ffffff';
    } catch (e) { return '#ffffff'; }
}

/**
 * Return a slightly darker shade of a hex colour (for borders on category rows).
 */
function _darken(hex) {
    try {
        let h = String(hex || '').replace('#', '');
        if (h.length === 3) h = h.split('').map(c => c + c).join('');
        const factor = 0.75;
        const r = Math.floor(parseInt(h.substr(0, 2), 16) * factor);
        const g = Math.floor(parseInt(h.substr(2, 2), 16) * factor);
        const b = Math.floor(parseInt(h.substr(4, 2), 16) * factor);
        return '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');
    } catch (e) { return '#999999'; }
}

// ============================================================================
// ORDER FORM TEMPLATE ROW MANAGEMENT
// ============================================================================

/**
 * Add a product row to the ORDER_FORM template sheet when a new product is created.
 *
 * Rules:
 *  - If the REF already exists in the sheet, skip (no duplicate).
 *  - Otherwise insert a new row just above the Shipping row (last pink footer row).
 *  - Populates: REF col, Name col, packaging col, Price/Unit col.
 *
 * @param {Object} product   ProductModel object with ref, name, category, price, etc.
 * @param {string} sheetName The ORDER_FORM sheet name (from getOrderFormSheetName)
 */
function addProductToOrderFormSheet(product, sheetName) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            Logger.log('addProductToOrderFormSheet: sheet not found: ' + sheetName);
            return;
        }

        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow < 1 || lastCol < 1) return;

        const allVals = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        const ref = String(product.ref || '').trim().toUpperCase();

        // Skip if REF already appears in column A (or B) of any product-like row
        for (let r = 0; r < allVals.length; r++) {
            const cellA = String(allVals[r][0] || '').trim().toUpperCase();
            const cellB = String(allVals[r][1] || '').trim().toUpperCase();
            if ((cellA === ref || cellB === ref) && ref) {
                Logger.log('addProductToOrderFormSheet: REF ' + ref + ' already in ' + sheetName + ' — skipping');
                return;
            }
        }

        // Find the "Shipping" row (last coloured footer row before the secondary table)
        let insertBeforeRow = lastRow; // default: append at end
        for (let r = allVals.length - 1; r >= 0; r--) {
            const rowText = allVals[r].join(' ').toLowerCase();
            if (/shipping/.test(rowText) && r > 2) {
                insertBeforeRow = r + 1; // 1-based row number of Shipping row
                break;
            }
        }

        // Insert a blank row just above Shipping
        sheet.insertRowBefore(insertBeforeRow);

        // Build the new row: detect which columns hold REF, name, price
        // by reading the template structure (assumes cols: A=REF, B=Name, rest dynamic)
        const newRow = sheet.getRange(insertBeforeRow, 1, 1, lastCol);
        const vals = new Array(lastCol).fill('');
        vals[0] = ref;                                            // col A = REF
        vals[1] = String(product.name || '').trim();              // col B = Name

        // Try to put packaging info in col C if present
        const packing = [product.variation3, product.variation4]
            .filter(v => v && String(v).trim() && String(v).trim() !== '1')
            .join(' ');
        if (packing && lastCol >= 3) vals[2] = packing;          // col C = Packaging

        // Price into first "Price/Unit"-like column after col C (col D if present)
        const price = parseFloat(product.price) || 0;
        if (price > 0 && lastCol >= 4) vals[3] = '$' + price.toFixed(2); // col D = Price/Unit

        newRow.setValues([vals]);

        // Copy row formatting from the row above (inherits borders / font)
        if (insertBeforeRow > 1) {
            sheet.getRange(insertBeforeRow - 1, 1, 1, lastCol)
                .copyFormatToRange(sheet, 1, lastCol, insertBeforeRow, insertBeforeRow);
        }

        Logger.log('addProductToOrderFormSheet: inserted ' + ref + ' into ' + sheetName + ' at row ' + insertBeforeRow);
    } catch (e) {
        Logger.log('addProductToOrderFormSheet ERROR: ' + e.message);
        // Do NOT re-throw — product save must not fail because of this
    }
}
