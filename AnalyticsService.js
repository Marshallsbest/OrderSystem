/**
 * AnalyticsService.gs
 * Handles data aggregation for the Daily Operations Dashboard
 */

function refreshDailyOperationsDashboard() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.DAILY_OPERATIONS);
    if (!sheet) {
        sheet = setupDailyOperationsSheet();
    }

    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const data = orderSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const now = new Date();
    const startOfWeek = new Date(now.setDate(now.getDate() - now.getDay()));
    startOfWeek.setHours(0, 0, 0, 0);

    // 1. Filter Orders Waiting for Processing (STATE: New)
    const pendingOrders = rows.filter(r => r[1] === "New");

    // 2. Filter Orders for Current Week
    const weekOrders = rows.filter(r => {
        const orderDate = new Date(r[2]); // Column C: TIMESTAMP
        return orderDate >= startOfWeek;
    });

    // 3. Aggregate Weekly Metrics
    let totalWeeklyComm = 0;
    const productStats = {}; // SKU -> Total Qty
    weekOrders.forEach(r => {
        totalWeeklyComm += parseFloat(r[7]) || 0; // Column H: COMMISSION

        // Parse items from Column K onwards (Index 10)
        // Format: [qty|@sku|$price|sale?]
        for (let i = 10; i < r.length; i++) {
            const itemStr = String(r[i]);
            const match = itemStr.match(/\[(\d+)\|@?([^\|]+)\|/);
            if (match) {
                const qty = parseInt(match[1]) || 0;
                const sku = match[2].trim();
                productStats[sku] = (productStats[sku] || 0) + qty;
            }
        }
    });

    // Find Most Popular Product
    let topProduct = "None";
    let maxQty = 0;
    for (const sku in productStats) {
        if (productStats[sku] > maxQty) {
            maxQty = productStats[sku];
            topProduct = sku;
        }
    }

    // 4. Breakdown By Customer (Time Groupings)
    const customerStats = {}; // Customer -> { week: {amt, count}, month: {amt, count}, quarter: {amt, count} }

    rows.forEach(r => {
        const customer = String(r[4] || "Unknown"); // Column E: CLIENT
        const amount = parseFloat(r[5]) || 0;       // Column F: TOTAL
        const oDate = new Date(r[2]);               // Column C: TIMESTAMP

        if (!customerStats[customer]) {
            customerStats[customer] = {
                week: { amt: 0, count: 0 },
                month: { amt: 0, count: 0 },
                quarter: { amt: 0, count: 0 }
            };
        }

        // Check date validity
        if (!oDate || isNaN(oDate.getTime())) return;

        // Check if in current week
        if (oDate >= startOfWeek) {
            customerStats[customer].week.amt += amount;
            customerStats[customer].week.count++;
        }

        // Check if in current month
        if (oDate.getMonth() === new Date().getMonth() && oDate.getFullYear() === new Date().getFullYear()) {
            customerStats[customer].month.amt += amount;
            customerStats[customer].month.count++;
        }

        // Check if in current quarter
        const curQuarter = Math.floor(new Date().getMonth() / 3);
        const orderQuarter = Math.floor(oDate.getMonth() / 3);
        if (curQuarter === orderQuarter && oDate.getFullYear() === new Date().getFullYear()) {
            customerStats[customer].quarter.amt += amount;
            customerStats[customer].quarter.count++;
        }
    });

    // 5. Render Output
    sheet.clear();

    // Header Style
    const titleRange = sheet.getRange("B2:G2");
    titleRange.merge().setValue("DAILY OPERATIONS & KPI DASHBOARD").setFontSize(16).setFontWeight("bold").setBackground("#0d2131").setFontColor("#ffffff").setHorizontalAlignment("left"); // Left aligned, deeper contrast

    // Column Group 1: Pending Orders
    sheet.getRange("B4").setValue("⏳ PENDING ORDERS (STATE: NEW)").setFontWeight("bold");
    const pendingData = pendingOrders.length > 0 ? pendingOrders.map(r => [r[0], r[4], r[5], new Date(r[2])]) : [["No pending orders", "", "", ""]];
    sheet.getRange(5, 2, 1, 4).setValues([["Invoice", "Customer", "Total", "Date"]]).setFontWeight("bold").setBackground("#f3f3f3");
    sheet.getRange(6, 2, pendingData.length, 4).setValues(pendingData);

    // Column Group 2: Weekly Summary
    const sumRow = 6 + pendingData.length + 2;
    sheet.getRange(sumRow, 2).setValue("📊 THIS WEEK'S SUMMARY").setFontWeight("bold");
    const summaryData = [
        ["Total Commission", "Top Product (SKU)", "Top Product Qty"],
        [totalWeeklyComm.toFixed(2), topProduct, maxQty]
    ];
    sheet.getRange(sumRow + 1, 2, 2, 3).setValues(summaryData).setBorder(true, true, true, true, true, true);
    sheet.getRange(sumRow + 1, 2, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");

    // Column Group 3: Customer Breakdown
    const breakRow = sumRow + 5;
    sheet.getRange(breakRow, 2).setValue("👥 CUSTOMER PERFORMANCE BREAKDOWN").setFontWeight("bold");
    const breakHeaders = [["Customer", "Wk Amt", "Wk #", "Mo Amt", "Mo #", "Qtr Amt", "Qtr #"]];
    const breakData = Object.keys(customerStats).map(c => [
        c,
        customerStats[c].week.amt.toFixed(2), customerStats[c].week.count,
        customerStats[c].month.amt.toFixed(2), customerStats[c].month.count,
        customerStats[c].quarter.amt.toFixed(2), customerStats[c].quarter.count
    ]);

    sheet.getRange(breakRow + 1, 2, 1, 7).setValues(breakHeaders).setFontWeight("bold").setBackground("#f3f3f3");
    if (breakData.length > 0) {
        sheet.getRange(breakRow + 2, 2, breakData.length, 7).setValues(breakData);
    }

    // Formatting
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 120);

    SpreadsheetApp.flush();
}
