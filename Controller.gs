/**
 * Controller.gs
 * Interface layer between the Spreadsheet UI/Web App and the Service logic
 */

/**
 * Main Web App Entry Point
 */
function doGet(e) {
    const clientId = e.parameter.clientId || '';
    const editOrderId = e.parameter.orderId || '';
    let prefillData = null;

    if (editOrderId) {
        prefillData = getOrderById(editOrderId);
    }

    const template = HtmlService.createTemplateFromFile('index');
    template.clientId = clientId;
    template.editOrderId = editOrderId;
    template.prefillData = prefillData;
    template.categorySettings = getCategorySettings();
    template.appConfig = getAppConfig();
    template.version = CURRENT_VERSION;

    return template.evaluate()
        .setTitle(APP_TITLE)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Spreadsheet Menu Trigger
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Order System');

    menu.addItem('📊 Admin Dashboard', 'setupDashboard')
        .addSeparator()
        .addItem('⚙️ Setup / Refresh Sheets', 'setupSheets')
        .addItem('📄 Generate PDF for Selection', 'generateSelectedOrderPdf')
        .addItem('📋 Populate Staging for Selection', 'populateStagingFromSelectedOrder')
        .addSeparator()
        .addItem('➕ Add New Product', 'showAddProductSidebar')
        .addItem('🛒 Open Order Form (Web Style)', 'showOrderFormDialog')
        .addSeparator()
        .addItem('📦 Update Export Summary', 'refreshExportSummary')
        .addItem('📈 Update Daily Ops Summary', 'refreshDailyOperationsDashboard')
        .addSeparator()
        .addItem('🧹 Cleanup & Deduplicate', 'cleanupProductSheet')
        .addItem('🌈 Style Product Headers', 'styleProductHeaders')
        .addToUi();
}

/**
 * UI Component Launchers
 */
function showAddProductSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('addProductSidebar')
        .setTitle('Add/Edit Product')
        .setWidth(450);
    SpreadsheetApp.getUi().showSidebar(html);
}

function showOrderFormDialog() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    const activeVal = String(activeCell.getValue()).trim();

    let editOrderId = '';
    let clientId = '';
    let prefillData = null;

    if (activeVal && (activeVal.startsWith('ORD-') || /^\d+$/.test(activeVal))) {
        try {
            const data = getOrderById(activeVal);
            if (data) {
                editOrderId = data.id;
                clientId = data.clientId;
                prefillData = data;
            }
        } catch (e) { }
    }

    const template = HtmlService.createTemplateFromFile('index');
    template.clientId = clientId;
    template.editOrderId = editOrderId;
    template.prefillData = prefillData;
    template.categorySettings = getCategorySettings();
    template.appConfig = getAppConfig();
    template.version = CURRENT_VERSION;

    const html = template.evaluate();
    html.setWidth(1200);
    html.setHeight(850);
    html.setTitle('Manual Order Entry');

    ui.showModalDialog(html, ' ');
}

/**
 * Install Dashboard Click Trigger
 * Run this once to enable dashboard button clicks
 */
function installDashboardTrigger() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Remove existing onSelectionChange triggers to avoid duplicates
    const triggers = ScriptApp.getUserTriggers(ss);
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onSelectionChange') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // Install new trigger
    ScriptApp.newTrigger('onSelectionChange')
        .forSpreadsheet(ss)
        .onSelectionChange()
        .create();

    SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard buttons are now active! Click any action in Column B.', 'Trigger Installed', 5);
}

/**
 * Client-Side API Wrappers (Called by google.script.run)
 */
function updateConfigSetting(key, value) {
    return Operations.updateConfigSetting(key, value);
}

function createExternalTemplates() {
    return Setup.createExternalTemplates();
}
