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
    template.appStyles = getAppStyles();
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

    menu.addItem('🛒 Open Order Form (Web App)', 'showOrderFormDialog')
        .addItem('➕ Add New Product', 'showAddProductSidebar')
        .addSeparator()
        .addItem('📄 Generate PDF for Selection', 'generateSelectedOrderPdf')
        .addItem('📂 Open PDF Exports Folder', 'openExportFolder')
        .addSeparator()
        .addSubMenu(ui.createMenu('🛡️ Header Protection')
            .addItem('📸 Backup Sheet Headers', 'backupSheetHeaders')
            .addItem('🔍 Compare Headers to Backup', 'compareSheetHeaders')
            .addItem('♻️ Reset Headers from Backup', 'resetSheetHeaders')
        )
        .addSubMenu(ui.createMenu('📋 Deployment')
            .addItem('🔗 Get Copy Link', 'showCopyLink')
            .addItem('📦 Create New Copy for Colleague', 'createCleanCopy')
            .addItem('🔄 Check for Updates', 'checkForUpdates')
            .addItem('📡 Register as Master', 'registerAsMaster')
            .addSeparator()
            .addItem('⬇️ Pull Updates from Master', 'pullUpdatesFromMaster')
        )
        .addSubMenu(ui.createMenu('🔧 System')
            .addItem('🏭 Factory Reset', 'factoryResetSpreadsheet')
        )
        .addToUi();

    // Auto-backup headers on first open (initialization lifecycle)
    initializeOnFirstOpen_();
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
    template.appStyles = getAppStyles();
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

/**
 * Show the "Copy this Spreadsheet" link in a dialog.
 * The link uses Google Sheets' built-in /copy URL pattern.
 * Anyone with access who clicks it will be prompted to make their own copy.
 */
function showCopyLink() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = ss.getId();
    const copyUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/copy';
    const qrUrl = 'https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=' + encodeURIComponent(copyUrl);

    const html = HtmlService.createHtmlOutput(
        '<div style="font-family: Roboto, sans-serif; padding: 20px;">' +
        '<h3 style="color: #006c4c; margin-top: 0;">📋 Copy Link</h3>' +
        '<p style="font-size: 13px; color: #666;">Share this link with a colleague. When they open it, Google will prompt them to make their own copy of this spreadsheet.</p>' +
        '<div style="background: #f5f5f5; padding: 12px; border-radius: 8px; border: 1px solid #e0e0e0; margin: 16px 0; word-break: break-all; font-family: Consolas, monospace; font-size: 12px;" id="linkBox">' +
        copyUrl +
        '</div>' +
        '<div style="display: flex; gap: 8px; margin-bottom: 20px;">' +
        '<button onclick="copyLink()" style="background: #006c4c; color: white; border: none; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px;" id="copyBtn">📋 Copy Link</button>' +
        '<a href="' + copyUrl + '" target="_blank" style="background: #1a73e8; color: white; padding: 10px 20px; border-radius: 8px; text-decoration: none; font-size: 14px;">🔗 Open</a>' +
        '</div>' +
        '<hr style="margin: 16px 0; border: none; border-top: 1px solid #e0e0e0;">' +
        '<p style="font-size: 12px; color: #666; margin-bottom: 8px;">QR Code (scan to copy):</p>' +
        '<img src="' + qrUrl + '" width="200" height="200" style="border: 1px solid #eee; border-radius: 8px;">' +
        '<script>' +
        'function copyLink() {' +
        '  var text = "' + copyUrl + '";' +
        '  navigator.clipboard.writeText(text).then(function() {' +
        '    document.getElementById("copyBtn").textContent = "✅ Copied!";' +
        '    setTimeout(function() { document.getElementById("copyBtn").textContent = "📋 Copy Link"; }, 2000);' +
        '  }).catch(function() {' +
        '    var ta = document.createElement("textarea");' +
        '    ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);' +
        '    document.getElementById("copyBtn").textContent = "✅ Copied!";' +
        '    setTimeout(function() { document.getElementById("copyBtn").textContent = "📋 Copy Link"; }, 2000);' +
        '  });' +
        '}' +
        '</script>' +
        '</div>'
    ).setWidth(480).setHeight(480);

    SpreadsheetApp.getUi().showModalDialog(html, 'Share Spreadsheet Copy Link');
}
