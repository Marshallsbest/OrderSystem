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
        .addItem('📄 Generate Invoice PDF for Selection', 'generateSelectedOrderPdf')
        .addItem('📋 Generate Order Form PDF for Selection', 'generateSelectedOrderFormPdf')
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
    const template = HtmlService.createTemplateFromFile('addProductSidebar');
    const html = template.evaluate()
        .setTitle('Add/Edit Product');
    SpreadsheetApp.getUi().showSidebar(html);
}

function showOrderFormDialog() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    const activeVal = String(activeCell.getValue()).trim();

    let editOrderId = '';
    let clientId = '';

    if (activeVal && (activeVal.startsWith('ORD-') || /^\d+$/.test(activeVal))) {
        try {
            const data = getOrderById(activeVal);
            if (data) {
                editOrderId = data.id;
                clientId = data.clientId;
            }
        } catch (e) { }
    }

    // Read WEB_APP_URL from SETTINGS
    let webAppUrl = '';
    try {
        const r = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('WEB_APP_URL');
        if (r) webAppUrl = String(r.getValue() || '').trim();
    } catch (e) { }

    if (!webAppUrl) {
        ui.alert('WEB_APP_URL is not set in SETTINGS.\n\nPlease deploy the web app and paste the URL into SETTINGS \u2192 WEB_APP_URL.');
        return;
    }

    // Append query params for client/order pre-fill
    const parts = [];
    if (clientId) parts.push('clientId=' + encodeURIComponent(clientId));
    if (editOrderId) parts.push('orderId=' + encodeURIComponent(editOrderId));
    const finalUrl = webAppUrl + (parts.length ? '?' + parts.join('&') : '');

    // A tiny launcher dialog opens a proper movable + resizable popup then closes itself.
    // NOTE: we do NOT call google.script.host.close() here — doing so kills the popup
    // because the popup is a child of the dialog's window context in Chrome.
    // The user closes this small launcher manually after the popup is open.
    const launcher = HtmlService.createHtmlOutput(
        '<html><head><script>' +
        'window.onload=function(){' +
        '  var sw=window.screen.availWidth, sh=window.screen.availHeight;' +
        '  var pw=Math.min(1200,Math.round(sw*0.85));' +
        '  var ph=Math.min(900, Math.round(sh*0.85));' +
        '  var lft=Math.round((sw-pw)/2), top=Math.round((sh-ph)/4);' +
        '  var popup=window.open("' + finalUrl + '","OrderForm",' +
        '    "width="+pw+",height="+ph+",left="+lft+",top="+top+' +
        '    ",resizable=yes,scrollbars=yes,toolbar=no,menubar=no,location=no,status=no");' +
        '  if(popup){' +
        '    document.getElementById("msg").textContent="\u2713 Order Form opened in a new window.";' +
        '    document.getElementById("msg").style.color="#188038";' +
        '  } else {' +
        '    document.getElementById("msg").innerHTML="Popup blocked. <a href=\\"' + finalUrl + '\\" target=\\"_blank\\">Click here</a> to open.";' +
        '    document.getElementById("msg").style.color="#c5221f";' +
        '  }' +
        '};' +
        '<\/script><\/head>' +
        '<body style="font-family:Roboto,sans-serif;padding:16px;text-align:center;">' +
        '<p id="msg" style="font-size:13px;margin:0 0 12px;">Opening Order Form\u2026<\/p>' +
        '<button onclick="google.script.host.close()" ' +
        '  style="background:#1a73e8;color:#fff;border:none;padding:8px 20px;' +
        '         border-radius:6px;cursor:pointer;font-size:13px;">Close<\/button>' +
        '<\/body><\/html>'
    ).setWidth(320).setHeight(110);

    ui.showModalDialog(launcher, 'Order System');
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
