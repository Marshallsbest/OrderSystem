/**
 * Code.gs - Main entry point for the Order System
 */

// --- Global Constants ---
const APP_TITLE = "Order System";

/**
 * Serve the Web App
 */
function doGet(e) {
  // Identify the client from URL parameters if present
  const clientId = e.parameter.clientId || '';

  // Create the template
  const template = HtmlService.createTemplateFromFile('index');

  // Pass variables to the template
  template.clientId = clientId;
  template.categorySettings = getCategorySettings(); // Fetch colors

  // Return the evaluated HTML with mobile-friendly meta tags
  return template.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * OnOpen Trigger - Add Menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Order System');

  // Add items
  menu.addItem('Setup / Refresh Sheets', 'setupSheets')
    .addItem('Force PDF Generation', 'createOrderPdf') // Manual Trigger if needed
    .addItem('Add New Product', 'showAddProductSidebar') // NEW
    .addItem('Duplicate Selected Order', 'duplicateSelectedOrder') // SPREADSHEET TOOL
    .addItem('Populate Staging Sheet (ORDER_DATA)', 'populateOrderDataStaging') // NEW STAGING TOOL
    .addItem('Refresh Category Colors', 'applyCategoryColorsVisuals')
    .addToUi();
}

/**
 * Show Add Product Sidebar
 */
function showAddProductSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('addProductSidebar')
    .setTitle('Add New Product')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Setup/Refresh All Sheets
 * Wrapper for SheetService setup functions
 */
function setupSheets() {
  setupSettingsSheet();
  setupOrderDataSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast("Sheets setup complete.");
}

/**
 * Show the Web App URL in a dialog
 */
function showWebLink() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('The Web App is not yet deployed. Please deploy it as a Web App first (Execute as Me, Access: Anyone/Anyone with Google Account).');
    return;
  }

  const htmlOutput = HtmlService.createHtmlOutput(`<p>Use this link for your clients:</p><p><a href="${url}" target="_blank">${url}</a></p>`)
    .setWidth(400)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Order Form Web Link');
}

/**
 * Include partial HTML files (for css.html and js.html)
 */
function debugClientLookup(id) {
  return SheetService.debugClientLookup(id);
}

function debugGetSettingsData() {
  return SheetService.debugGetSettingsData();
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
