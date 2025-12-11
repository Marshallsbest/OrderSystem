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
  
  // Return the evaluated HTML with mobile-friendly meta tags
  return template.evaluate()
      .setTitle(APP_TITLE)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Add Custom Menu to Spreadsheet
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu(APP_TITLE)
      .addItem('Get Order Form Web Link', 'showWebLink')
      .addToUi();
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
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
