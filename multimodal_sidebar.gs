/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function doMultimodal() {
  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  SpreadsheetApp.setActiveSheet(sh.getSheetByName('Matriz'));
  var ui = HtmlService.createTemplateFromFile('multimodal_sidebar_html')
      .evaluate()
      .setTitle('Análise multimodal');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Returns the active row.
 *
 * @return {Object[]} The headers & values of all cells in row.
 * Adapted from: https://stackoverflow.com/questions/30628894/how-do-i-make-a-sidebar-display-values-from-cells#30634581
 */
function getRecord() {
  // Retrieve and return the information requested by the sidebar.
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rowNum = sheet.getActiveCell().getRow();
  if (rowNum > data.length) return [];
  var record = [];
  for (var col=0;col<headers.length;col++) {
    var cellval = data[rowNum-1][col];
    // Dates must be passed as strings - use a fixed format for now
    if (typeof cellval == "object") {
      cellval = Utilities.formatDate(cellval, Session.getScriptTimeZone() , "M/d/yyyy");
    }
    // TODO: Format all cell values using SheetConverter library
    record.push({ heading: headers[col],cellval:cellval });
  }
  return record;
}