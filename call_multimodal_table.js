function doMultimodalTable() {
  var html = HtmlService.createHtmlOutputFromFile('multimodal_table')
  .setWidth(1400)
  .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Análise multimodal');
}

function scan_Multimodal_data() {
  // Get active cell values for Fator_Normativo and Fator_Determinativo
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getCurrentCell().getRow();
  var rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Open sheet "Dados" and retrieve data
  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  var ss = sh.getSheetByName('Dados');
  var data1 = ss.getDataRange().getValues();// get all data from Dados 

  // Retorna todos os itens com georeferenciamento
  var sql_query = "Select Referencia, TipoReferencia, Texto, Destaque, FatorAtual, Aspecto From ? Where (Referencia <> 'IMG') and (FatorAtual = '" + rowValues[2] + "' or FatorAtual = '" + rowValues[6] + "')";

  //var ui = SpreadsheetApp.getUi(); 
  //ui.alert(sql_query);

  var itensTable = SUPERSQL(sql_query,data1);

  itensTable.shift(); // remove header

    //Logger.log(itensTable);
 
  return itensTable;
}
