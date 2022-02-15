function doTable() {
  var html = HtmlService.createHtmlOutputFromFile('table')
  .setWidth(1400)
  .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Análise simples');
}

function scan_Table_data() {
  // Open sheet "Dados"
  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  var ss = sh.getSheetByName('Dados');
  var data1=ss.getDataRange().getValues();// get all data from Dados 

  // Retorna todos os itens com georeferenciamento
  var itensTable = SUPERSQL("Select Referencia, TipoReferencia, Destaque, FatorAtual, Significancia, FOFA, Aspecto, FOFA+'='+Significancia as Valor From ?",data1);

  itensTable.shift(); // remove header

    //Logger.log(itensTable);
  
  return itensTable;
}