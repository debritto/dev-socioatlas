function doGet() {
  var html = HtmlService.createHtmlOutputFromFile('index')
  .setWidth(1400)
  .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gráfico Multimodal');
}
