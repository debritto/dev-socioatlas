function doMap() {
  var html = HtmlService.createHtmlOutputFromFile('mapping')
  .setWidth(1400)
  .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Georreferenciamento Multimodal');
}

function scan_GeoJSON_data() {

  // Open sheet "Dados"
  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  var ss = sh.getSheetByName('Dados');
  var data1=ss.getDataRange().getValues();// get all data from Dados 

  // Retorna todos os itens com georeferenciamento
  var itensGEO = SUPERSQL("Select Destaque, FatorAtual, FOFA, Aspecto, Longitude, Latitude, url From ? WHERE Longitude <> 0",data1);

  itensGEO.shift(); // remove header

  // Create an array of objects to record pre formatted GeoJSON data comming from multimodal factors
  var pre_GeoJSON_data = new Array();


  for (let i in itensGEO) {
    var imageURL = "https://drive.google.com/uc?export=view&id=" + getIdFromUrl(itensGEO[i][6]);
    var popupText = '<div style="text-align: center;"><img src = "' + imageURL + '" height = "40%" width = "40%"/><br><a href="' + itensGEO[i][6] + '" target="_blank">Editar Imagem</a>  </div><br>' + itensGEO[i][0] + '<br><b>' + itensGEO[i][1] + '</b> <i>' + itensGEO[i][3] + '</i><br>' + itensGEO[i][2];

    pre_GeoJSON_data.push({"popupContent": popupText,"lng":itensGEO[i][4],"lat": itensGEO[i][5]});
  }
  var GeoJSON_data = GeoJSON.parse(pre_GeoJSON_data, {Point: ['lat','lng'], include: ['popupContent']});

  //Logger.log(GeoJSON_data);

  return GeoJSON_data;
}
