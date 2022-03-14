function EzraToSocioAtlas() {

  var file = DriveApp.getFileById("1aVmU8V2QAwI_YB1i3g3Ewe6lrNHXNX1E");
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  csvData.shift();  

  var SocioAtlasSheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sh = SpreadsheetApp.openById("1lFtV9bUkJLMI2aI_q9l-DbW7dqxadSvOgV3ikAAGh-c"); 
  SpreadsheetApp.setActiveSpreadsheet(sh);
  var ss = SpreadsheetApp.setActiveSheet(sh.getSheetByName('BLivre'));

  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  var Bible = ss.getDataRange().getValues();// get all Biblical text 
  var SQL = new gSQL();

  var BibleVerses = []; 


  for (let zz in csvData) {

    //Logger.log(csvData[zz][0].toLocaleUpperCase().substr(0,3));

    var BibleVersesAux = "SELECT Scripture FROM ? WHERE Book = '" + csvData[zz][0].toLocaleUpperCase().substr(0,3) + "' AND  Chapter = '" + csvData[zz][1] + "' AND Verse = '" + csvData[zz][2] + "'"; 
    BibleVerses = SUPERSQL(BibleVersesAux,Bible);

    Logger.log(BibleVersesAux);
    Logger.log(BibleVerses[1]);
  }

  SpreadsheetApp.setActiveSpreadsheet(SocioAtlasSheet);


}