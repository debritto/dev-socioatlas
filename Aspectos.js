/**
* Simple trigger that runs each time the user edits a cell in the spreadsheet.
*
* @param {Object} e The onOpen() event object.
* Adapted from: https://support.google.com/docs/thread/40589349/onedit-triggered-by-edit-in-one-specific-cell-not-any-cell-in-sheet?hl=en
* and from: https://stackoverflow.com/questions/12583187/google-spreadsheet-script-check-if-edited-cell-is-in-a-specific-range
*/

function onEdit(e) {
  // Run only if there's not a script running
  if (checkRunning() !== 'true'){
    if (!e) {
      throw new Error('Please do not run the script in the script editor window. It runs automatically when you edit the spreadsheet.');
    }
    // Verify sheet name
    var SheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

    if (SheetName == "Fatores"){
      var myRange = SpreadsheetApp.getActiveSheet().getRange('FatoresAspectos'); //<<< Change Your Named Ranged Name Here
      var myRangeSubst = SpreadsheetApp.getActiveSheet().getRange('FatoresSubstituicao').getValues(); // Column 'Substituição'
      var myRangeFator = SpreadsheetApp.getActiveSheet().getRange('FatoresFator').getValues(); // Column 'Substituição'


      //Let's get the row & column indexes of the active cell
      var row = e.range.getRow();
      var col = e.range.getColumn();

      //Mostra janela com valores da atual linha, para depuração de erros.
      //SpreadsheetApp.getUi().alert("Fator: " + myRangeFator[row-1].toString() + "  Substituicao: " + myRangeSubst[row-1].toString());

      //Check that your active cell is within your named range
      if (col >= myRange.getColumn() && col <= myRange.getLastColumn() && row >= myRange.getRow() && row <= myRange.getLastRow()) { //As defined by your Named Range

        //Update all factors with the same name automatically
        setScriptIsRunning(true); // Marca variável que diz que este script está rodando - com isso evita que triggers sejam disparados pela própria função OnEdit.

        var SQL = new gSQL();
        var AspectoSubstituicao = "";
        AspectoSubstituicao = e.value;

        var FatorSubstituicao = myRangeSubst[row-1].toString();
        var FatorSubstituicao_Aux = myRangeFator[row-1].toString();

        if ((FatorSubstituicao === "") || (FatorSubstituicao === null)){
          // Quando Substituicao = ""
          FatorSubstituicao = FatorSubstituicao_Aux;
          SQL.DB('').TABLE('Fatores').UPDATE('Aspecto').WHERE('Substituicao','=',FatorSubstituicao.toString()).VALUES(AspectoSubstituicao.toString()).setVal();
          SQL.DB('').TABLE('Fatores').UPDATE('Aspecto').WHERE('Fator','=',FatorSubstituicao.toString()).VALUES(AspectoSubstituicao.toString()).setVal();
          showMessage_("Todas as células com o fator '" + FatorSubstituicao + "' foram atualizadas com o aspecto '" + AspectoSubstituicao + "' !");
        
        } else {
          // Quando Substituicao <> ""
          SQL.DB('').TABLE('Fatores').UPDATE('Aspecto').WHERE('Substituicao','=',FatorSubstituicao.toString()).VALUES(AspectoSubstituicao.toString()).setVal();
          SQL.DB('').TABLE('Fatores').UPDATE('Aspecto').WHERE('Fator','=',FatorSubstituicao.toString()).AND('Substituicao','=','').VALUES(AspectoSubstituicao.toString()).setVal();
        showMessage_("Todas as células com o fator '" + FatorSubstituicao + "' foram atualizadas com o aspecto '" + AspectoSubstituicao + "' !");


        }
        setScriptIsRunning(false); // Marca variável que diz que este script está rodando - com isso evita que triggers sejam disparados.
      }
    }
    SpreadsheetApp.flush();
  }  
}

function test_db(){

  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  SpreadsheetApp.setActiveSheet(sh.getSheetByName('Fatores'));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data1=ss.getDataRange().getValues();// get all data from Dados 

  // Retorna todos os itens juntamente com o número de fatores em cada
  var retorno1 = alasql("SELECT * FROM ?",data1)
  var retorno2 = alasql("UPDATE Fatores SET Aspecto = '160-Ético' WHERE Substituicao = 'INTEGRAÇÃO LOCAL'",retorno1);

   
  Logger.log(retorno2);

}


function myupdate(){
  var SQL = new gSQL();

  var AspectoSubstituicao = "010-Numérico";
  //var AspectoSubstituicao = "000-INDEFINIDO";
  var FatorSubstituicao = 'INTEGRAÇÃO LOCAL';

  //Update the Lenon's age  
  SQL.DB('1UnUIeGpUR66xWlfe3l5x5AzZ1IKhQh70_awMiZihHeM').TABLE('Fatores').UPDATE('Aspecto').WHERE('Substituicao','=',FatorSubstituicao).VALUES(AspectoSubstituicao).setVal();
  SQL.DB('1UnUIeGpUR66xWlfe3l5x5AzZ1IKhQh70_awMiZihHeM').TABLE('Fatores').UPDATE('Aspecto').WHERE('Fator','=',FatorSubstituicao).AND('Substituicao','=','').VALUES(AspectoSubstituicao).setVal();
}

/**
* Shows a message in a pop-up.
*
* @param {String} message The message to show.
*/
function showMessage_(message, timeoutSeconds) {
  SpreadsheetApp.getActive().toast(message, 'Aviso:', timeoutSeconds || 5);
}