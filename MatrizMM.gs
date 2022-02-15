function GeraMatrizMM(){

  // Cria validação de dados para coluna VINCULO
  //Source: https://stackoverflow.com/questions/58746916/google-sheets-data-validation-drop-down-menu-show-values-but-paste-formula

  var setas = new Array();
  setas[0] = '...';
  setas[1] = '⇦';
  setas[2] = '⬅⇨';
  setas[3] = '⬅';
  setas[4] = '⬅➡';
  setas[5] = '⇦➡';
  setas[6] = '⇨';
  setas[7] = '➡';
  setas[8] = '⇦⇨';

  // Get all data from Matriz
  var sh = SpreadsheetApp.getActiveSpreadsheet();
  var ssMatriz = sh.getSheetByName('Matriz');
  var MultimodalMatrix = ssMatriz.getDataRange().getValues();
  MultimodalMatrix.shift(); // remove header


  // Get all data from Dados
  var ss = sh.getSheetByName('Dados');
  var data1=ss.getDataRange().getValues();

  var matrizMM = []; 

  // Retorna todos os itens juntamente com o número de fatores em cada
  var fatoresMM = SUPERSQL("Select Comentario, count(Comentario) As ItemCount From ? GROUP BY Comentario",data1);
  fatoresMM.shift(); // remove header
  
  //Logger.log(fatoresMM);

  SpreadsheetApp.setActiveSheet(sh.getSheetByName('Matriz'));

  for (let i in fatoresMM) {
    // Considera apenas comentários com mais de um fator

    if (fatoresMM[i][1] > 1) {

      var matrizMM_aux = "SELECT a.Comentario AS aComentario, a.Aspecto AS aAspecto, a.FatorAtual AS aFator, b.FatorAtual AS bFator, b.Aspecto AS bAspecto, b.Comentario AS bComentario FROM ? a JOIN ? b ON a.Aspecto <= b.Aspecto WHERE (a.Aspecto <> '' AND b.Aspecto <> '') AND (a.FatorAtual <> b.FatorAtual) AND (a.Comentario = '" + fatoresMM[i][0] + "') AND (b.Comentario = '" + fatoresMM[i][0] + "') ORDER BY a.Aspecto"; 
      
      matrizMM = SUPERSQL(matrizMM_aux,data1,data1);

      if (matrizMM !== null && matrizMM !== undefined){
        matrizMM.shift();
        for (let m in matrizMM){

          var aModalidade_aux = matrizMM[m][1]; 
          var aFator_aux = matrizMM[m][2];
          var bFator_aux = matrizMM[m][3]; 
          var bModalidade_aux = matrizMM[m][4];
          

          // Verifica se dados já existem na tabela (array) Matriz
          var testSQL = "SELECT ID FROM ? WHERE (Fator_Normativo = '" + aFator_aux + "') AND (Fator_Determinativo ='" + bFator_aux + "')";
          var testData = SUPERSQL(testSQL,MultimodalMatrix);

          //Logger.log(testSQL);
          //Logger.log(testData);


          if (testData == null){

            // Add only new links and ignore similar combinations like A-B and B-A if the table Matrix already has one
            var testSQL_contrary = "SELECT ID FROM ? WHERE (Fator_Normativo = '" + bFator_aux + "') AND (Fator_Determinativo ='" + aFator_aux + "')";
            var testData_contrary = SUPERSQL(testSQL_contrary,MultimodalMatrix);

            if (testData_contrary == null){

              var insertTest = [Utilities.getUuid(),aModalidade_aux,aFator_aux,'','...','',bFator_aux,bModalidade_aux,'*'];
              MultimodalMatrix.push(insertTest);

            }

          } else {

            // If data exists on matrix array update only modalities 
            for( var ii = 0, len = MultimodalMatrix.length; ii < len; ii++ ) {
              if( MultimodalMatrix[ii][0] === testData ) {

                MultimodalMatrix[ii][1] = aModalidade_aux;
                MultimodalMatrix[ii][7] = bModalidade_aux;
                MultimodalMatrix[ii][8] = '*';

                break;
              }
            }

          }
        }
      }
    }
  }

  // Delete all rows that where not updated (it means that they are not exist anymore)
  var Updated_MultimodalMatrix = [];
  for( var n = 0, len = MultimodalMatrix.length; n < len; n++ ) {
    if( MultimodalMatrix[n][8] === '*' ) {
      var updatedValue = [MultimodalMatrix[n][0],MultimodalMatrix[n][1],MultimodalMatrix[n][2],MultimodalMatrix[n][3],MultimodalMatrix[n][4],MultimodalMatrix[n][5],MultimodalMatrix[n][6],MultimodalMatrix[n][7],''];
      Updated_MultimodalMatrix.push(updatedValue);
    }
  }

  //Logger.log(Updated_MultimodalMatrix.length);
  //Logger.log(Updated_MultimodalMatrix[0].length);
  //To understand how range works: 
  //https://stackoverflow.com/questions/11947590/sheet-getrange1-1-1-12-what-does-the-numbers-in-bracket-specify#11947798

  ssMatriz.getRange(2,1,Updated_MultimodalMatrix.length,Updated_MultimodalMatrix[0].length).setValues(Updated_MultimodalMatrix);

  // Sort rows by Normative Modality (1st) and Normative Factor (2nd) - false = descending / true = ascending
  ssMatriz.sort(3,true).sort(2,false);

  
  // Set data validation for populated columns only
  ssMatriz.getRange("E2:E").setDataValidation(null);
  var countVinculos = ssMatriz.getRange("E2:E").getValues();
  var count = countVinculos.filter(String).length;
  count++;
  var vinculos = ssMatriz.getRange("E2:E" + count);
  var rule = SpreadsheetApp.newDataValidation()
             .requireValueInList(setas)
             .setAllowInvalid(false)
             .build();

  vinculos.setDataValidation(rule);
  
 }