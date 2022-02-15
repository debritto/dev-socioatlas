// Based on code: https://stackoverflow.com/questions/39016714/google-apps-script-google-docs-trigger-on-comment-add-or-change#39021121 

function ExtrairComentarios() {
 
  setScriptIsRunning(true); // Marca variável que diz que este script está rodando - com isso evita que triggers sejam disparados.
// veja aqui: https://stackoverflow.com/questions/18749591/encode-html-entities-in-javascript/39243641#39243641
// Parece que funciona melhor que esta opção
  // Function to decode and encode html entities
  var decodeHtmlEntity = function(str) {
    return str.replace(/&#(\d+);/g, function(match, dec) {
      return String.fromCharCode(dec);
    });
  };

  var encodeHtmlEntity = function(str) {
    var buf = [];
    for (var i=str.length-1;i>=0;i--) {
      buf.unshift(['&#', str[i].charCodeAt(), ';'].join(''));
    }
    return buf.join('');
  };
  // Source: https://gist.github.com/CatTail/4174511
  // var entity = '&#39640;&#32423;&#31243;&#24207;&#35774;&#35745;';
  // var str = '高级程序设计';
  // console.log(decodeHtmlEntity(entity) === str);
  // console.log(encodeHtmlEntity(str) === entity);  
 
  // Cria uma array para cada coluna

  var hList = [], cList = [], vFactor = [], vSWOT_Type = [], iSignificance = [], vAuthor = [], vDate = [], vIDcomment = [], vBibliografia = [], vReferencia = [], vBib = [], vTipoReferencia = [], vURL = [], vLongitude = [], vLatitude = [];

  // --------------------------------------------------------------------------- Cria matriz com dados da tabela fontes

  var sh_fontes = SpreadsheetApp.getActiveSpreadsheet();  
  var ss_fontes = sh_fontes.getSheetByName('Fontes');
  var dados_fontes = ss_fontes.getDataRange().getValues();

  var SQL = new gSQL();
  var fontes = new Array();

  for(nn_fontes in dados_fontes){
      var fontes_aux = dados_fontes[nn_fontes][0];
      // Read documents
      if ( fontes_aux.includes('document')){
        fontes.push( dados_fontes[nn_fontes] );
      }

      // Read files
      if ( fontes_aux.includes('file')){
        fontes.push( dados_fontes[nn_fontes] );
      }

      // Read folders with documents
      if ( fontes_aux.includes('folders')){
        var fontesEmPastas = listFolderContents( fontes_aux );
        for (let zz in fontesEmPastas){
          fontes.push(fontesEmPastas[zz]);
        }
      }
  }
  
  //fontes.shift(); // Remove header
  
  //Logger.log(fontes);

  for (let xx in fontes) {

    // Change docId into your document's ID - See below on how to
    var docId = getIdFromUrl(fontes[xx][0]);

    Logger.log(docId);

    // Discover mimeType: adapted from: https://www.labnol.org/code/19912-mime-types-google-drive
    var mimeType = DriveApp.getFileById(docId).getMimeType();
    var fileType = getFileType(mimeType);

    Logger.log( fileType );

    // Process Ezra Bible Notes
    if (fileType == "CSV"){

      var csvFile = DriveApp.getFileById(docId);
      var csvData = Utilities.parseCsv(csvFile.getBlob().getDataAsString());
      csvData.shift(); // Remove header

      // Read Bible and search for verses
      var sh = SpreadsheetApp.getActiveSpreadsheet();
      var ss = sh.getSheetByName('BLivre');
      var Bible = ss.getDataRange().getValues();// get all Biblical text 
      var SQL = new gSQL();
      var BibleVerses = [];
      var strFatores = []; 
  
      for (let zz in csvData) {

        //Logger.log(csvData[zz][0]);

        var BibleVersesAux = "SELECT Scripture FROM ? WHERE Book = '" + csvData[zz][0].toLocaleUpperCase().substr(0,3) + "' AND  Chapter = '" + csvData[zz][3] + "' AND Verse = '" + csvData[zz][4] + "'"; 


        BibleVerses = SUPERSQL(BibleVersesAux,Bible);

        //Logger.log(BibleVerses)
        
        if (BibleVerses && BibleVerses.length > 0) {

          var Scripture = csvData[zz][0].toLocaleUpperCase().substr(0,3) + " " + csvData[zz][3] + ":" + csvData[zz][4] + " " + BibleVerses[1]; 
          var ScriptureREF = csvData[zz][0].toLocaleUpperCase().substr(0,3) + " " + csvData[zz][3] + ":" + csvData[zz][4];
          var strFatores = csvData[zz][5].toLocaleUpperCase().split(';'); // cria uma array com os fatores 

          if (strFatores.length > 0){ // processa apenas itens com fatores
            for (let yy in strFatores) {

              hList.unshift([ Scripture  ]);
              cList.unshift([ csvData[zz][6] ]);
              if (strFatores[yy].length > 0){
                vFactor.unshift([ strFatores[yy] ]);
              } else {
                vFactor.unshift([ '>>INDEFINIDO<<' ]);
              }  
              vSWOT_Type.unshift(["N/A"]);
              iSignificance.unshift(["0"]);
              vAuthor.unshift([ csvFile.getOwner() ]);
              vDate.unshift([ csvFile.getDateCreated() ]);

              vIDcomment.unshift([ScriptureREF]);
              // Comentário tem que ser o versículo bíblico - para saber se há um ou mais fatores em um único versículo

              vBibliografia.unshift(["Bíblia Livre"]);
              vReferencia.unshift(["TEO"]);
              vTipoReferencia.unshift(["*"]);
              vBib.unshift([ csvData[zz][0].toLocaleUpperCase().substr(0,3) ]);
              vURL.unshift(["N/A"]);
              vLongitude.unshift(['']);
              vLatitude.unshift(['']);
            }
          }
        }
      }  

    } else if (fileType == "JPEG"){
      // Ref. file properties: https://developers.google.com/drive/api/v2/reference/files 

      var file = Drive.Files.get(docId);

      //
      var fileTitle = file.title;
      var fileData = fileTitle.split('_'); // cria uma array com fatores, comentários e significância do nome dor arquiv IMG 
      // Retira apenas comentário interno - que está no nome do arquivo.
      var Comment_Delimiter = fileData[2].indexOf("&")+1;
      if (Comment_Delimiter > 0){
        var Comment_Content = fileData[2].substr(Comment_Delimiter, fileData[2].length-1);
      } else {
        var Comment_Content = '';
      }  


      // Busca comentários feitos diretamente na foto dentro do  Google Drive  :) Youpyyy
      var comments = Drive.Comments.list(docId);
      var text = Comment_Content;
      if (comments.items && comments.items.length > 0) {
        for (var i = 0; i < comments.items.length; i++) {
          var comment = comments.items[i];
          if (text == ''){
            text = comment.content;
          } else {    
            text = text + ' | ' + comment.content;
          }  
        }
      }    


      var regex = /\#(\w+) /g;
      var AllFactors;
      while ( AllFactors = regex.exec(file.title) ) {
        var FactorAux = AllFactors[1];
        vFactor.unshift([FactorAux.toLocaleUpperCase().trim()]);
        var imgThumbnail = '=IMAGE("https://drive.google.com/uc?export=view&id=' + docId + '")';

        hList.unshift([ imgThumbnail ]);  // add image link
        cList.unshift([ text ]);  // add image note
        vAuthor.unshift([file.ownerNames]);
        vDate.unshift([ file.imageMediaMetadata.date ]);
        vIDcomment.unshift([docId]);
        vBibliografia.unshift(["Imagem"]);
        vReferencia.unshift(["IMG"]);
        vBib.unshift([""]);
        vURL.unshift([ fontes[xx][0] ]);

        // SWOT type
        var TagSWOT_Delimiter = fileTitle.indexOf("@")+1;
        var vSWOT = fileTitle.substr(TagSWOT_Delimiter, 2);
        var vSWOTref = fileTitle.substr(TagSWOT_Delimiter, 3);
        var vSignificance = vSWOTref.charAt(vSWOTref.length-1);

        if (vSignificance =="0" || vSignificance =="1" || vSignificance =="2") {
          iSignificance.unshift([vSignificance]);
        } else {
          iSignificance.unshift(["0"]);
        } 

        switch(vSWOT.toLowerCase())
        {
        case "fo":
          vSWOT_Type.unshift(["(INT) Força"]);
          vTipoReferencia.unshift(["+"]);
          break;
        case "fr":
          vSWOT_Type.unshift(["(INT) Fraqueza"]);
          vTipoReferencia.unshift(["-"]);
          break;
        case "op":
          vSWOT_Type.unshift(["(EXT) Oportunidade"]);
          vTipoReferencia.unshift(["+"]);
          break;
        case "am":
          vSWOT_Type.unshift(["(EXT) Ameaça"]);
          vTipoReferencia.unshift(["-"]);
          break;
        default:
          vSWOT_Type.unshift(["(*) INDEFINIDO"]);
          vTipoReferencia.unshift([""]);
          break;
        }

        // Test if 'longitude' property exists inside EXIF data
        if( testProperty(file, 'imageMediaMetadata.location.longitude') ){
          vLongitude.unshift([ file.imageMediaMetadata.location.longitude ]);
          vLatitude.unshift([ file.imageMediaMetadata.location.latitude ]);
        } else {
          vLongitude.unshift(['']);
          vLatitude.unshift(['']);
        }
      }

    } else if (fileType == "Google Docs"){ 

      var bibliografia = fontes[xx][1];
      var URL_aux = fontes[xx][0];
  
      //Logger.log( docId );

      var options = {
        'maxResults': 99  // caso precise baixar mais de 100 comentários ver aqui: https://stackoverflow.com/questions/47297128/how-to-specify-request-params-with-drive-comments-list e aqui https://developers.google.com/drive/api/v3/reference/comments/list
      };


      var comments = Drive.Comments.list(docId, options);
      // Get list of comments

      if (comments.items && comments.items.length > 0) {
        for (var i = 0; i < comments.items.length; i++) {
          var comment = comments.items[i];
          
          
          //Logger.log(comments.items[i].content);

          // Loop until identify all factors in the present comment 
          // Adapted from: https://stackoverflow.com/questions/1493027/javascript-return-string-between-square-brackets#1493071
          
          var text = comment.content;

          var regex = /\[([^\][]*)]/g;
          var AllFactors;
          while ( AllFactors = regex.exec(text) ) {
            var FactorAux = AllFactors[1];
            vFactor.unshift([FactorAux.toLocaleUpperCase().trim()]);
            //Logger.log(AllFactors[1]);


            //var TagDelimiter = comment.content.indexOf("]")+1;
            //var FactorAux = comment.content.substr(0, TagDelimiter);
            //vFactor.unshift([FactorAux.toLocaleUpperCase()]);


            // Documentação sobre comentários e textos marcados -> https://developers.google.com/drive/api/v2/reference/comments
            //Logger.log( comment.anchor.endsWith );

            if ( 'context' in comment ){    // Verifica se a propriedade context (texto marcado) está presente
              //hList.unshift([comment.context.value]);
              hList.unshift([decodeHtmlEntity(comment.context.value)]); 
            } else {
              hList.unshift(["Indefinido"]);
            }

            cList.unshift([comment.content]);  // add comment
            vAuthor.unshift([comment.author.displayName]);
            vDate.unshift([comment.modifiedDate]);
            vIDcomment.unshift([comment.commentId]);

            // SWOT type
            var TagSWOT_Delimiter = comment.content.indexOf("{")+1;
            var vSWOT = "";
            var vSWOT = comment.content.substr(TagSWOT_Delimiter, 2);
            var ref = 0;

            switch(vSWOT.toLowerCase())
            {
            case "fo":
              vSWOT_Type.unshift(["(INT) Força"]);
              vReferencia.unshift(["COM"]);
              vTipoReferencia.unshift(["+"]);
              vBib.unshift([""]);
              break;
            case "fr":
              vSWOT_Type.unshift(["(INT) Fraqueza"]);
              vReferencia.unshift(["COM"]);
              vTipoReferencia.unshift(["-"]);
              vBib.unshift([""]);
              break;
            case "op":
              vSWOT_Type.unshift(["(EXT) Oportunidade"]);
              vReferencia.unshift(["COM"]);
              vTipoReferencia.unshift(["+"]);
              vBib.unshift([""]);
              break;
            case "am":
              vSWOT_Type.unshift(["(EXT) Ameaça"]);
              vReferencia.unshift(["COM"]);
              vTipoReferencia.unshift(["-"]);
              vBib.unshift([""]);
              break;
            case "re":
              vSWOT_Type.unshift(["N/A"]);
              vReferencia.unshift(["REF"]);
              vBib.unshift([""]);
              ref = 1;
              break;
            case "te":
              vSWOT_Type.unshift(["N/A"]);
              vReferencia.unshift(["TEO"]);
              
              // Extract bible book name using regex
              // adapted from: https://stackoverflow.com/questions/22254746/bible-verse-regex
              const regex = /(?:\d\s*)?[A-Z]?[a-z]+\s*\d+(?:[:-]\d+)?(?:\s*-\s*\d+)?(?::\d+|(?:\s*[A-Z]?[a-z]+\s*\d+:\d+))?/i;
              var texto = hList[i];

              var BibAux = regex.exec( texto );
              // Como manipular objetos regex: https://stackoverflow.com/questions/432493/how-do-you-access-the-matched-groups-in-a-javascript-regular-expression
              // Eles retornam uma array
              //Logger.log(BibAux[0].substr(0,3));

              //Logger.log(BibAux);

              if (BibAux === null){
                vBib.unshift([ 'N/C' ]);
              } else {     
                vBib.unshift([ testBibleBookName(BibAux[0].substr(0,3)) ]);
              }
              ref = 1;
              break;            
            default:
              vSWOT_Type.unshift(["(*) INDEFINIDO"]);
              vReferencia.unshift([""]);
              vTipoReferencia.unshift([""]);
              vBib.unshift([""]);
              break;
            }

            if (ref == 0){
              // Significance
              var vSignificance = comment.content.substr(TagSWOT_Delimiter+2, 1);
              if (vSignificance =="0" || vSignificance =="1" || vSignificance =="2") {
                iSignificance.unshift([vSignificance]);
              } else {
                iSignificance.unshift(["0"]);
              } 
            } else {
              iSignificance.unshift(["0"]);
              // Ref Type
              var vTipoReferenciaAux = comment.content.substr(TagSWOT_Delimiter+3, 1);
              if (vTipoReferenciaAux =="-" || vTipoReferenciaAux =="+" || vTipoReferenciaAux =="*") {
                vTipoReferencia.unshift([vTipoReferenciaAux]);
              } else {
                vTipoReferencia.unshift(["*"]);
              }  
            }
            // Ref. Bibliografica
            vBibliografia.unshift([bibliografia]);

            // URL do texto vinculado à bibliografia
            vURL.unshift([URL_aux]);
            vLongitude.unshift(['']);
            vLatitude.unshift(['']);
          }
        }
      }
    }
  }
  
  // Sheet header
  hList.unshift([ "Texto" ]);
  cList.unshift([ "Destaque"]);
  vSWOT_Type.unshift(["FOFA"]);
  vFactor.unshift(["Fator"]);
  iSignificance.unshift(["Significancia"]);
  vAuthor.unshift(["Autor"]);
  vDate.unshift(["Data"]);
  vIDcomment.unshift(["Comentario"]);
  vBibliografia.unshift(["ReferênciaBibliografica"]);
  vReferencia.unshift(["Referencia"]);
  vTipoReferencia.unshift(["TipoReferencia"]);
  vBib.unshift(["Bib"]);
  vURL.unshift(["URL"]);
  vLongitude.unshift(["Longitude"]);
  vLatitude.unshift(["Latitude"]);

  // Set values to cells
  //var sheet = SpreadsheetApp.getActiveSheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados");
  sheet.getRange("A1:Q").clearContent(); // Clear all previews content
  sheet.getRange("A1:Q").clearFormat(); // Clear all previews format


  sheet.getRange("A1:A" + hList.length).setValues(hList);
  sheet.getRange("B1:B" + cList.length).setValues(cList);
  sheet.getRange("C1:C" + vFactor.length).setValues(vFactor);
  sheet.getRange("D1:D" + vSWOT_Type.length).setValues(vSWOT_Type);
  sheet.getRange("E1:E" + iSignificance.length).setValues(iSignificance);
  sheet.getRange("F1:F" + vAuthor.length).setValues(vAuthor);
  sheet.getRange("G1:G" + vDate.length).setValues(vDate);
  sheet.getRange("H1:H" + vIDcomment.length).setValues(vIDcomment);
  sheet.getRange("I1:I" + vBibliografia.length).setValues(vBibliografia);

  sheet.getRange("A1:Q1").setFontColor("#FFFFFF");
  sheet.getRange("A1:Q1").setBackground("#2D273D");
  sheet.getRange("A1:Q1").setFontWeight("bold");
  sheet.getRange("A1:Q" + hList.length).setWrap(true);

  keepUnique();

  sheet.getRange("J1").setValue("FatorAtual");
  sheet.getRange("J2").setValue("=ARRAYFORMULA(IF(ISBLANK(IFERROR(VLOOKUP(C2:C; Fatores!A2:C; 2; 0);)); IFERROR(VLOOKUP(C2:C; Fatores!A2:C; 1; 0););IFERROR(VLOOKUP(C2:C; Fatores!A2:C; 2; 0);) ))");

  sheet.getRange("K1").setValue("Aspecto");
  sheet.getRange("K2").setValue("=ARRAYFORMULA(IFERROR(VLOOKUP(C2:C; Fatores!A2:C; 3; 0);))");

  sheet.getRange("J1:O1").setFontColor("#FFFFFF");
  sheet.getRange("J1:O1").setBackground("#2D273D");
  sheet.getRange("J1:O1").setFontWeight("bold");
  sheet.getRange("J1:O" + hList.length).setWrap(true);
  
  sheet.getRange("L1:L" + vReferencia.length).setValues(vReferencia);
  sheet.getRange("M1:M" + vTipoReferencia.length).setValues(vTipoReferencia);  
  sheet.getRange("N1:N" + vBib.length).setValues(vBib);  
  sheet.getRange("O1:O" + vURL.length).setValues(vURL);  
  sheet.getRange("P1:P" + vLongitude.length).setValues(vLongitude);  
  sheet.getRange("Q1:Q" + vLatitude.length).setValues(vLatitude);  


  setScriptIsRunning(false); // Avisa que script não está mais rodando

  var ui = SpreadsheetApp.getUi();
  ui.alert('Planilha atualizada!');
  
}

/**
 * Return the Google Docs file IDUpdate a comment's content.
 *
 * @param {String} url with fileId ID
 */
function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

/**
 * Update a comment's content.
 *
 * @param {String} fileId ID of the file to update comment for.
 * @param {String} commentId ID of the comment to update.
 * @param {String} newContent The new text content for the comment.
 * https://developers.google.com/drive/api/v2/reference/comments/update#javascript
 */
function updateComment(fileId, commentId, newContent) {
  // First retrieve the comment from the API.
  var request = Drive.Comments.get({
    'fileId': fileId,
    'commentId': commentId
  });
  request.execute(function(resp) {
    resp.content = newContent;
    var updateRequest = Drive.Comments.update({
      'fileId': fileId,
      'commentId': commentId,
      'resource': resp
    });
    updateRequest.execute(function(resp) { });
  });
}

function teste() {

updateComment('1q0n6g9PVoO6ew6uGBxjii7RdMEKai4pmn4lx_tzlz6A','AAAAMFUQQqU','[regularização] Refugiados necessitam de vistos permanentes para trazerem suas famílias {am2} Oiii');

}

/**
 * Adapted from:
 *
 * https://www.weirdgeek.com/2020/05/find-and-replace-in-google-apps-script/
 */

function FindandReplace(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data")
  var lastRow = sheet.getLastRow()
  var lastColumn = sheet.getLastColumn()
  var range = sheet.getRange(1, 1, lastRow, lastColumn)
  var to_replace = "Blank";
  var replace_with = "";
  var data  = range.getValues();
 
    var oldValue="";
    var newValue="";
    var cellsChanged = 0;
 
    for (var r=0; r<data.length; r++) {
      for (var i=0; i<data[r].length; i++) {
        oldValue = data[r][i];
        newValue = data[r][i].toString().replace(to_replace, replace_with);
        if (oldValue!=newValue)
        {
          cellsChanged++;
          data[r][i] = newValue;
        }
      }
    }
    range.setValues(data);
}



/**
 * This JavaScript function (specially meant for Google Apps Script or GAS) Get the lookup (vertical) value from a multi-dimensional array.
 *
 * @version 1.1.0
 * Form: https://gist.github.com/narottamdas/26aca662b6eb19322789ec98a445eb18
 * @param {Object} searchValue The value to search for the lookup (vertical).
 * @param {Array} array The multi-dimensional array to be searched.
 * @param {Number} searchIndex The column-index of the array where to search.
 * @param {Number} returnIndex The column-index of the array from where to get the returning matching value.
 * @return {Object} Returns the matching value found else returns null.
 */
function arrayLookup(searchValue,array,searchIndex,returnIndex) 
{
  var returnVal = null;
  var i;
  for(i=0; i<array.length; i++)
  {
    if(array[i][searchIndex]==searchValue)
    {
      returnVal = array[i][returnIndex];
      break;
    }
  }
  
  return returnVal;
}

/**
* Get unique values in column, assuming data starts after header
* Adapted from: https://stackoverflow.com/questions/17372322/gathering-all-the-unique-values-from-one-column-and-outputting-them-in-another-c
*/
function keepUnique(){
  var col = 2 ; // choose the column you want to use as data source (0 indexed, it works at array level)

  // Cria validação de dados para coluna ASPECTO
  //Source: https://stackoverflow.com/questions/58746916/google-sheets-data-validation-drop-down-menu-show-values-but-paste-formula

  var option = new Array();
  option[0]="160-Ético";
  option[1]="150-Estético";
  option[2]="140-Jurídico";
  option[3]="130-Operacional/Técnico";
  option[4]="120-Econômico";
  option[5]="110-Social/Relacional";
  option[6]="100-Conhecimento";
  option[7]="090-Linguístico/Informacional";
  option[8]="080-Histórico/Cultural";
  option[9]="070-Fiducial (fé/convicção)";
  option[10]="060-Sensitivo (psiquico)";
  option[11]="050-Biótico";
  option[12]="040-Físico";
  option[13]="030-Cinemático (movimento)";
  option[14]="020-Espacial";
  option[15]="010-Numérico";
  option[16]="000-INDEFINIDO"


  // -----------------------------------------------------    Cria matriz com tabela Dados (apenas coluna Fatores)
  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  var ss = sh.getSheetByName('Dados');
  var data = ss.getDataRange().getValues();  // get all data from Dados
  
  var dados_fatores_aux = new Array();
  for(nn in data){
    var duplicate = false;
    for(j in dados_fatores_aux){
      if(data[nn][col] == dados_fatores_aux[j][0]){
        duplicate = true;
      }
    }
    if(!duplicate){
      dados_fatores_aux.push([data[nn][col]]);
    }
  }
  dados_fatores_aux.shift(); // remove header
  dados_fatores_aux.sort(); 

  // Aplica filtro e retira elementos em branco, ou nulos, da matriz
  var dados_fatores = dados_fatores_aux.filter(el => {
    return el != null && el != '';
  });
  
  //Logger.log(dados_fatores);

  // --------------------------------------------------------------------------- Cria matriz com tabela fatores

  var sh2 = SpreadsheetApp.getActiveSpreadsheet();  
  var ss2 = sh2.getSheetByName('Fatores');
  var data2 = ss2.getDataRange().getValues();

  //Logger.log(data2);

  var fatores_substituicao = new Array();

  for(nn2 in data2){
      fatores_substituicao.push(data2[nn2]);   
  }
  fatores_substituicao.shift(); // remove header
  
  //Logger.log(fatores_substituicao);
 
   // -------------------------------------------------------- Verifica se fatores da tabela dados estão na tabela fatores
  for (let i in dados_fatores) {
    var procura_fator = arrayLookup(dados_fatores[i],fatores_substituicao,0,0);

    if ( procura_fator === null) {

      // Se NÃO encontrar fatores na tabela de substituicao - inclui
      dados_fatores[i][2] = "000-INDEFINIDO";
      dados_fatores[i][3] = "*";
      fatores_substituicao.push(dados_fatores[i]);
    } else {
      // Se encontrar altera coluna controle apenas
      //                                OR
      if (( dados_fatores[i][2] === "") || ( dados_fatores[i][2] === null )){
        dados_fatores[i][2] = "000-INDEFINIDO";
      }
      dados_fatores[i][3] = "*";

      // procura posição do fator atual na array
      var fator_index;
      for( var x = 0, len = fatores_substituicao.length; x < len; x++ ) {
          if( fatores_substituicao[x][0] === procura_fator ) {
              fator_index = x;
              break;
          }
      }
      // altera a array a partir da posição encontrada
      //fatores_substituicao[fator_index][2] = dados_fatores[i][2];
      fatores_substituicao[fator_index][3] = dados_fatores[i][3];
    }
  }
  //Logger.log(fatores_substituicao);
  
 
  // Remover todos os fatores com Controle = ""
  for( var x = 0, len = fatores_substituicao.length; x < len; x++ ) {
      if ( x >= fatores_substituicao.length ){
        break
      } 

      if ( fatores_substituicao[x][3] === "" ){
        fatores_substituicao.splice(x,1);
      }
  }

  
  //Logger.log(fatores_substituicao);


  // trocar todos os controles de * para ""
  for( var x = 0, len = fatores_substituicao.length; x < len; x++ ) {
      fatores_substituicao[x][3] = "";
  }
  
  //Logger.log(fatores_substituicao);

  // Organiza os fatores em ordem ascendente
  arraySort(fatores_substituicao);

  //Logger.log(fatores_substituicao);

  // ------------------------------------            Write all updated values to the sheet, at once
  var ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fatores");
  ss3.getRange("A2:C").clearContent(); // Clear all previous content
  ss3.getRange("A2:C").clearFormat(); // Clear all previous format
  ss3.getRange(2,1,fatores_substituicao.length,fatores_substituicao[0].length).setValues(fatores_substituicao);

  var aspectos = ss3.getRange("C2:C" + fatores_substituicao.length);
  var rule = SpreadsheetApp.newDataValidation()
             .requireValueInList(option)
             .setAllowInvalid(false)
             .build();
  aspectos.setDataValidation(rule);


}

// Multidimensional array sort function
// Adapted from: https://stackoverflow.com/questions/6993302/javascript-sort-multidimensional-array#6993632
function arraySort(pArray)
{
pArray.sort(
  function(a,b)
  {
    var len=a.length;
    for (var i=0;i<len;i++)
    {
      if (a[i]>b[i]) return 1;
      else if (a[i]<b[i]) return -1;
    }
    return 0;
  }
);
}

// From: https://pt.stackoverflow.com/questions/237762/remover-acentos-javascript
function removeAcento(text){

  text = text.toLocaleUpperCase();                                                         
  text = text.replace(new RegExp('[ÁÀÂÃ]','gi'), 'A');
  text = text.replace(new RegExp('[ÉÈÊ]','gi'), 'E');
  text = text.replace(new RegExp('[ÍÌÎ]','gi'), 'I');
  text = text.replace(new RegExp('[ÓÒÔÕ]','gi'), 'O');
  text = text.replace(new RegExp('[ÚÙÛ]','gi'), 'U');
  text = text.replace(new RegExp('[Ç]','gi'), 'C');
  return text; 
}

function testBibleBookName(bookName){
  switch(bookName){
    case "GEN": 
    case "EXO": 
    case "LEV": 
    case "NUM": 
    case "DEU": 
    case "JOS": 
    case "JUD": 
    case "RUT": 
    case "1SA": 
    case "2SA": 
    case "1KI": 
    case "2KI": 
    case "1CH": 
    case "2CH": 
    case "EZR": 
    case "NEH": 
    case "EST": 
    case "JOB": 
    case "PSA": 
    case "PRO": 
    case "ECC": 
    case "SON": 
    case "ISA": 
    case "JER": 
    case "LAM": 
    case "EZE": 
    case "DAN": 
    case "HOS": 
    case "JOE": 
    case "AMO": 
    case "OBA": 
    case "JON": 
    case "MIC": 
    case "NAH": 
    case "HAB": 
    case "ZEP": 
    case "HAG": 
    case "ZEC": 
    case "MAL": 
    case "MAT": 
    case "MAR": 
    case "LUK": 
    case "JOH": 
    case "ACT": 
    case "ROM": 
    case "1CO": 
    case "2CO": 
    case "GAL": 
    case "EPH": 
    case "PHP": 
    case "COL": 
    case "1TH": 
    case "2TH": 
    case "1TI": 
    case "2TI": 
    case "TIT": 
    case "PHM": 
    case "HEB": 
    case "JAM": 
    case "1PE": 
    case "2PE": 
    case "1JO": 
    case "2JO": 
    case "3JO": 
    case "JDE": 
    case "REV":     
      break;
    default:
      bookName = "";
      break;
  }
return bookName;
}

// Cria uma variável que avisa que há scripts rodando, com isso é possível evitar que trigers sejam disparados, por exemplo.
// Adapted from: https://stackoverflow.com/questions/48809246/check-if-function-is-running-with-google-apps-script-and-google-web-apps#48812935
//
function setScriptIsRunning(value){

  CacheService.getScriptCache().put("isRunning", value.toString());

}

// Verifica se há scripts rodando
function checkRunning() {
var currentState = CacheService.getScriptCache().get("isRunning");

 return (currentState == 'true');
}

// Return mime type of a Google Drive file
// Adapted from: https://www.labnol.org/code/19912-mime-types-google-drive
function getFileType(mimeType) {

  var filetype = "";

  switch (mimeType) {
    case MimeType.GOOGLE_APPS_SCRIPT: filetype = 'Google Apps Script'; break;
    case MimeType.GOOGLE_DRAWINGS: filetype = 'Google Drawings'; break;
    case MimeType.GOOGLE_DOCS: filetype = 'Google Docs'; break;
    case MimeType.GOOGLE_FORMS: filetype = 'Google Forms'; break;
    case MimeType.GOOGLE_SHEETS: filetype = 'Google Sheets'; break;
    case MimeType.GOOGLE_SLIDES: filetype = 'Google Slides'; break;
    case MimeType.FOLDER: filetype = 'Google Drive folder'; break;
    case MimeType.BMP: filetype = 'BMP'; break;
    case MimeType.GIF: filetype = 'GIF'; break;
    case MimeType.JPEG: filetype = 'JPEG'; break;
    case MimeType.PNG: filetype = 'PNG'; break;
    case MimeType.SVG: filetype = 'SVG'; break;
    case MimeType.PDF: filetype = 'PDF'; break;
    case MimeType.CSS: filetype = 'CSS'; break;
    case MimeType.CSV: filetype = 'CSV'; break;
    case MimeType.HTML: filetype = 'HTML'; break;
    case MimeType.JAVASCRIPT: filetype = 'JavaScript'; break;
    case MimeType.PLAIN_TEXT: filetype = 'Plain Text'; break;
    case MimeType.RTF: filetype = 'Rich Text'; break;
    case MimeType.OPENDOCUMENT_GRAPHICS: filetype = 'OpenDocument Graphics'; break;
    case MimeType.OPENDOCUMENT_PRESENTATION: filetype = 'OpenDocument Presentation'; break;
    case MimeType.OPENDOCUMENT_SPREADSHEET: filetype = 'OpenDocument Spreadsheet'; break;
    case MimeType.OPENDOCUMENT_TEXT: filetype = 'OpenDocument Word'; break;
    case MimeType.MICROSOFT_EXCEL: filetype = 'Microsoft Excel'; break;
    case MimeType.MICROSOFT_EXCEL_LEGACY: filetype = 'Microsoft Excel'; break;
    case MimeType.MICROSOFT_POWERPOINT: filetype = 'Microsoft PowerPoint'; break;
    case MimeType.MICROSOFT_POWERPOINT_LEGACY: filetype = 'Microsoft PowerPoint'; break;
    case MimeType.MICROSOFT_WORD: filetype = 'Microsoft Word'; break;
    case MimeType.MICROSOFT_WORD_LEGACY: filetype = 'Microsoft Word'; break;
    case MimeType.ZIP: filetype = 'ZIP'; break;
    default: filetype = "Unknown";
  }

  return filetype;

}

// Test if a specific property exists inside an object
// Adapted from: https://stackoverflow.com/questions/4676223/check-if-object-member-exists-in-nested-object
function testProperty(obj, prop) {
    var parts = prop.split('.');
    for(var i = 0, l = parts.length; i < l; i++) {
        var part = parts[i];
        if(obj !== null && typeof obj === "object" && part in obj) {
            obj = obj[part];
        }
        else {
            return false;
        }
    }
    return true;
}
