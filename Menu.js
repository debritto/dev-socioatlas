/**
 * @OnlyCurrentDoc
 * ver: https://developers.google.com/apps-script/guides/services/authorization
 */

/**
 * The event handler triggered when installing the add-on.
 * @param {Event} e The onInstall event.
 */
function onInstall(e) {
  onOpen(e);
}

// The onOpen function is executed automatically every time a Spreadsheet is loaded
function onOpen(e) {
  console.info(e);
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('📥 Importar fontes', 'ExtrairComentarios')
    .addSeparator()
    .addItem('Relação de fatores', 'doFactors')
    .addItem('Análise simples', 'doTable')
    .addItem('🌎 Georreferenciamento', 'doMap')
    .addSeparator()
    .addItem('🔄 Gera matriz multimodal', 'GeraMatrizMM')
    .addItem('Análise multimodal', 'doMultimodal')
    .addItem('Gráfico multimodal', 'doGet')
    .addSeparator()
    .addItem('Tutorial', 'openTutorial')
    .addSeparator()
    .addItem('Versão atual', 'DoVersaoAtual')
    .addItem('Start!', 'DoStart')
    .addToUi();
}

function DoStart(){
  // Verify spreadsheet structure, and create one if nedded
  // Adapted from: https://stackoverflow.com/questions/58090521/in-google-sheets-apps-script-how-to-check-if-a-sheet-exists-and-if-it-doesnt

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if( ss.getSheetByName("Fontes") == null) {
    //if returned null means the sheet doesn't exist, so create it
    newSheet = ss.insertSheet("Fontes");
    newSheet.getRange("A1").setValue("Fontes");
    newSheet.getRange("B1").setValue("Referências bibliográficas");
    newSheet.getRange("A1:B1").setFontColor("#FFFFFF");
    newSheet.getRange("A1:B1").setBackground("#2D273D");
    newSheet.getRange("A1:B1").setFontWeight("bold");
    newSheet.getRange("A1:B1").setWrap(true);
    newSheet.setColumnWidth(1,700);
    newSheet.setColumnWidth(2,700);
  }

  if( ss.getSheetByName("Aspectos") == null) {
    //if returned null means the sheet doesn't exist, so create it
    newSheet = ss.insertSheet("Aspectos");
    var aspecto = [], nucleo = [], identificacao = [];

    aspecto.push(['Aspecto']);
    aspecto.push(['160-Ético']);
    aspecto.push(['150-Estético']);
    aspecto.push(['140-Jurídico']);
    aspecto.push(['130-Operacional/Técnico']);
    aspecto.push(['120-Econômico']);
    aspecto.push(['110-Social/Relacional']);
    aspecto.push(['100-Conhecimento']);
    aspecto.push(['090-Linguístico/Informacional']);
    aspecto.push(['080-Histórico/Cultural']);
    aspecto.push(['070-Fiducial (fé/convicção)']);
    aspecto.push(['060-Sensitivo (psiquico)']);
    aspecto.push(['050-Biótico']);
    aspecto.push(['040-Físico']);
    aspecto.push(['030-Cinemático (movimento)']);
    aspecto.push(['020-Espacial']);
    aspecto.push(['010-Numérico']);
    aspecto.push(['000-INDEFINIDO']);

    nucleo.push(['Núcleo_de_Sentido']);
    nucleo.push(['Ágape']);
    nucleo.push(['Harmonia']);
    nucleo.push(['Justiça']);
    nucleo.push(['Trabalho (vocação)']);
    nucleo.push(['Frugalidade']);
    nucleo.push(['Sociação (co-operação)']);
    nucleo.push(['Sabedoria']);
    nucleo.push(['Significação']);
    nucleo.push(['Criatividade']);
    nucleo.push(['Convicção']);
    nucleo.push(['Senciência']);
    nucleo.push(['Vida']);
    nucleo.push(['Energia']);
    nucleo.push(['Movimento']);
    nucleo.push(['Extensão']);
    nucleo.push(['Magnitude']);
    nucleo.push(['N/D']);

    identificacao.push(['Como_Identificar']);
    identificacao.push(['Quão ético? Há promessas cumpridas ou quebradas? É amável, cuidadoso, sacrificial, seguro?']);
    identificacao.push(['Quão agradável e prazeroso? Há alguma alusão desafiadora, alguma nuance?']);
    identificacao.push(['Quão justo? É justo e correto para todos os envolvidos? A ação ou decisão pode ser justificada? Há muita, ou pouca, regulamentação?']);
    identificacao.push(['Quão realizador é o trabalho? Qual o trabalho necessário?']);
    identificacao.push(['Quão valioso? É acessível, econômico, gerenciável?']);
    identificacao.push(['Quão sociável? Há cooperação e encorajamento? Quais comunidades e associações estão presentes?']);
    identificacao.push(['Quão inteligível e aplicável? Há coerência interna e externa? Qual o grau de aplicabilidade?']);
    identificacao.push(['Quão claro? Há comunicação aberta? Qual linguagem ou símbolos estão sendo utilizados?']);
    identificacao.push(['Quão criativo? Os desenvolvimentos são culturalmente apropriados e úteis?']);
    identificacao.push(['Quão confiável? Quais crenças, cosmovisões, ideologias estão em jogo?']);
    identificacao.push(['Quão estimulante? Os impactos percebidos são ameaçadores ou bem vindos? Quais são os sentimentos, emoções?']);
    identificacao.push(['Quão produtivo? Há relação fecunda, promotora de saúde, entre as coisas vivas?']);
    identificacao.push(['Quão reativo? Há algum uso efetivo, não poluente, sustentável, dos recursos naturais?']);
    identificacao.push(['Quão rápido? Quais processos, fatores, são constantes nesta situação?']);
    identificacao.push(['Quão grande? Há cobertura, solução, resposta adequada em detalhe e alcançe?']);
    identificacao.push(['Quantos? Todas as medições e avaliações foram feitas corretamente?']);
    identificacao.push(['N/D']);

    newSheet.getRange("A1:A" + aspecto.length).setValues(aspecto);
    newSheet.getRange("B1:B" + nucleo.length).setValues(nucleo);
    newSheet.getRange("C1:C" + identificacao.length).setValues(identificacao);
    newSheet.getRange("A1:C1").setFontColor("#FFFFFF");
    newSheet.getRange("A1:C1").setBackground("#2D273D");
    newSheet.getRange("A1:C1").setFontWeight("bold");
    newSheet.getRange("A1:C" + aspecto.length).setWrap(true);
    newSheet.setColumnWidth(1,300);
    newSheet.setColumnWidth(2,300);
    newSheet.setColumnWidth(3,900);
  }

  if( ss.getSheetByName("Dados") == null) {
    //if returned null means the sheet doesn't exist, so create it
    newSheet = ss.insertSheet("Dados");
    newSheet.getRange("A1").setValue("Texto");
    newSheet.getRange("B1").setValue("Destaque");
    newSheet.getRange("C1").setValue("Fator");
    newSheet.getRange("D1").setValue("FOFA");
    newSheet.getRange("E1").setValue("Significancia");
    newSheet.getRange("F1").setValue("Autor");
    newSheet.getRange("G1").setValue("Data");
    newSheet.getRange("H1").setValue("Comentario");
    newSheet.getRange("I1").setValue("ReferênciaBibliografica");
    newSheet.getRange("J1").setValue("FatorAtual");
    newSheet.getRange("K1").setValue("Aspecto");
    newSheet.getRange("L1").setValue("Referencia");
    newSheet.getRange("M1").setValue("TipoReferencia");
    newSheet.getRange("N1").setValue("Bib");
    newSheet.getRange("O1").setValue("URL");
    newSheet.getRange("P1").setValue("Longitude");
    newSheet.getRange("Q1").setValue("Latitude");
    newSheet.getRange("A1:Q1").setFontColor("#FFFFFF");
    newSheet.getRange("A1:Q1").setBackground("#2D273D");
    newSheet.getRange("A1:Q1").setFontWeight("bold");
    newSheet.getRange("A1:Q1").setWrap(true);
    newSheet.setColumnWidth(1,286);
    newSheet.setColumnWidth(2,286);
    newSheet.setColumnWidth(3,162);
    newSheet.setColumnWidth(4,145);
    newSheet.setColumnWidth(5,115);
    newSheet.setColumnWidth(6,223);
    newSheet.setColumnWidth(7,200);
    newSheet.setColumnWidth(8,133);
    newSheet.setColumnWidth(9,570);
    newSheet.setColumnWidth(10,149);
    newSheet.setColumnWidth(11,187);
    newSheet.setColumnWidth(12,99);
    newSheet.setColumnWidth(13,118);
    newSheet.setColumnWidth(14,66);
    newSheet.setColumnWidth(15,600);
    newSheet.setColumnWidth(16,130);
    newSheet.setColumnWidth(17,130);
  }  

  if( ss.getSheetByName("Fatores") == null) {
    //if returned null means the sheet doesn't exist, so create it
    newSheet = ss.insertSheet("Fatores");
    newSheet.getRange("A1").setValue("Fator");
    newSheet.getRange("B1").setValue("Substituicao");
    newSheet.getRange("C1").setValue("Aspecto");
    newSheet.getRange("A1:C1").setFontColor("#FFFFFF");
    newSheet.getRange("A1:C1").setBackground("#2D273D");
    newSheet.getRange("A1:C1").setFontWeight("bold");
    newSheet.getRange("A1:C1").setWrap(true);
    newSheet.setColumnWidth(1,440);
    newSheet.setColumnWidth(2,440);
    newSheet.setColumnWidth(3,440);
  }

  if( ss.getSheetByName("Matriz") == null) {
    //if returned null means the sheet doesn't exist, so create it
    newSheet = ss.insertSheet("Matriz");
    newSheet.getRange("A1").setValue("ID");
    newSheet.getRange("B1").setValue("Modalidade_Normativa");
    newSheet.getRange("C1").setValue("Fator_Normativo");
    newSheet.getRange("D1").setValue("a_Comentario");
    newSheet.getRange("E1").setValue("Vinculo");
    newSheet.getRange("F1").setValue("b_Comentario");
    newSheet.getRange("G1").setValue("Fator_Determinativo");
    newSheet.getRange("H1").setValue("Modalidade_Determinativa");
    newSheet.getRange("I1").setValue("Controle");

    newSheet.getRange("A1:I1").setFontColor("#FFFFFF");
    newSheet.getRange("A1:I1").setBackground("#2D273D");
    newSheet.getRange("A1:I1").setFontWeight("bold");
    newSheet.getRange("A1:I1").setWrap(true);
    newSheet.setColumnWidth(1,50);
    newSheet.setColumnWidth(2,187);
    newSheet.setColumnWidth(3,284);
    newSheet.setColumnWidth(4,99);
    newSheet.setColumnWidth(5,99);
    newSheet.setColumnWidth(6,99);
    newSheet.setColumnWidth(7,284);
    newSheet.setColumnWidth(8,187);
    newSheet.setColumnWidth(9,60);
  }
}

function DoVersaoAtual(){
  var versaoAtual = "0.0 beta9";
  
  var ui = SpreadsheetApp.getUi();
  ui.alert('Versão Atual: ' + versaoAtual + ' | Verifique se há uma versão nova em 👉️ www.socioatlas.xyz');
}

/**
 * Open a URL in a new tab.
 * Based on: https://stackoverflow.com/questions/10744760/google-apps-script-to-open-a-url
 */
function openTutorial(){
  var url = "https://www.socioatlas.xyz/docs/"
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Abrindo painel ..." );
}