/**
 * @OnlyCurrentDoc
 * ver: https://developers.google.com/apps-script/guides/services/authorization
 */

// [START apps_script_triggers_oninstall]
/**
 * The event handler triggered when installing the add-on.
 * @param {Event} e The onInstall event.
 */
function onInstall(e) {
  onOpen(e);
}
// [END apps_script_triggers_oninstall]


// Based on code at: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#addMenu(String,Object)

// The onOpen function is executed automatically every time a Spreadsheet is loaded
function onOpen() {

  /*
  var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp.
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    // Add a normal menu item (works in all authorization modes).
    menu.addItem('Start workflow', 'startWorkflow');
  } else {
    // Add a menu item based on properties (doesn't work in AuthMode.NONE).
    var properties = PropertiesService.getDocumentProperties();
    var workflowStarted = properties.getProperty('workflowStarted');
    if (workflowStarted) {
      menu.addItem('Check workflow status', 'checkWorkflow');
    } else {
      menu.addItem('Start workflow', 'startWorkflow');
    }
  }
  menu.addToUi();

  */
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  /*
  switch(Session.getActiveUserLocale())
  {
  case "en":
    menuEntries.push({name: "📥 Import fonts", functionName: "ExtrairComentarios"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "Factors list", functionName: "doFactors"});
    menuEntries.push({name: "Simple analysis", functionName: "doTable"});
    menuEntries.push({name: "🌎 Georeferencing", functionName: "doMap"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "🔄 Create Multimodal matrix", functionName: "GeraMatrizMM"});
    menuEntries.push({name: "Multimodal analysis", functionName: "doMultimodal"});
    menuEntries.push({name: "Multimodal chart", functionName: "doGet"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "Tutorial", functionName: "openTutorial"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "Version", functionName: "DoVersaoAtual"});
    break;
  default:
  */
    menuEntries.push({name: "📥 Importar fontes", functionName: "ExtrairComentarios"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "Relação de fatores", functionName: "doFactors"});
    menuEntries.push({name: "Análise simples", functionName: "doTable"});
    menuEntries.push({name: "🌎 Georreferenciamento", functionName: "doMap"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "🔄 Gera matriz multimodal", functionName: "GeraMatrizMM"});
    menuEntries.push({name: "Análise multimodal", functionName: "doMultimodal"});
    menuEntries.push({name: "Gráfico multimodal", functionName: "doGet"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "Tutorial", functionName: "openTutorial"});
    menuEntries.push(null); // line separator
    menuEntries.push({name: "Versão atual", functionName: "DoVersaoAtual"});
    menuEntries.push({name: "Start!", functionName: "DoStart"});

    /*
    break;
  }
  */
  ss.addMenu("SocioAtlas", menuEntries);
}


function DoStart(){
  // Verify spreadsheet structure, and create one if nedded
  // Adapted from: https://stackoverflow.com/questions/58090521/in-google-sheets-apps-script-how-to-check-if-a-sheet-exists-and-if-it-doesnt

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if( ss.getSheetByName("Fontess") == null) {
    //if returned null means the sheet doesn't exist, so create it
  
    /*  
    var val = Utilities.formatDate(new Date(A2), "GMT+1", "MM/dd/yyyy");
    ss.insertSheet(val, ss.getSheets().length, {template: templateSheet});
    sheet2 = ss.insertSheet(A2);
    */
    newSheet = ss.insertSheet("Fontess");
    Logger.log('Criei');
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