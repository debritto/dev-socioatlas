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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
  // executed.
  menuEntries.push({name: "📥 Importar fontes", functionName: "ExtrairComentarios"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "🔖 Relação de fatores", functionName: "doFactors"});
  menuEntries.push({name: "📑 Análise simples", functionName: "doTable"});
  menuEntries.push({name: "🌎 Georreferenciamento", functionName: "doMap"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "🔄 Gera matriz multimodal", functionName: "GeraMatrizMM"});
  menuEntries.push({name: "🔁 Análise multimodal", functionName: "doMultimodal"});
  menuEntries.push({name: "Gráfico multimodal", functionName: "doGet"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "💡 Tutorial", functionName: "openTutorial"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "Versão atual", functionName: "DoVersaoAtual"});


 
  ss.addMenu("SocioAtlas", menuEntries);
}

function DoVersaoAtual(){
  var versaoAtual = "0.0beta7";
  
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