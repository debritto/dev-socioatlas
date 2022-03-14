function myGraphViz() {
  var script = `digraph G {labelloc="b"; size="8,6"; ratio=fill;label="(cc) www.socioatlas.xyz";labeljust=right;
node[fontsize=20,shape=plaintext,fontcolor=blue];edge[color=white];`;

  var sh = SpreadsheetApp.getActiveSpreadsheet();  
  var ss = sh.getSheetByName('Matriz');
  var DadosMatriz = ss.getDataRange().getValues(); // get Matrix data 

  var modalidades = [], modalidadesNormativas = [], modalidadesDeterminativas = [], modalidades_aux = [], modalidadesFatoresNormativos = [], modalidadesFatoresDeterminativos = [], modalidadesFatores = [], modalidadesFatores_aux = [];

  // Retorna todos os itens juntamente com o número de fatores em cada
  modalidadesNormativas = SUPERSQL("Select Distinct Modalidade_Normativa From ? ",DadosMatriz);
  modalidadesNormativas.shift(); // remove header
  
  modalidadesDeterminativas = SUPERSQL("Select Distinct Modalidade_Determinativa From ? ",DadosMatriz);
  modalidadesDeterminativas.shift(); // remove header

  modalidades_aux = modalidadesNormativas.concat(modalidadesDeterminativas);

  // Adapted from: https://stackoverflow.com/questions/16747798/delete-duplicate-elements-from-an-array#16747921
  // Remove modalidades duplicadas
  modalidades = (function(modalidades_aux){
  var m = {}, modalidades = []
  for (var i=0; i<modalidades_aux.length; i++) {
    var v = modalidades_aux[i];
    if (!m[v]) {
      modalidades.push(v);
      m[v]=true;
    }
  }
  return modalidades;
  })(modalidades_aux);
  //
  modalidades.sort(); // sorteira em ordem ascendente
  modalidades.reverse(); // reverte ordem da matrix, deixando-a em ordem descendente. Assim a modalidade mais determinativa fica no final.
  // 
  var number_of_modalities = modalidades.length -1; 
  for (let m in modalidades){
    script += '"' + modalidades[m];
    if (m < number_of_modalities){
      script += '" -> ';
    } else {
      script += '";'
    }
  }
  script +=`node[fontsize=18,shape=note,fontcolor=black];
edge[color=black];`;


  modalidadesFatoresNormativos = SUPERSQL(`select  '{rank=same;"' + Modalidade_Normativa + '" "' + Fator_Normativo + '";}' as modfat from ? `,DadosMatriz);
  modalidadesFatoresNormativos.shift(); // remove header
  
  modalidadesFatoresDeterminativos = SUPERSQL(`select  '{rank=same;"' + Modalidade_Determinativa + '" "' + Fator_Determinativo + '";}' as modfat from ? `,DadosMatriz);
  modalidadesFatoresDeterminativos.shift(); // remove header
  modalidadesFatores_aux = modalidadesFatoresNormativos.concat(modalidadesFatoresDeterminativos);

  // Remove duplicados
  modalidadesFatores = (function(modalidadesFatores_aux){
  var m = {}, modalidadesFatores = []
  for (var i=0; i<modalidadesFatores_aux.length; i++) {
    var v = modalidadesFatores_aux[i];
    if (!m[v]) {
      modalidadesFatores.push(v);
      m[v]=true;
    }
  }
  return modalidadesFatores;
  })(modalidadesFatores_aux);
  //
  modalidadesFatores.sort(); // sorteira em ordem ascendente
  modalidadesFatores.reverse(); // reverte ordem da matrix, deixando-a em ordem descendente. Assim a modalidade mais determinativa fica no final.
  
  //The join() method creates and returns a new string by concatenating all elements in an array, separated by a specified separator string. 
  script += modalidadesFatores.join(' ');

  // analisa vínculos
  var vinculos_aux = SUPERSQL(`Select '"' + Fator_Normativo + '"->"' + Fator_Determinativo + '"' as vinculos, Vinculo  From ? `,DadosMatriz);
  vinculos_aux.shift(); // remove header

  var vinculo = "";
  var setas = "";
  for (let m in vinculos_aux){
    vinculo = vinculos_aux[m][0];
    setas = vinculos_aux[m][1];

    switch(setas)
    {
    case "⇦": // 1
      script += `edge[dir="back"]` + vinculo + `[arrowtail="onormal"][arrowhead="none"][arrowsize="1"][color=blue][style=filled];`;
      break;
    case "⬅⇨": // 2
      script += `edge[dir="back"]` + vinculo + `[arrowtail="normal"][arrowhead="none"][arrowsize="1"][color=red][style=filled];`;
      // Make diferent lines For both arrow types
      script += `edge[dir="forward"]` + vinculo + `[arrowtail="none"][arrowhead="onormal"][arrowsize="1"][color=blue][style=filled];`;
      break;
    case "⬅": // 3
      script += `edge[dir="back"]` + vinculo + `[arrowtail="normal"][arrowhead="none"][arrowsize="1"][color=red][style=filled];`;
      break;
    case "⬅➡": //4
      script += `edge[dir="both"]` + vinculo + `[arrowtail="normal"][arrowhead="normal"][arrowsize="1"][color=red][style=filled];`;
      break;
    case "⇦➡": // 5
      script += `edge[dir="back"]` + vinculo + `[arrowtail="onormal"][arrowhead="none"][arrowsize="1"][color=blue][style=filled];`;
      // Make diferent lines For both arrow types
      script += `edge[dir="forward"]` + vinculo + `[arrowtail="none"][arrowhead="normal"][arrowsize="1"][color=red][style=filled];`;
      break;
    case "⇨":  //6
      script += `edge[dir="forward"]` + vinculo + `[arrowtail="none"][arrowhead="onormal"][arrowsize="1"][color=blue][style=filled];`;
      break;            
    case "➡": //7
      script += `edge[dir="forward"]` + vinculo + `[arrowtail="none"][arrowhead="normal"][arrowsize="1"][color=red][style=filled];`;
      break;            
    case "⇦⇨": //8
      script += `edge[dir="both"]` + vinculo + `[arrowtail="onormal"][arrowhead="onormal"][arrowsize="1"][color=blue][style=filled];`;
      break;            
    default:
      script += `edge[dir="both"]` + vinculo + `[arrowtail="none"][arrowhead="none"][arrowsize="3"][color=black][style=dotted];`;
      break;
    }
  }
  // 
  script += "}";
  //Logger.log(script);

  //var ChartImageURL = encodeURI("https://image-charts.com/chart?cht=gv:dot&chof=.svg&chl=" + script);
  //openUrl(ChartImageURL);
  
  //Logger.log(ChartImageURL);
 
  return script;
}

/**
 * Open a URL in a new tab.
 * Adapted from: https://stackoverflow.com/questions/10744760/google-apps-script-to-open-a-url#10745403
 */
function openUrl( url ){
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
  +'<body style="word-break:break-word;font-family:sans-serif;">Erro ao abrir automaticamente. <a href="'+url+'" target="_blank" onclick="window.close()">Clique AQUI para continuar!</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Abrindo ..." );
}


function BibleViz(){


  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  var ss = sh.getSheetByName('Dados');
  var Dados = ss.getDataRange().getValues();

var BibleUse = SUPERSQL("select FatorAtual as `Fatores`, count(case when Bib = 'GEN' then Bib end) as `GEN`, count(case when Bib = 'EXO' then Bib end) as `EXO`, count(case when Bib = 'LEV' then Bib end) as `LEV`, count(case when Bib = 'NUM' then Bib end) as `NUM`, count(case when Bib = 'DEU' then Bib end) as `DEU`, count(case when Bib = 'JOS' then Bib end) as `JOS`, count(case when Bib = 'JUD' then Bib end) as `JUD`, count(case when Bib = 'RUT' then Bib end) as `RUT`, count(case when Bib = '1SA' then Bib end) as `1SA`, count(case when Bib = '2SA' then Bib end) as `2SA`, count(case when Bib = '1KI' then Bib end) as `1KI`, count(case when Bib = '2KI' then Bib end) as `2KI`, count(case when Bib = '1CH' then Bib end) as `1CH`, count(case when Bib = '2CH' then Bib end) as `2CH`, count(case when Bib = 'EZR' then Bib end) as `EZR`, count(case when Bib = 'NEH' then Bib end) as `NEH`, count(case when Bib = 'EST' then Bib end) as `EST`, count(case when Bib = 'JOB' then Bib end) as `JOB`, count(case when Bib = 'PSA' then Bib end) as `PSA`, count(case when Bib = 'PRO' then Bib end) as `PRO`, count(case when Bib = 'ECC' then Bib end) as `ECC`, count(case when Bib = 'SON' then Bib end) as `SON`, count(case when Bib = 'ISA' then Bib end) as `ISA`, count(case when Bib = 'JER' then Bib end) as `JER`, count(case when Bib = 'LAM' then Bib end) as `LAM`, count(case when Bib = 'EZE' then Bib end) as `EZE`, count(case when Bib = 'DAN' then Bib end) as `DAN`, count(case when Bib = 'HOS' then Bib end) as `HOS`, count(case when Bib = 'JOE' then Bib end) as `JOE`, count(case when Bib = 'AMO' then Bib end) as `AMO`, count(case when Bib = 'OBA' then Bib end) as `OBA`, count(case when Bib = 'JON' then Bib end) as `JON`, count(case when Bib = 'MIC' then Bib end) as `MIC`, count(case when Bib = 'NAH' then Bib end) as `NAH`, count(case when Bib = 'HAB' then Bib end) as `HAB`, count(case when Bib = 'ZEP' then Bib end) as `ZEP`, count(case when Bib = 'HAG' then Bib end) as `HAG`, count(case when Bib = 'ZEC' then Bib end) as `ZEC`, count(case when Bib = 'MAL' then Bib end) as `MAL`, count(case when Bib = 'MAT' then Bib end) as `MAT`, count(case when Bib = 'MAR' then Bib end) as `MAR`, count(case when Bib = 'LUK' then Bib end) as `LUK`, count(case when Bib = 'JOH' then Bib end) as `JOH`, count(case when Bib = 'ACT' then Bib end) as `ACT`, count(case when Bib = 'ROM' then Bib end) as `ROM`, count(case when Bib = '1CO' then Bib end) as `1CO`, count(case when Bib = '2CO' then Bib end) as `2CO`, count(case when Bib = 'GAL' then Bib end) as `GAL`, count(case when Bib = 'EPH' then Bib end) as `EPH`, count(case when Bib = 'PHP' then Bib end) as `PHP`, count(case when Bib = 'COL' then Bib end) as `COL`, count(case when Bib = '1TH' then Bib end) as `1TH`, count(case when Bib = '2TH' then Bib end) as `2TH`, count(case when Bib = '1TI' then Bib end) as `1TI`, count(case when Bib = '2TI' then Bib end) as `2TI`, count(case when Bib = 'TIT' then Bib end) as `TIT`, count(case when Bib = 'PHM' then Bib end) as `PHM`, count(case when Bib = 'HEB' then Bib end) as `HEB`, count(case when Bib = 'JAM' then Bib end) as `JAM`, count(case when Bib = '1PE' then Bib end) as `1PE`, count(case when Bib = '2PE' then Bib end) as `2PE`, count(case when Bib = '1JO' then Bib end) as `1JO`, count(case when Bib = '2JO' then Bib end) as `2JO`, count(case when Bib = '3JO' then Bib end) as `3JO`, count(case when Bib = 'JDE' then Bib end) as `JDE`, count(case when Bib = 'REV' then Bib end) as `REV` from ? group by FatorAtual", Dados);

Logger.log(BibleUse);
}