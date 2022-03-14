// Adapted from: https://opensourcelibs.com/lib/alasqlgs
/*
  The code here uses http://alasql.org library:
  
  Downlaad it from here:
    https://raw.githubusercontent.com/agershun/alasql/develop/dist/alasql.min.js
  or here (not tested)
    https://cdn.jsdelivr.net/npm/alasql
    
    My sample sheet is here:
    https://docs.google.com/spreadsheets/d/1V0kHvuS0QfzgYTvkut9UkwcgK_51KV2oHDxKE6dMX7A/copy

    Adapted from here by Software DeBritto: https://stackoverflow.com/questions/17930389/can-the-google-spreadsheet-query-function-be-used-in-google-apps-script#28751946
*/

function test_AlaSqlQuery()
{

//  var file = SpreadsheetApp.getActive();
//  var sheet1 = file.getSheetByName('East');
//  var range1 = sheet1.getDataRange();
//  var data1 = range1.getValues();
  
//  var sheet2 = file.getSheetByName('Reps');
//  var range2 = sheet2.getDataRange();   
//  var data2 = range2.getValues();


  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  SpreadsheetApp.setActiveSheet(sh.getSheetByName('Dados'))
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data1=ss.getDataRange().getValues();// get all data from Dados 
 

  var sql = `
  SELECT a.Comentario AS aComentario, a.Aspecto AS aAspecto, a.Fator AS aFator, 
  b.Comentario AS bComentario, b.Aspecto AS bAspecto, b.Fator AS bFator 
  FROM ? a 
  JOIN ? b ON a.Aspecto <= b.Aspecto 
  WHERE 
  (a.Aspecto NOT NULL) AND 
  (b.Aspecto NOT NULL) AND 
  (a.Fator <> b.Fator) AND 
  (a.Comentario = aComentario) AND 
  (b.Comentario = bComentario) 
  ORDER BY a.Aspecto`;

  var sql= "select * from ?";
  
  var data = getAlaSql(sql, data1, data1);

  Logger.log(data);  


 
//  var sql = "select a.Col1, a.Col3, reps.Col2, a.Col7 from ? a left join ? reps on reps.Col1 = a.Col3";
//  var data = getAlaSql(sql, data1, data2);
  
//  Logger.log(data);  


}

function getAlaSql(sql)
{
  var tables = Array.prototype.slice.call(arguments, 1);  
  var request = convertToAlaSql_(sql);
  var res = alasql(request, tables);
  //return JSON.stringify(res);

  return convertAlaSqlResultToArray_(res);
}



function test_AlaSqlSelect()
{
  var file = SpreadsheetApp.getActive();
  var sheet = file.getSheetByName('East');
  var range = sheet.getDataRange();
  var data = range.getValues();
  
  var sql = "select * from ? where Col5 > 50 and Col3 = 'Jones'"
  Logger.log(convertAlaSqlResultToArray_(getAlaSqlSelect_(data, sql)));
  /*
  [  
     [  
        Sun Jan 07 12:38:56      GMT+02:00      2018,
        East,
        Jones,
        Binder,
        60.0,
        4.99,
        299.40000000000003                                            // error: precision =(
     ],
    ...
  ]
  
  */

}


function getAlaSqlSelect_(data, sql)
{
  var request = convertToAlaSql_(sql);
  var res = alasql(request, [data]);
  // [{0=2016.0, 1=a, 2=1.0}, {0=2016.0, 1=a, 2=2.0}, {0=2018.0, 1=a, 2=4.0}, {0=2019.0, 1=a, 2=5.0}]
  return convertAlaSqlResultToArray_(res);
}


function convertToAlaSql_(string)
{
  var result = string.replace(/(Col)(\d+)/g, "[$2]");
  result = result.replace(/\[(\d+)\]/g, function(a,n){ return "["+ (+n-1) +"]"; });
  return result;
}


function convertAlaSqlResultToArray_(res)
{
  var result = [];
  var row = [];
  res.forEach
  (
  function (elt)
  {
    row = [];
    for (var key in elt) { row.push(elt[key]); }
    result.push(row);
  }  
  );
  return result;
}