function prepara_substituicao(){

// Put all sheet data inside _arr_ matrix. Note that you could have a memory overflow.
let Dados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados").getDataRange().getValues();
let Fatores = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Substituir").getDataRange().getValues();

let Dados_headers = Dados.shift();
let Fatores_headers = Fatores.shift();

let Dados_data = Dados.map( r => {
	let obj = {};
	r.forEach((cell,i) => {
		obj[Dados_headers[i]] = cell;
	})

	return obj;
});

let Fatores_data = Fatores.map( r => {
	let obj = {};
	r.forEach((cell,i) => {
		obj[Fatores_headers[i]] = cell;
	})

	return obj;
});

  var res = alasql( "select distinct Fator from ?  order by Fator", [Dados_data]);
  var res2 = alasql( "select distinct * from ?", [Fatores_data]);

  console.log(res);

  console.log(res2);


}



// Adapted from: https://stackoverflow.com/questions/42150450/google-apps-script-for-multiple-find-and-replace-in-google-sheets#42174750

function runReplaceInSheet(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Substituir");
  //  get the current data range values as an array
  //  Fewer calls to access the sheet -> lower overhead 
  var values = sheet.getDataRange().getValues();  

  // Replace Subject Names
  replaceInSheet(values, /\d\dART\d\d/g, "Art");
  replaceInSheet(values, /\d\dCCL\d\d/g, "Communication & Culture");
  replaceInSheet(values, /\d\dDLT\d\d/g, "Digital Technology");
  replaceInSheet(values, /\d\dDRA\d\d/g, "Drama");

  // Replace Staff Names
  replaceInSheet(values, 'TED', 'Tahlee Edward');
  replaceInSheet(values, 'TLL', 'Tyrone LLoyd');
  replaceInSheet(values, 'TMA', 'Timothy Mahone');
  replaceInSheet(values, 'TQU', 'Tom Quebec');

  // Write all updated values to the sheet, at once
  sheet.getDataRange().setValues(values);
}

function replaceInSheet(values, to_replace, replace_with) {
  //loop over the rows in the array
  for(var row in values){
    //use Array.map to execute a replace call on each of the cells in the row.
    var replaced_values = values[row].map(function(original_value) {
      return original_value.toString().replace(to_replace,replace_with);
    });

    //replace the original row values with the replaced values
    values[row] = replaced_values;
  }
}