function doFactors() {
  var html = HtmlService.createHtmlOutputFromFile('fatores')
  .setWidth(1400)
  .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Relação de fatores');
}

function scan_Factors() {
  // Open sheet "Aspectos" and retrieve data
  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  var ss = sh.getSheetByName('Aspectos');
  var data1 = ss.getDataRange().getValues();// get all data from Dados 
  var aspectsTable = SUPERSQL("Select Aspecto From ?",data1);
  aspectsTable.shift(); // remove header


  // Open sheet "Dados" and retrieve data
  var sh = SpreadsheetApp.getActiveSpreadsheet();   
  var ss = sh.getSheetByName('Fatores');
  var data2 = ss.getDataRange().getValues();// get all data from Dados 


  /*
  Adapted from: https://www.geeksforgeeks.org/how-to-print-a-circular-structure-in-a-json-like-format-using-javascript/
  Convert a circular structure in a JSON like format
  */
  const circularReplacer = () => {
    
      // Creating new WeakSet to keep 
      // track of previously seen objects
      const seen = new WeakSet();
        
      return (key, value) => {
    
          // If type of value is an 
          // object or value is null
          if (typeof(value) === "object" 
                    && value !== null) {
            
          // If it has been seen before
          if (seen.has(value)) {
                  return;
              }
                
              // Add current value to the set
              seen.add(value);
        }
          
        // return the value
        return value;
    };
  };

  /*
  Adapted from: https://code.tutsplus.com/articles/data-structures-with-javascript-tree--cms-23393

  Methods:
  find(data): return the node containing the given data
  add(data, parentData): add a new node to the parent containing the given data
  remove(data): remove the node containing the given data
  */

  //elements of the tree
  class Node {
      constructor(name) {
          this.name = name;
          this.parent = "null";
          this.children = [];
      }
  }

  //the tree namestructure
  class Tree {
      constructor(name) {
        let node = new Node(name);
        this._root = node;
      }
      
      //returns the node that has the given name, or null
      //use a depth-first search 
      find(name, node = this._root) {
          //if the current node matches the name, return it
          if (node.name == name)
              return node;

          //recurse on each child node
          for (let child of node.children) {
              //if the name is found in any child node it will be returned here 
              if (this.find(name, child))
                  return child;
          }   
          
          //otherwise, the name was not found
          return null;
      }

      //create a new node with the given name and add it to the specified parent node
      add(name, parentname) {
          let node = new Node(name);
          let parent = this.find(parentname);

          //if the parent exists, add this node
          if (parent) {
              parent.children.push(node);
              node.parent = parent;

              //return this node
              return node;
          }
          //otherwise throw an error
          else {
              throw new Error(`Cannot add node: parent with name ${parentname} not found.`);
          }
      }
      
      //removes the node with the specified name from the tree
      remove(name) {
          //find the node
          let node = this.find(name)

          //if the node exists, remove it from its parent
          if (node) {
              //find the index of this node in its parent
              let parent = node.parent;
              let indexOfNode = parent.children.indexOf(node);
              //and delete it from the parent
              parent.children.splice(indexOfNode, 1);
          }
          //otherwise throw an error
          else {
              throw new Error(`Cannot remove node: node with name ${name} not found.`);
          }
      }

      //depth-first tree traversal
      //starts at the root
      forEach(callback, node = this._root) {
          //recurse on each child node
          for (let child of node.children) {
              //if the name is found in any child node it will be returned here 
              this.forEach(callback, child);
          }   
          
          //otherwise, the name was not found
          callback(node);
      }

      //breadth-first tree traversal
      forEachBreadthFirst(callback) {
          //start with the root node
          let queue = [];
          queue.push(this._root);
      
          //while the queue is not empty
          while (queue.length > 0) {
              //take the next node from the queue  
              let node = queue.shift();
              
              //visit it
              callback(node);

              //and enqueue its children
              for (let child of node.children) {
                  queue.push(child);
              }
          }
      }
  }

  /*
          SCAN FACTORS
  */
  var factorsTree = new Tree("Fatores");
  
  for(n_aspects in aspectsTable){
    // Insere todos os aspectos
    //factorsTree.add(aspectsTable[n_aspects][0], "Realidade");

    // Fatores com substituição primeiro
    var factors_query = "Select Distinct Substituicao From ? Where Aspecto ='" + aspectsTable[n_aspects][0] + "'";
    var factorsTable = SUPERSQL(factors_query,data2);
    
    if ( factorsTable) {  
      factorsTable.shift(); // remove header

      for(n_factors in factorsTable){
        if (factorsTable[n_factors][0] !==""){

          //factorsTree.add(factorsTable[n_factors][0], aspectsTable[n_aspects][0]);
          
          var isThereThisFactor = factorsTree.find(factorsTable[n_factors][0]);
          if (!isThereThisFactor){
            factorsTree.add(factorsTable[n_factors][0], "Fatores");
          }
          var factorsAUX_query = "Select Fator From ? Where Substituicao ='" + factorsTable[n_factors][0] + "' Order By Fator";
          var factorsAUXTable = SUPERSQL(factorsAUX_query,data2);

          if ( factorsAUXTable) {  
            factorsAUXTable.shift(); // remove header
            for(n_factorsAUX in factorsAUXTable){

              //factorsTree.add(factorsAUXTable[n_factorsAUX][0], factorsTable[n_factors][0]);
              
              var isThereThisFactor = factorsTree.find(factorsAUXTable[n_factorsAUX][0]);
              if (!isThereThisFactor){
                factorsTree.add(factorsAUXTable[n_factorsAUX][0], factorsTable[n_factors][0]);
              }  
            }
          }
        }
      }
    }

    // Fatores sem substituição agora
    var factors_query = "Select Distinct Fator, Substituicao From ? Where (Substituicao = '' and Aspecto ='" + aspectsTable[n_aspects][0] + "') Order By Fator";
    var factorsTable = SUPERSQL(factors_query,data2);
    
    if ( factorsTable) {  
      factorsTable.shift(); // remove header

      for(n_factors in factorsTable){
          //factorsTree.add(factorsTable[n_factors][0], aspectsTable[n_aspects][0]);

          var isThereThisFactor = factorsTree.find(factorsTable[n_factors][0]);
          if (!isThereThisFactor){
            factorsTree.add(factorsTable[n_factors][0], "Fatores");
          }  
        }
      }

  }  
   
  var jsonString = JSON.stringify(factorsTree, circularReplacer());
  var scannedFactorsAUX = JSON.parse(jsonString);
  //var scannedFactors = scannedFactorsAUX["_root"];
  const scannedFactors = Object.values(scannedFactorsAUX);

  Logger.log(scannedFactors);

  return scannedFactors;
}