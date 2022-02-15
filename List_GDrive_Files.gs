// Adapted from: https://www.acrosswalls.org/ortext-datalinks/list-google-drive-folder-file-names-urls/
// List all files in a specific Google Drive Folder
// Retorna uma array com a lista de urls para ser incluída automaticamente no final da lista da planilha do Socio.Atlas 
// Assim os usuários podem preencher a planilha fontes com arquivos ou pastas (com arquivos)
function listFolderContents(linkAux) {
  var docsList = new Array();
  var folderID = getIdFromUrl(linkAux);
  var folder = DriveApp.getFolderById(folderID);
  var contents = folder.getFiles();
  var file;

  while(contents.hasNext()) {
    file = contents.next();
    // inclui apenas links de documentos

    if (file.getUrl().includes('document') || file.getUrl().includes('file')){
      docsList.push([file.getUrl(), file.getName()]);
    }
  }  
  return docsList;
};