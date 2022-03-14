//
// Adapted from: https://stackoverflow.com/questions/30582284/how-to-extract-exif-data-from-jpeg-files-in-drive
// Documentation: https://developers.google.com/apps-script/guides/services/advanced#enable_advanced_services
//
// Demo use of getPhotoExif()
// Logs all files in a folder named "Photos".
function listPhotos() {
  var files = DriveApp.getFoldersByName("Photos").next().getFiles();
  var fileInfo = [];
  while (files.hasNext()) {
    var file = files.next();
    Logger.log("File: %s, Date taken: %s",
               file.getName(),
               getPhotoExif(file.getId()).date || 'unknown');
  }
}

/**
 * Retrieve imageMediaMetadata for given file. See Files resource
 * representation for details.
 * (https://developers.google.com/drive/v2/reference/files)
 *
 * @param {String} fileId    File ID to look up
 *
 * @returns {object}         imageMediaMetadata object
 */
function getPhotoExif( fileId ) {
  var file = Drive.Files.get(fileId);
  var metaData = file.imageMediaMetadata;

  Logger.log( file.ownerNames )
  Logger.log( file.etag )
  Logger.log (file.mimeType )

  // If metaData is 'undefined', return an empty object
  return metaData ? metaData : {};
}


function TesteFoto(){

Logger.log( getPhotoExif('14BykIRVmkWSYs6wpGvpKp98W6x6sf-ZS') );

// https://drive.google.com/file/d/14BykIRVmkWSYs6wpGvpKp98W6x6sf-ZS/view?usp=sharing
}


