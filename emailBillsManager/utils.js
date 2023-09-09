function testCheckFolder() {
  var assert = require("assert");
  var parentFol = DriveApp.getFilesByName('testFolder');
  var folderName = 'testFolder';
  var result = checkFolder(parentFol,folderName);
  assert.equal(result.getName(), 'testFolder');

  var parentFol = DriveApp.getFilesByName('testFolder');
  var folderName = 'testFolder2';
  var result = checkFolder(parentFol,folderName);
  assert.equal(result.getName(), 'testFolder2');
}
function checkFolder(parentFolder,folderName) {
  /* 
  Check if folder exists, if not create it

  :param: parentFolder: parent folder
  :param: folderName: name of the folder to check

  :return: newFolder: folder created or folder already existing

  */
  do_create = true
  var allFolders = parentFolder.getFolders()
  while (allFolders.hasNext()) {
    var folder = allFolders.next();
    if (folder.getName() == folderName.toString()) {
      do_create = false
      return newFolder = folder
    }
  }
  if (do_create) {
    return newFolder = parentFolder.createFolder(folderName.toString())
  }
}


function testCheckFile() { 
  var assert = require("assert");
  var parentFol = DriveApp.getFilesByName('testFolder');
  var filename = 'testFile';  
  var result = checkFile(parentFol,filename);
  assert.equal(result, false);  
} 
function checkFile(parentFolder,filename) {
  /*
  Check if file exists

  :param: parentFolder: parent folder
  :param: filename: name of the file to check

  :return: results: true if file exists, false if not

  */
  var results;
  var haBDs = parentFolder.getFilesByName(filename)
  //Does not exist
  if (!haBDs.hasNext()) {
    results = haBDs.hasNext();
  }
  //Does exist
  else {
    results = haBDs.hasNext();
  }
  return results;
}