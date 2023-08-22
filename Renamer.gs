// Google Apps Script for batch folder name updates in Google Drive
// By: Alex O'Brien - Office of Research Services, UPEI

// Variables
let ourSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let folderId = "1qkc0PAf6xk_42X6vYN5O4ZmgWNnG_FSM"; // in this case, this folder is a grandparent to what we want to edit

// Function for performing batch changes to Google Drive folders
function mainFunction() {
  // Google Drive root
  var temp_merge_folder = DriveApp.getFolderById(folderId);
  // root contains researcher folders
  var researcher_folders = temp_merge_folder.getFolders();

  while(researcher_folders.hasNext()){
    // get next researcher folder
    var sub_folder = researcher_folders.next();
    // get researcher content folders
    var sub_folder_content = sub_folder.getFolders()
    
    while(sub_folder_content.hasNext()){
      var curFolder = sub_folder_content.next();
      // RENAME - pass in id of researcher content folders
      renameFoldersFunction(curFolder.getId());
    }
  }
}

// Function to execute folder name change
function renameFoldersFunction(folderId) {

  var curFolder = DriveApp.getFolderById(folderId);
  curFolder.setName("ORS - " + curFolder.getName())
  
}
