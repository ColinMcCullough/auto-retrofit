var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetsCount = ss.getNumSheets();
var sheets = ss.getSheets();

function onOpen() { 
  // Try New Google Sheets method
  try{
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Retrofitting Tools')
    .addItem('Show Sheets', 'showSheets')
    .addItem('Hide External Sheets by Keyword', 'hideSheets')
    .addItem('Delete Sheets All Sheets From Sheet #', 'deleteAllSheetsFrom')
    .addItem('Delete Sheets', 'deleteSheets')
    .addItem('Copy All Visible Sheets', 'copyAllVisibleSheets') 
    .addItem('Copy Sheets By Keyword', 'copySheets') 
    .addToUi(); 
  } 
  
  // Log the error
  catch (e){Logger.log(e)}
  
  // Use old Google Spreadsheet method
  finally{
    var items = [
      {name: 'Hide Sheets', functionName: 'hideSheets'},
      {name: 'Show Sheets', functionName: 'showSheets'},
      {name: 'Delete Sheets All Sheets From Sheet #', functionName: 'deleteAllSheetsFrom'},
      {name: 'Delete Sheets', functionName: 'deleteSheets'},
      {name: 'Copy Sheets', functionName: 'copySheets'},
      
    ];
      ss.addMenu('Retrofitting Tools', items);
  }
}

function deleteSheets() {
  var deleteSheetsContaining = Browser.inputBox("Delete sheets with names containing:"); 
      deleteSheetsContaining = deleteSheetsContaining.toLowerCase();
    if (sheetMatch(deleteSheetsContaining)){
      for (var i = 0; i < sheetsCount; i++){
        var sheet = sheets[i]; 
        var sheetName = sheet.getName();
        sheetName = sheetName.toLowerCase();
        Logger.log(sheetName);
      if (sheetName.indexOf(deleteSheetsContaining.toString()) !== -1){
        Logger.log("DELETE!");
        ss.deleteSheet(sheet);
      }
    } 
  } else {
    noMatchAlert();
  }
}

function deleteAllSheetsFrom() {
  var deleteSheetsFrom = Browser.inputBox("Delete all sheet from sheet number:"); 
      for (var i = deleteSheetsFrom -1 ; i < sheetsCount; i++)
      {
      
          ss.deleteSheet(sheets[i]);
      }       
}


function hideSheets() {
  var hideSheetsContaining = Browser.inputBox("Hide sheets with names containing:");
  var destinationId = Browser.inputBox("Enter the destination spreadsheet ID:"); // update
  var destination = SpreadsheetApp.openById(destinationId);
  var extsheetsCount = destination.getNumSheets();
  var extsheets = destination.getSheets();
  if (extSheetMatch(hideSheetsContaining,extsheetsCount,extsheets)){
    for (var i = 0; i < sheetsCount; i++){
      var extsheet = extsheets[i]; 
      var extsheetName = extsheet.getName();
      if (extsheetName.indexOf(hideSheetsContaining.toString()) !== -1){
        renameAndArchive(destination,extsheet,extsheetName);  
      }
    }
  } else { 
    noMatchAlert();
  }
}

function renameAndArchive(destination,extsheet,sheetName) {
  var newName = sheetName + "(Archived)";
  extsheet.activate();
  destination.renameActiveSheet(newName);
  extsheet.hideSheet();
}

function showSheets() {
  var showSheetsContaining = Browser.inputBox("Show sheets with names containing:"); 
  if (sheetMatch(showSheetsContaining)){
    for (var i = 0; i < sheetsCount; i++){
      var sheet = sheets[i]; 
      var sheetName = sheet.getName();
      Logger.log(sheetName); 
      if (sheetName.indexOf(showSheetsContaining.toString()) !== -1){
        Logger.log("SHOW!");
        sheet.showSheet();
      }
    } 
  } else {
    noMatchAlert();
  }
}

function copyAllVisibleSheets() {
  var destinationId = Browser.inputBox("Enter the destination spreadsheet ID:");
  var destination = SpreadsheetApp.openById(destinationId);
  var extsheets = destination.getSheets();
    for (var i = 0; i < sheetsCount; i++){
      var sheet = sheets[i]; 
      var sheetName = sheet.getName();
      var destinationSheetName = destination.getSheetByName(sheetName);
      if (!sheet.isSheetHidden() && destinationSheetName != null){ //current spreadsheets tab is not hidden and destination sheet has identical name
        destinationArchive(destination,destinationSheetName,sheetName);
        copyToAndRename(sheet,destination,sheetName);
        }
      else if (!sheet.isSheetHidden() && destinationSheetName == null){
        copyToAndRename(sheet,destination,sheetName);
        }
      }
    successAlert('copied')
}

function copyToAndRename(sheet,destination,sheetName) {
  sheet.copyTo(destination);
  var destinationActive = destination.getSheetByName("Copy of " + sheetName);
  destinationActive.activate();
  destination.renameActiveSheet(sheetName);
}

function destinationArchive(destination,destinationMatch,sheetName) {
  var day = Utilities.formatDate(new Date(), "PST","MM/dd/yy");
  destinationMatch.activate();
  destination.renameActiveSheet(sheetName + "(Archived" + day + ")");
  destinationMatch.hideSheet();
}


function copySheets() {
  var copySheetsContaining = Browser.inputBox("Copy sheets with names containing:");
  var destinationId = Browser.inputBox("Enter the destination spreadsheet ID:");
  var destination = SpreadsheetApp.openById(destinationId);
  if (sheetMatch(copySheetsContaining)){
    for (var i = 0; i < sheetsCount; i++){
      var sheet = sheets[i]; 
      var sheetName = sheet.getName();
      var destinationMatch = destination.getSheetByName(sheetName);
      if (sheetName.indexOf(copySheetsContaining.toString()) !== -1 && destinationMatch != null){
        destinationArchive(destination,destinationMatch,sheetName);
        copyToAndRename(sheet,destination,sheetName);
      }
      else if (!sheet.isSheetHidden() && destinationMatch == null){
        copyToAndRename(sheet,destination,sheetName);
      }
    }
    successAlert('copied')
  } else {
    noMatchAlert();
  }
}



// determine if any sheets match the user input
function sheetMatch(sheetMatch){
  for (var i = 0; i < sheetsCount; i++){
    var sheetName = sheets[i].getName(); 
    //if(!sheets[i].isSheetHidden() && sheetName.indexOf(sheetMatch.toString()) !== -1) {
    if (sheetName.indexOf(sheetMatch.toString()) !== -1) {
      return true
    }
  }
  return false
}

// determine if any sheets match the user input
function extSheetMatch(sheetMatch,extsheetsCount,extsheets){
  for (var i = 0; i < extsheetsCount; i++){
    var sheetName = extsheets[i].getName(); 
    if (sheetName.indexOf(sheetMatch.toString()) !== -1){
      return true
    }
  }
  return false
}

// alert if no sheets matched the user input
function noMatchAlert() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'No Sheets Matched Your Input',
     "Try again and make sure you aren't using quotes.",
      ui.ButtonSet.OK);
}

// alert after succesful action (only used in copy)
function successAlert(action) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Success!',
     "You're sheets were " + action + " successfully.",
      ui.ButtonSet.OK);
}