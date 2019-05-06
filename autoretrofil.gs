var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetsCount = ss.getNumSheets();
var sheets = ss.getSheets();
var SIDEBAR_TITLE = "Auto Retro Fit Tool";

function onOpen() {
 SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Retrofit Tool', 'showSidebar')
      .addToUi();
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE);
  SpreadsheetApp.getUi().showSidebar(ui);
}

//test function to delete test sheets other than tab named home
function deleteAllExternalSheets() {
  var destinationId = Browser.inputBox("Enter the destination spreadsheet ID:"); // update
  var destination = SpreadsheetApp.openById(destinationId);
  var extsheetsCount = destination.getNumSheets();
  var extsheets = destination.getSheets();
  for (var i = 0; i < extsheetsCount; i++){
    var extsheet = extsheets[i]; 
    var extsheetName = extsheet.getName();
    if(extsheetName != "home") {
      destination.deleteSheet(extsheet);
    }
  }
} 
/*
  This function will delete any external sheets matching the exact keyword entered in the prompt
*/
function deleteSheetsByKeyword() {
  var deleteSheetsContaining = Browser.inputBox("Delete sheets with names containing:");
  var destinationId = Browser.inputBox("Enter the destination spreadsheet ID:"); // update
  var destination = SpreadsheetApp.openById(destinationId);
  var extsheetsCount = destination.getNumSheets();
  var extsheets = destination.getSheets();
  if (extSheetMatch(deleteSheetsContaining,extsheetsCount,extsheets)){
    for (var i = 0; i < extsheetsCount; i++){
      var extsheet = extsheets[i]; 
      var extsheetName = extsheet.getName();
      if (extsheetName.indexOf(deleteSheetsContaining.toString()) !== -1){
        destination.deleteSheet(extsheet);
      }
    }
    successAlert('deleted')
  } else { 
    noMatchAlert();
  }
}

/*
  This function will delete any internal sheets matching starting at the sheet number entered in the prompt
  EX: if you enter '2', it will delete all sheets after the first sheet
*/
function deleteAllSheetsFrom() {
  var deleteSheetsFrom = Browser.inputBox("Delete all sheet from sheet number:"); 
      for (var i = deleteSheetsFrom -1 ; i < sheetsCount; i++)
      {
          ss.deleteSheet(sheets[i]);
      }       
}

/*
  This function will hide all external sheets matching the exact keyword entered in the prompt.
  It will also rename them using the format (name + "archived" + date)
  If a matching archived name already exists it will rename the sheet being hidden with the time appended to the end of the name to not create a conflict
  EX: if you enter '2', it will delete all sheets after the first sheet
*/
function hideSheets() {
  var hideSheetsContaining = Browser.inputBox("Hide sheets with names containing:");
  var destinationId = Browser.inputBox("Enter the destination spreadsheet ID:"); // update
  var destination = SpreadsheetApp.openById(destinationId);
  var extsheetsCount = destination.getNumSheets();
  var extsheets = destination.getSheets();
  if (extSheetMatch(hideSheetsContaining,extsheetsCount,extsheets)){
    for (var i = 0; i < extsheetsCount; i++){
      var extsheet = extsheets[i]; 
      var extsheetName = extsheet.getName();
      if (extsheetName.indexOf(hideSheetsContaining.toString()) !== -1 && !extsheet.isSheetHidden()){
        //renameAndArchive(destination,extsheet,extsheetName); 
        destinationArchive(destination,extsheet,extsheetName);
      }
    }
    successAlert('hidden')
  } else { 
    noMatchAlert();
  }
}

/*
  @param destination: the destination Spreadsheet object
  @param extsheet: the destination sheet tab object
  @param sheetName: the name of the 
*/
function renameAndArchive(destination,extsheet,sheetName) {
  var newName = sheetName + "(Archived)";
  extsheet.activate();
  destination.renameActiveSheet(newName);
  extsheet.hideSheet();
}

/*
  This function will show all sheets with names matching the text entered in the prompt
*/
function showSheets() {
  var hideSheetsContaining = Browser.inputBox("Show sheets with names containing:");
  var destinationId = Browser.inputBox("Enter the destination spreadsheet ID:"); // update
  var destination = SpreadsheetApp.openById(destinationId);
  var extsheetsCount = destination.getNumSheets();
  var extsheets = destination.getSheets();
  if (extSheetMatch(hideSheetsContaining,extsheetsCount,extsheets)){
    for (var i = 0; i < extsheetsCount; i++){
      var extsheet = extsheets[i]; 
      var extsheetName = extsheet.getName();
      if (extsheetName.indexOf(hideSheetsContaining.toString()) !== -1 && extsheet.isSheetHidden()){
        //renameAndArchive(destination,extsheet,extsheetName); 
        extsheet.showSheet();
      }
    }
    successAlert('unhidden')
  } else { 
    noMatchAlert();
  }
}


/*
  This function will copy all visible sheets to the destination sheet
  If the destination sheet has the same tab name, it will hide it and rename appending 'Archived'
*/
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

/*
  @param sheet: Current Spreadsheet tab
  @param destination: Destination spreadsheet object
  @param sheetName: name of sheet tab being passed in
  Copies sheet to destination and renames the sheet to remove the prefix 'Copy of'
*/
function copyToAndRename(sheet,destination,sheetName) {
  sheet.copyTo(destination);
  var destinationActive = destination.getSheetByName("Copy of " + sheetName);
  destinationActive.activate();
  destination.renameActiveSheet(sheetName);
}

/*
  @param destination: Destination spreadsheet object
  @param destinationMatch: Destination spreadsheet Tab
  @param sheetName: Name of sheet
  This function hides and renames sheets
*/
function destinationArchive(destination,destinationMatch,sheetName) {
  var day = Utilities.formatDate(new Date(), "PST","MM/dd/yy");
  destinationMatch.activate();
  var archivedName = sheetName + "(Archived" + day + ")";
  var destinationArchiveMatch = destination.getSheetByName(archivedName);
  if(destinationArchiveMatch == null) { //does not find a sheet matching the standard archived naming convention
    destination.renameActiveSheet(sheetName + "(Archived" + day + ")");
  } else if (destinationArchiveMatch != null) { //finds a sheet matching the standard archived naming convention
    var formattedDate = Utilities.formatDate(new Date(), "PST","MM/dd/yy' Time:'hh:mm a");
    destination.renameActiveSheet(sheetName + "(Archived" + formattedDate + ")");
  }
  destinationMatch.hideSheet();
}

/*
  This function will copy all visible sheets matching the text entered in the prompt to the destination sheet
  If the destination sheet has the same tab name, it will hide it and rename appending 'Archived'
*/
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


/*
  @param sheetMatch: user input from prompt
  @return boolean: returns true if it finds a sheet matching the param text false if not
  determine if any sheets match the user input
*/
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

/*
  @param sheetMatch: user input from prompt
  @return boolean: returns true if it finds a sheet matching the param text false if not
  determine if any sheets match the user input on the destination sheet
*/
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