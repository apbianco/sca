// Code implementing the trigger.
//
// The goal of the trigger is to generate a new fresh sheet from
// template, performing some minimal validation of the input.
//
// 1- The user enters a family name
// 2- The family is upper-cased and prefixed with the sheet's name
//    which will carry the name of the operator (Xavier, Alex, etc...)
// 3- A new sheet is generated from the "FACTURE VIERGE 2020/2021"
//    template and stored in a directory create in the db/ directory
// 4- A link to that sheet is inserted back into the sheet from which
//    this script runs, so that the user can start directly interact with
//    it.

// BUG TO FIX:
// - Create an already existing family from drop down
// - Error message appears
// - But content for row 1 and 2 isn't clear

// Some globals defined here to make changes easy:
//
// The ID of the empty invoice to use to create content. Adjust
// this ID for the new season
var empty_invoice = '1thTPqNLroAAaa5D82IUanJSQ_JY-NJkFMsJ4E56kuwo';

// The DB folder for the 2020/2021 season
var previous_db_folder = '1BDEV3PQULwsrqG3QjsTj1EGcnqFp6N6r'
// Ranges to copy from an entry filed last season:
var ranges_previous_season = ["C8:G12", "B16:G20"];

// The DB folder for the current season
var db_folder = '1UmSA2OIMZs_9J7kEFu0LSfAHaYFfi8Et';
// Ranges to copy to for an entry filed this season:
var ranges_current_season = ["C8:G12", "B16:G20"];

var allowed_user = 'inscriptions.sca@gmail.com'

// Spreadsheet parameters (row, columns, etc...)
var coord_family_name = [1, 2];
var coord_family_last_season = [2, 2];
var coord_status_info = [3, 2];
var coord_download_link = [9, 2];

var last_season_list = [
  "PETIT-BIANCO",
  "BADOUARD",
];

function Debug(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

function setRangeTextColor(sheet, coord, text, color) {
  var x = coord[0];
  var y = coord[1];
  sheet.getRange(x,y).setValue(text);
  sheet.getRange(x,y).setFontColor(color);
}

function clearRange(sheet, coord) {
  sheet.getRange(coord[0],coord[1]).clear();
}

function setLastSeasonFamilyList(sheet) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(
    last_season_list, true).build();
  var x = coord_family_last_season[0];
  var y = coord_family_last_season[1];
  sheet.getRange(x, y).clearDataValidations().clearContent().setDataValidation(rule);
}

// Retrieve and sanitize a family name.
function getFamilyName(sheet, coord) {
  var x = coord[0];
  var y = coord[1];
  if (sheet.getRange(x, y)) {
    return sheet.getRange(x, y).getValue().toString().toUpperCase().
      replace(/\s/g, "-").  // No spaces
      replace(/\d+/g, "").  // No numbers
      replace(/\//g, "-").  // / into -
      replace(/\./g, "-").  // . into -
      replace(/_/g, "-").   // _ into -
      replace(/-+/g, "-")   // Many - into a single one.
  } else {
    return "";
  }
} 

function setFamilyName(sheet, coord, value) {
  sheet.getRange(coord[0], coord[1]).setValue(value);  
}

// Warn if the validated family name is already in db/ under any form.
// When something is found, return a filename/file_id tuple, otherwise
// return a tuple of empty strings.
function checkAlreadyExists(folder, family_name) {
  var files = DriveApp.getFolderById(folder).getFolders()
  while (files.hasNext()) {
    file = files.next();
    if (file.getName().includes(family_name)) {
      return [file.getName(), file.getId()];
    }
  }
  return ['', ''];
}

function activeRangeInCoord(r, coord) {
  return r.getRow() == coord[0] && r.getColumn() == coord[1];
}

function resetSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  clearRange(sheet, coord_family_name);
  setLastSeasonFamilyList(sheet);
  setRangeTextColor(sheet, coord_status_info, "Créer ou importer une famille", "green");
}

// onEdit runs when the cell where you enter the family name sees its
// content changed. It serves two purposes:
//
// 1- Validate the selected content
// 2- Give direction in the status cell on what to do next.
function onEdit(event){
  var ui = SpreadsheetApp.getUi();
  var r = event.source.getActiveRange();
  var sheet = SpreadsheetApp.getActiveSheet();

  // Check whether the event happened entering a new family name.
  // If that value can be validated, indicate we might proceed.
  if (activeRangeInCoord(r, coord_family_name)) {
    // Do not allow a last season entry to be present at the same time.
    // FIXME: Display an error message
    if (getFamilyName(sheet, coord_family_last_season) != '') {
      clearRange(sheet, coord_status_info);      
      return;
    }
    var family_name = getFamilyName(sheet, coord_family_name);
    if (family_name == '') {
      clearRange(sheet, coord_status_info);      
      return;
    }
    setRangeTextColor(sheet, coord_family_name, family_name, "black");
    setRangeTextColor(sheet, coord_status_info,
                      "Cliquez sur 'Créer une nouvelle inscription' " +
                      "pour créer un dossier " + family_name,
                      "green");
    clearRange(sheet, coord_download_link);
    return
  }
  
  if (activeRangeInCoord(r, coord_family_last_season)) {
    // Do not allow an entered family name to be present at the same time.
    // FIXME: Display and error message
    if (getFamilyName(sheet, coord_family_name) != '') {
      clearRange(sheet, coord_status_info);      
      return;
    }
    var family_name = getFamilyName(sheet, coord_family_last_season);
    if (family_name == '') {
      clearRange(sheet, coord_status_info);      
      return;
    }
    setRangeTextColor(sheet, coord_status_info,
                      "Cliquez sur 'Créer une nouvelle inscription' " +
                      "pour importer le dossier de " + family_name,
                      "green");
    clearRange(sheet, coord_download_link);
    return    
  }
}

// onChange just calls on edit...
function onChange(event) {
  onEdit(event);
}

// When the sheet is opened or loaded, clear the family name,
// install the name of the old families in a drop-down and
// update the status bar.
function onOpen(event) {
  resetSheet();
}

// Display an error panel with some text, collect OK, 
// clear the status cell and the family input cell
function displayErrorPannel(sheet, message) {
  setRangeTextColor(sheet, coord_status_info, "⚠️ Erreur...", "red");
  SpreadsheetApp.flush();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(message, ui.ButtonSet.OK);
  resetSheet();
}

function createNewFamilySheet(sheet, family_name) {
  // This is the final name for the directory that will contain the
  // subscriptions
  var final_name = sheet.getName().toString() + ":" + family_name
    
  // Create a directory in db/<final_name>.
  var family_dir_id = DriveApp.createFolder(final_name).getId();
  DriveApp.getFileById(family_dir_id).moveTo(DriveApp.getFolderById(db_folder));
  
  // Make a copy of the template file and move it to the 'db/'
  // directory, then rename it.
  var document_id = DriveApp.getFileById(empty_invoice).makeCopy().getId();
  document_id = DriveApp.getFileById(document_id).moveTo(
    DriveApp.getFolderById(family_dir_id)).getId();  
  DriveApp.getFileById(document_id).setName(final_name);
  
  // Assemble the URL that leads to the new file and insert that in
  // the sheet... Change the readiness indicator.
  var url = "https://docs.google.com/spreadsheets/d/" +
	        document_id + "/edit#gid=0";
  var link = '=HYPERLINK("' + url + '"; "Ouvrir ' + final_name + '")'

  // Set the download link  
  x = coord_download_link[0];
  y = coord_download_link[1];
  sheet.getRange(x, y).setFormula(link);
  
  // Reset the family name, update the status bar.
  clearRange(sheet, coord_family_name);
  setRangeTextColor(sheet, coord_status_info,
		    "Terminé - cliquez sur le lien en bas de cette page " +
		    "pour charger la nouvelle feuille", "green");
  Logger.log("Created " + final_name + ", stored in " + url);
  Logger.log("User " + Session.getEffectiveUser());
  
  return document_id;
}

function createNewFamilySheetFromOld(sheet, family_name) {
  var new_sheet_id = createNewFamilySheet(sheet, family_name);
  Debug('new_sheet_id='+new_sheet_id);
  var old = checkAlreadyExists(previous_db_folder, family_name);
  var old_sheet_name = old[0];
  var old_folder_id = old[1];
  var old_sheet_id = '';
  
  // Inside the old folder, search the sheet and determine its id
  var files = DriveApp.getFolderById(old_folder_id).getFiles()
  while (files.hasNext()) {
    file = files.next();
    if (file.getName() == old_sheet_name) {
      old_sheet_id = file.getId();
      break;
    }
  }
  if (old_sheet_id == '') {
    // FIXME: Error message
    return;
  }
  
  // Load the old and new sheets
  var old_sheet = SpreadsheetApp.openById(old_sheet_id).getSheetByName('Inscription');
  var new_sheet = SpreadsheetApp.openById(new_sheet_id).getSheetByName('Inscription');

  // Iterate over all the ranges to copy and copy the content from the old to the new.
  if (ranges_previous_season.length != ranges_current_season.length) {
    // FIXME: Error message
    return;
  }
  // FIXME: Add verification for the old/new sheets
  for (var index in ranges_previous_season) {
    // FIXME: Verify ranges exist.
    var old_sheet_range = old_sheet.getRange(ranges_previous_season[index]);
    var new_sheet_range = new_sheet.getRange(ranges_current_season[index]);
    new_sheet_range.setValues(old_sheet_range.getValues());
  }
}

// This call back is attached to the button used to create a new entry.
function GenerateEntry() {
  // Get a handle on the current sheet, make sure only an allowed
  // user runs this.
  var sheet = SpreadsheetApp.getActiveSheet();
  if (Session.getEffectiveUser() != allowed_user) {
    displayErrorPannel(
      sheet,
      "⚠️ Vous n'utilisez pas cette feuille en tant que " +
      allowed_user + ".\n\n" +
      "Veuillez vous connecter d'abord à ce compte avant d'utiliser " +
      "cette feuille.");
    return;
  }
  
  // Clear the status bar.
  clearRange(sheet, coord_status_info);
  
  // Capture the sanitized family name. Warn if that doesn't work and return.
  var new_family = true;
  var family_name = getFamilyName(sheet, coord_family_name);
  if (family_name == '') {
    family_name = getFamilyName(sheet, coord_family_last_season);
    if (family_name == '') {
      displayErrorPannel(
        sheet,
        "⚠️ Veuillez correctement saisir ou selectionner un nom de famille.\n\n" +
        "N'avez vous pas oublié de valider par [return] ou [enter] 😋 ?");
      return;
    }
    new_family = false;
  }
  
  var already_exists = checkAlreadyExists(db_folder, family_name);
  if (already_exists[0] != '') {
      displayErrorPannel(
        sheet,
        "⚠️ Un dossier d'inscription existe déjà sous cette dénomination: " +
        already_exists[0]);
    return;
  }
  
  // We have a valid family name, indicate that we're preparing the data, clear
  // the old download link.
  clearRange(sheet, coord_download_link);
  // If this is a new family, we force the cell value to be the normalized name
  if (new_family) {
    setFamilyName(sheet, coord_family_name, family_name)
  }

  
  if (new_family) {
    setRangeTextColor(sheet, coord_status_info,
                      "Preparation de " + family_name + "...", "orange");
    SpreadsheetApp.flush();
    createNewFamilySheet(sheet, family_name);
  } else {
    setRangeTextColor(sheet, coord_status_info,
                      "Importation de " + family_name + "...", "orange");
    SpreadsheetApp.flush();    
    createNewFamilySheetFromOld(sheet, family_name);
  }
}
