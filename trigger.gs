// Code implementing the trigger for the 2022/2023 seasons
//
// The goal of the trigger is to generate a new fresh sheet from
// template, performing some minimal validation of the input.
//
// 1- The user enters a family name
// 2- The family is upper-cased and prefixed with the sheet's name
//    which will carry the name of the operator (Xavier, Alex, etc...)
// 3- A new sheet is generated from the "FACTURE VIERGE <SEASON>"
//    template and stored in a directory create in the db/ directory
// 4- A link to that sheet is inserted back into the sheet from which
//    this script runs, so that the user can start directly interacting
//    with it.
// 5- Previous season data can be loaded and converted into a new
//    sheet.

// TODO:
//  - All FIXMEs
//  - On click on link replaces the status with a ready message

// Some globals defined here to make changes easy:
//
// The ID of the empty invoice to use to create content. Adjust
// this ID for the new season
var empty_invoice = '1enIvk0cW9RtsXLzVIZmwpK8HwSB89xAZOwcjNA2t0k0';

// The DB folder for the PREVIOUS season
var previous_db_folder = '1UmSA2OIMZs_9J7kEFu0LSfAHaYFfi8Et'
// Ranges to copy from an entry filed last season:
// FIXME: Season 2023/2024: second range features one more row.
//        and we will copy again up to column G.
// NOTE: Season 2022/2023: we're not copying the level so we stop
//       at F20
var ranges_previous_season = ["C8:G12", "B16:F20"];

// The DB folder for the CURRENT season
var db_folder = '1apITLkzOIkqCI7oIxpiKA5W_QR0EM3ey';
// Ranges to copy to for an entry filed this season:
// FIXME: Season 2023/2024: second range will feature one more row.
//        for the names: families will be up to 6 members and we
//        need to copy an additional column so include column G again.
var ranges_current_season = ["C6:G10", "B14:F18"];

var allowed_user = 'inscriptions.sca@gmail.com'

// Spreadsheet parameters (row, columns, etc...)
var coord_family_name = [1, 2];
var coord_family_last_season = [2, 2];
var coord_status_info = [3, 2];
var coord_download_link = [9, 2];

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

function setCellValueAtCoord(sheet, coord, value) {
  sheet.getRange(coord[0], coord[1]).setValue(value);  
}

// Return true when the active range is at coord.
function activeRangeInCoord(r, coord) {
  return r.getRow() == coord[0] && r.getColumn() == coord[1];
}

// If no data exists, read available family name from the previous
// season folder and use them to build a pull down menu that is used
// to populate a cell and the existing entry selection is reset.
// If data already exists, nothing happens, but the existing entry
// selection is reset.
//
// Loading the folder content doesn't work directly so it's better to
// just install the values in cell B2 as "Validation des données"
// "List d'éléments":
//
// 1- Export last year db-YYYY-YYYY as a zip file from the Drive UI.
// 2- Extract the list of registered families:
//
//    LC_CTYPE=C && LANG=C && \
//    unzip -l db-2021-2022-20220914T190110Z-001.zip | \
//    egrep db- | sed 's@^.*2022/@@g' | sort -u | sed 's/.*_//g' | \
//    sort > LIST
//
// 3- Edit LIST to remove undesirable entries
// 4- Turn the list in to a CSV list:
//
//    clear; for i in $(cat LIST); do echo -n "$i,"; done | \
//    sed 's/,$//g'; echo
//
//  5- Install in cell B2 as previously detailed
  
function setLastSeasonFamilyList(sheet) {
  return;

  // The rest of this code never runs and that's fine.  
  var x = coord_family_last_season[0];
  var y = coord_family_last_season[1];
  if (sheet.getRange(x, y).getDataValidation() != null) {
    clearRange(sheet, coord_family_last_season);
    return;
  }
	
  var last_season_entries = getFolderFolderNames(previous_db_folder);
  var last_season_list = [];
  for (var index in last_season_entries) {
    family = last_season_entries[index].split(":")[1];
    if (family != '' && family != undefined) {
      last_season_list.push(family)
    }
  }
  if (last_season_list.length != 0) {
    last_season_list.sort();
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(last_season_list, true).build();
    sheet.getRange(x, y).clearDataValidations()
      .clearContent().setDataValidation(rule);
  }
  clearRange(sheet, coord_family_last_season);
}

// Warn if the validated family name is already in db/ under any form
// (<operator>:<family_name>).  When something is found, return a
// filename/file_id tuple, otherwise return a tuple of empty strings.
function checkAlreadyExists(folder, family_name) {
  var files = DriveApp.getFolderById(folder).getFolders()
  while (files.hasNext()) {
    file = files.next();
    if (file.getName().split(':')[1] == family_name) {
      return [file.getName(), file.getId()];
    }
  }
  return ['', ''];
}

// Get the file names of all folders found inside a folder. Return an
// array of strings.
function getFolderFolderNames(folder) {
  var folder_list = [];
  var files = DriveApp.getFolderById(folder).getFolders();
  while (files.hasNext()) {
    file = files.next();
    folder_list.push(file.getName());
  }
  return folder_list;
}

// Retrieve and sanitize a family name at coord, return it.
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
      replace(/:/g, "-").   // : into -
      replace(/-+/g, "-").  // Many - into a single one
      replace(/-$/g, "").   // Last - into nothing
      replace(/^-/g, "")    // First - into nothing
  } else {
    return "";
  }
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

// Clear the family name input, set the list of families from last
// season and indicate that the system is ready.
function resetSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  clearRange(sheet, coord_family_name);
  setLastSeasonFamilyList(sheet);
  setRangeTextColor(sheet, coord_status_info,
                    "➡️ Entrer le nom d'une nouvelle famille ou selectionner " +
                    "une famille de la saison précédente", "green");
  SpreadsheetApp.flush();
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
  setRangeTextColor(sheet, coord_status_info, "⏳ ...", "orange");

  // Check whether the event happened entering a new family name.
  // If that value can be validated, indicate we might proceed.
  if (activeRangeInCoord(r, coord_family_name)) {
    // Do not allow a last season entry to be present at the same time.
    if (getFamilyName(sheet, coord_family_last_season) != '') {
      displayErrorPannel(sheet,
                         'Vous ne pouvez pas définir une nouvelle famille ' +
                         'et selectionner une famille sur la saison précédente ' +
                         'en même temps...')
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
                      "➡️ Cliquez sur 'Créer une nouvelle inscription' " +
                      "pour créer un dossier " + family_name,
                      "green");
    clearRange(sheet, coord_download_link);
    return
  }
  
  if (activeRangeInCoord(r, coord_family_last_season)) {
    // Do not allow an entered family name to be present at the same time.
    if (getFamilyName(sheet, coord_family_name) != '') {
      displayErrorPannel(sheet,
                         'Vous ne pouvez pas selectionner une famille ' +
                         'sur la saison préceedente ' +
                         'et définir une nouvelle famille en même temps...')
      clearRange(sheet, coord_status_info);      
      return;
    }
    var family_name = getFamilyName(sheet, coord_family_last_season);
    if (family_name == '') {
      clearRange(sheet, coord_status_info);      
      return;
    }
    setRangeTextColor(sheet, coord_status_info,
                      "➡️ Cliquez sur 'Créer une nouvelle inscription' " +
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

// When the sheet is opened or loaded, just reset its state
function onOpen(event) {
  resetSheet();
}

function createHyperLinkFromDocId(doc_id, link_text) {
  var url = ("https://docs.google.com/spreadsheets/d/" +
             doc_id + "/edit#gid=0");
  return '=HYPERLINK("' + url + '"; "' + link_text + '")';
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
  var link = createHyperLinkFromDocId(document_id, "Ouvrir " + final_name);
  x = coord_download_link[0];
  y = coord_download_link[1];
  sheet.getRange(x, y).setFormula(link);
  
  // Reset the family name, update the status bar.
  clearRange(sheet, coord_family_name);
  setRangeTextColor(sheet, coord_status_info,
		    "✅ Terminé - cliquez sur le lien en bas de cette page " +
		    "pour charger la nouvelle feuille", "green");
  return document_id;
}

function getSheetFromFolderID(folder_id, sheet_name) {
  var old_sheet_id = ''
  var files = DriveApp.getFolderById(folder_id).getFiles()
  while (files.hasNext()) {
    file = files.next();
    if (file.getName() == sheet_name) {
      old_sheet_id = file.getId();
      break;
    }
  }
  return old_sheet_id;
}

function createNewFamilySheetFromOld(sheet, family_name) {
  var new_sheet_id = createNewFamilySheet(sheet, family_name);
  var old = checkAlreadyExists(previous_db_folder, family_name);
  // Inside the old folder, search the sheet and determine its id.
  // old[0] is the (old) sheet name, old[1] is  the (old) folder id.
  var old_sheet_id = getSheetFromFolderID(old[1], old[0]);

  if (old_sheet_id == '') {
    displayErrorPannel(sheet, 
                       'Facture pour ' + family_name +
                       ' introuvable sur la saison précédente')
    return;
  }
  
  // Load the old and new sheets
  var old_sheet = SpreadsheetApp.openById(
      old_sheet_id).getSheetByName('Inscription');
  var new_sheet = SpreadsheetApp.openById(
      new_sheet_id).getSheetByName('Inscription');

  // Iterate over all the ranges to copy and copy the content from the
  // old to the new.
  if (ranges_previous_season.length != ranges_current_season.length) {
    displayErrorPannel(sheet, 
                       'Difference de range.length pour la famile ' + family_name +
                       ': previous=' + ranges_previous_season.length +
                       'current=' + ranges_current_season.length)
    return;
  }
  // Perform a blind copy on the first range which contains the 
  // civility object.
  // FIXME: This is fragile - maybe this should create
  // a map to name the ranges and pick the ranges by name instead
  // of relying on the magic of an index...
  var old_civility_range = old_sheet.getRange(ranges_previous_season[0]);
  var new_civility_range = new_sheet.getRange(ranges_current_season[0]);
  new_civility_range.setValues(old_civility_range.getValues());
  // Adjustments of cell D10
  // TODO: Remove this adjustment in 2023/2024
  setCellValueAtCoord(new_sheet, [10, 4], "secondaire");
  
  // Perform a copy the second range which contains the family
  // information. Uppercase all family names, convert M/F into
  // Garcon/Fille
  var old_member_range = old_sheet.getRange(ranges_previous_season[1]);
  var new_member_range = new_sheet.getRange(ranges_current_season[1]);
  const rows = old_member_range.getNumRows();
  const columns = old_member_range.getNumColumns();

  for (var row = 1; row <= rows; row++) {
    for (let column = 1; column <= columns; column++) {
      var source_cell = old_member_range.getCell(row, column);
      var dest_cell = new_member_range.getCell(row, column);
      // First column is the first name. We should remove
      // anything that is placed in parenthesis or after //
      if (column == 1) {
        var first_name = source_cell.getValue().toString();
        var found = first_name.indexOf('(');
        if (found >= 0) {
          first_name = first_name.substring(0, found);
        }
        var found = first_name.indexOf('/');
        if (found >= 0) {
          first_name = first_name.substring(0, found);
        }
        dest_cell.setValue(first_name)
      }
      // Second column in that range is the familly name which is
      // upercassed for consistency
      else if (column == 2) {
        dest_cell.setValue(source_cell.getValue().toString().toUpperCase());
      }
      // Fifth column is sex conversion
      else if (column == 5) {
        var sex = source_cell.getValue().toString().toUpperCase();
        if (sex == 'M') {
          dest_cell.setValue('Garçon');
        } else if (sex == 'F') {
          dest_cell.setValue('Fille');
        } else {
          dest_cell.setValue(sex);
        }
      } else {
        dest_cell.setValue(source_cell.getValue());
      }
    }
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
        "⚠️ Veuillez correctement saisir ou selectionner un " +
	    "nom de famille.\n\n" +
        "N'avez vous pas oublié de valider par [return] ou [enter] 😋 ?");
      return;
    }
    new_family = false;
  }
  
  var already_exists = checkAlreadyExists(db_folder, family_name);
  if (already_exists[0] != '') {
    var sheet_id = getSheetFromFolderID(already_exists[1], already_exists[0]);
    var extra_text = '';
    if (sheet_id) {
      extra_text = ("\n\nVous pouvez directement l'ouvrir en copiant le lien ci-dessous " +
                    "et en le collant dans la barre de navigation du navitateur web:\n\n" +
                    "https://docs.google.com/spreadsheets/d/" + sheet_id + "/edit#gid=0");
    }
    displayErrorPannel(
      sheet,
      "⚠️ Un dossier d'inscription existe déjà sous cette dénomination: " +
      already_exists[0] + extra_text);
    return;
  }
  
  // We have a valid family name, indicate that we're preparing the
  // data, clear the old download link.
  clearRange(sheet, coord_download_link);
  // If this is a new family, we force the cell value to be the
  // normalized name
  if (new_family) {
    setCellValueAtCoord(sheet, coord_family_name, family_name)
  }
  
  if (new_family) {
    setRangeTextColor(sheet, coord_status_info,
                      "⏳ Préparation de " + family_name + "...", "orange");
    SpreadsheetApp.flush();
    createNewFamilySheet(sheet, family_name);
  } else {
    setRangeTextColor(sheet, coord_status_info,
                      "⏳ Importation de " + family_name + "...", "orange");
    SpreadsheetApp.flush();    
    createNewFamilySheetFromOld(sheet, family_name);
  }
  setLastSeasonFamilyList(sheet);
  SpreadsheetApp.flush();
}
