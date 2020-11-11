// Code implementing the trigger.
//
// The goal of the trigger is to generate a new fresh sheet from template, performing some minimal
// validation of the input.
//
// 1- The user enters a family name
// 2- The family is upper-cased and prefixed with the sheet's name which will carry the name of the
//    operator (Xavier, Alex, etc...)
// 3- A new sheet is generate from the "FACTURE VIERGE 2020/2021" template and stored in the db/
//    directory
// 4- A link to that sheet is inserted back into the sheet from which this script runs, so that the
//    user can start directly interact with it.

// Some globals defined here to make changes easy:
var empty_invoice = '1wh8HadLsEvbUOg00d7wfQt8MMpudBmwZceqSytWF5Ho'
var db_folder = '1BDEV3PQULwsrqG3QjsTj1EGcnqFp6N6r'

function setRangeTextColor(sheet, x, y, text, color) {
  sheet.getRange(x,y).setValue(text);
  sheet.getRange(x,y).setFontColor(color);
}

function clearRange(sheet, x, y) {
  sheet.getRange(x,y).clear();
}

function getFamilyName(sheet) {
  return sheet.getRange(1,2).getValue().toString().toUpperCase().replace(/\s/g, "");
}

// onEdit runs when the cell where you enter the family name sees its content changed.
// It serves two purposes:
// 1- Validate the content
// 2- Give direction in the status cell on what to do next.
function onEdit(event){
  var r = event.source.getActiveRange();
  // If we're entering a non empty family name, upper-case it and indicate that
  // we might proceed.
  if (r.getRow() == 1 && r.getColumn() == 2) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var family_name = ''
    if (sheet.getRange(1,2)) {
      var family_name = getFamilyName(sheet);
    }
    if (family_name == '') {
      clearRange(sheet, 2, 2);      
      return;
    }
    setRangeTextColor(sheet, 1, 2, family_name, "black");
    setRangeTextColor(sheet, 2, 2, "Cliquez sur 'Créer une nouvelle inscription'...","green");
    clearRange(sheet, 8, 2);
  }
}

// onChange just calls on edit...
function onChange(event) {
  onEdit(event);
}

function DisplayErrorPannel(sheet, message) {
  setRangeTextColor(sheet, 2,2, "Erreur", "red");
  SpreadsheetApp.flush();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(message, ui.ButtonSet.OK);
  clearRange(sheet, 2, 2);
}

// This call back is attached to the button used to create a new entry.
function GenerateEntry() {
  // Get a handle on the current sheet, clear old data...
  var sheet = SpreadsheetApp.getActiveSheet();
  clearRange(sheet, 2, 2);

  // Make sure only skicluballevardin@gmail.com runs this.
  if (Session.getEffectiveUser() != 'skicluballevardin@gmail.com') {
    DisplayErrorPannel(
      sheet,
      "Vous n'utilisez pas cette feuille en tant que skicluballevardin@gmail.com.\n\n" +
      "Veuillez vous connecter d'abord à ce compte avant d'utiliser cette feuille.");
    return;
  }
  
  // Capture the family name, transform to upper case. Warn if that doesn't work and return.
  var family_name = getFamilyName(sheet);
  if (family_name == '') {
    DisplayErrorPannel(
      sheet,
      "Veuillez correctement saisir un nom de famille.\n\n" +
      "N'avez vous pas oublié de valider par [return] ou [enter] 😋 ?")
    return;
  }
  // We have a valid family name, indicate that we're preparing the data, clear
  // the old download link.
  clearRange(sheet, 8, 2);
  sheet.getRange(1,2).setValue(family_name);
  setRangeTextColor(sheet, 2, 2, "Preparation de " + family_name+"...", "orange");
  SpreadsheetApp.flush();

  // Make a copy of the template file and move it to the 'db/' directory, then rename it.
  var documentId = DriveApp.getFileById(empty_invoice).makeCopy().getId();
  documentId = DriveApp.getFileById(documentId).moveTo(DriveApp.getFolderById(db_folder)).getId();  
  var final_name = sheet.getName().toString() + ":" + family_name
  DriveApp.getFileById(documentId).setName(final_name);
  
  // Assemble the URL that leads to the new file and insert that in the sheet... Change the
  // readiness indicator.
  var url = "https://docs.google.com/spreadsheets/d/" + documentId + "/edit#gid=0";
  var link = '=HYPERLINK("' + url + '"; "Ouvrir ' + final_name + '")'
  sheet.getRange(8,2).setFormula(link);
  clearRange(sheet, 1, 2);
  setRangeTextColor(sheet, 2, 2, "Terminé - cliquez sur le lien en bas de cette page pour charger la nouvelle feuille", "green");

  Logger.log("Created " + final_name + ", stored in " + url);
  Logger.log("User " + Session.getEffectiveUser())
}
