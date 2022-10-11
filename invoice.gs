// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.
// This script runs with duplicates of the following shared doc: 
// shorturl.at/EJM58

// Dev or prod? "dev" sends all email to email_dev. Prod is the
// real thing: family will receive invoices, and so will email_license,
// unless the trigger in use is the TEST trigger.
var dev_or_prod = "prod"

// Seasonal parameters - change for each season
// 
// - Name of the season
var season = "2022/2023"
//
// - Storage for the current season's database.
//
var db_folder = '1apITLkzOIkqCI7oIxpiKA5W_QR0EM3ey'
//
// - ID of attachements to be sent with the invoice - some may change
//   from one season to an other when they are refreshed.
//
// TODO: Change for 2022/2023
var parental_consent_pdf = '1TzWFmJUpp7eHdQTcWxGdW5Vze9XILw2e'
var rules_pdf = '1JKsqHWBIQc9PJrPesX3GkM9u22DSNVJN'
var parents_note_pdf = '10xRJwUWS_eJApxNgUJOfOAnAilNDdnWj'

// Spreadsheet parameters (row, columns, etc...). Adjust as necessary
// when the master invoice is modified.
// 
// - Locations of family details:
//
var coord_family_civility = [6, 3]
var coord_family_name = [6, 4]
var coord_family_email = [9, 3]
var coord_cc = [9, 5]
var coord_family_phone1 = [10, 3];
var coord_family_phone2 = [10, 5];
//
// - Locations of various status line and collected input, located
//   a the bottom of the invoice.
// 
var coord_personal_message = [78, 3]
var coord_callme_phone = [79, 7]
var coord_timestamp = [79, 2]
var coord_version = [79, 3]
var coord_parental_consent = [79, 5]
var coord_status = [81, 4]
var coord_generated_pdf = [81, 6]
//
// - Rows where the familly names are entered
// 
var coords_identity_rows = [14, 15, 16, 17, 18, 19];
//
// - Parameters defining the valid ranges to be retained during the
//   generation of the invoice's PDF
//
var coords_pdf_row_column_ranges = {'start': [1, 0], 'end': [80, 7]}
//
// - Definition of all possible license values
var attributed_licenses_values = [
  'Aucune',                    // Must match getNoLicenseString() accessor
  'CN Jeune (Loisir)',
  'CN Adulte (Loisir)',
  'CN Famille (Loisir)',
  'CN Dirigeant',              // Must match getExecutiveLicenseString() accessor
  'CN Jeune (Compétition)',
  'CN Adulte (Compétition)'];
//
// - DoB validation for a given type of license: change the
//   start of end of ranges not featuring a negative number
//
var attributed_licenses_dob_validation = {
  'Aucune':                  [-1,   2051],
  'CN Jeune (Loisir)':       [2007, 2050],
  'CN Adulte (Loisir)':      [1900, 2006],
  'CN Famille (Loisir)':     [-1,   2051],
  'CN Dirigeant':            [1900, 2004],
  'CN Jeune (Compétition)':  [2007, 2050],
  'CN Adulte (Compétition)': [1900, 2006]};
//
// Coordinates of where the various license purchases are indicated.
//
var coord_purchased_licenses = {
  // 'Aucune' doesn't exist.
  'CN Jeune (Loisir)':       [34, 5],
  'CN Adulte (Loisir)':      [35, 5],
  'CN Famille (Loisir)':     [36, 5],
  'CN Dirigeant':            [37, 5],
  'CN Jeune (Compétition)':  [44, 5],
  'CN Adulte (Compétition)': [45, 5]};
//
// Coordinates of where subscription purchases are indicated.
//
var coord_purchased_subscriptions_non_comp = [
  [38, 5],  // Riders
  [39, 5],  // Child 1
  [40, 5],  // Child 2
  [41, 5],  // Child 3
  [42, 5]   // Child 4
];

// Email configuration - these shouldn't change very often
var allowed_user = 'inscriptions.sca@gmail.com'
var email_loisir = 'sca.loisir@gmail.com'
var email_comp = 'skicluballevardin@gmail.com'
var email_dev = 'apbianco@gmail.com'
var email_license = 'licence.sca@gmail.com'

function isProd() {
  return dev_or_prod == 'prod'
}

function isDev() {
  return dev_or_prod == 'dev'
}

function isTest() {
  return getOperator() == 'TEST'
}

// Issue a key/value string followed by a carriage return
function hashToString(v) {
  var to_return = "";
  for (const [key, value] of Object.entries(v)) {
    to_return += (key + ": " + value + "\n");
  }
  return to_return;
}

// If we're in dev mode, return email_dev.
// If we're used from the TEST trigger, return allowed_user
// If not, return the provided email address.
// NOTE: dev mode takes precedence over TEST
function checkEmail(email) {
  if (isDev()) {
    return email_dev
  }
  if (isTest()) {
    return allowed_user
  }
  return email
}

function Debug(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

function createHyperLinkFromURL(url, link_text) {
  return '=HYPERLINK("' + url + '"; "' + link_text + '")';
}

function clearRange(sheet, coord) {
  sheet.getRange(coord[0],coord[1]).clear();
}

function setStringAt(coord, text, color) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var x = coord[0];
  var y = coord[1];
  sheet.getRange(x,y).setValue(text);
  sheet.getRange(x,y).setFontColor(color);
}

function getStringAt(coord) {
  var x = coord[0];
  var y = coord[1];
  return SpreadsheetApp.getActiveSheet().getRange(x, y).getValue().toString()
}

function getNoLicenseString() {
  return attributed_licenses_values[0];
}

function getNonCompJuniorLicenseString() {
  return attributed_licenses_values[1];
}

function getNonCompAdultLicenseString() {
  return attributed_licenses_values[2];
}

function getNonCompFamilyLicenseString() {
  return attributed_licenses_values[3];
}

function getExecutiveLicenseString() {
  return attributed_licenses_values[4];
}

function getCompJuniorLicenseString() {
  return attributed_licenses_values[5];
}

function getCompAdultLicenseString() {
  return attributed_licenses_values[6];
}

// Obtain a DoB at coords. Return an empty string if no DoB exists or if it
// can't be parsed properly as dd/MM/yyyy
function getDoB(coords) {
  var birth = getStringAt(coords);
  if (birth != "") {
    birth = new Date(getStringAt(coords));
    // Verify the format
    if (birth != undefined) {
      var dob_string = Utilities.formatDate(birth, "GMT", "dd/MM/yyyy");
      // Verify the date is reasonable
      var yob = Number(new RegExp("[0-9]+/[0-9]+/([0-9]+)", "gi").exec(dob_string)[1]);
      if (yob >= 1900 && yob <= 2050) {
        return dob_string;
      }
    }
  }
  return undefined;
}

function getFamilyDictionary() {
  var family = []
  var no_license = getNoLicenseString();
  for (var index in coords_identity_rows) {    
    var first_name = getStringAt([coords_identity_rows[index], 2]);
    var last_name = getStringAt([coords_identity_rows[index], 3]);    
    // After validation of the family entries, having a first name
    // guarantees a last name. Just check the existence of a
    // first name in order to skip that entry
    if (first_name == "") {
      continue;
    }
    // We can skip that entry if no license is required. That familly
    // member doesn't need to be reported in this dictionary.
    var license = getStringAt([coords_identity_rows[index], 7]);
    if (license == "" || license == no_license) {
      continue;
    }
    // DoB is guaranteed to be there if a license was requested
    var birth = getDoB([coords_identity_rows[index], 4]);
    var city = getStringAt([coords_identity_rows[index], 5])
    if (city == "") {
      city = "\\";
    }
    // Sex is guaranteed to be there
    var sex = getStringAt([coords_identity_rows[index], 6])

    family.push({'first': first_name, 'last': last_name,
                 'birth': birth, 'city': city, 'sex': sex,
                 'license': license})
  }
  return family
}

function savePDF(blob, fileName) {
  blob = blob.setName(fileName)
  var file = DriveApp.createFile(blob);
  pdfId = DriveApp.createFile(blob).getId()
  DriveApp.getFileById(file.getId()).moveTo(DriveApp.getFolderById(db_folder))
  return file;
}

function createPDF(sheet) {
  var ss_url = sheet.getUrl()
  var id = sheet.getSheetId()
  var url = ss_url.replace(/\/edit.*$/,'')  
      + '/export?exportformat=pdf'
      + '&format=pdf' + '&size=7' + '&portrait=true' + '&fitw=true'       
      + '&top_margin=0.1' + '&bottom_margin=0.1' + '&left_margin=0.1' + '&right_margin=0.1'           
      + '&sheetnames=false' + '&printtitle=true' + '&pagenum=CENTER' + '&gridlines=false'
      // Input range
      + '&fzr=false' + '&gid=' + id + '&ir=false' + '&ic=false'
      + '&r1=' + coords_pdf_row_column_ranges['start'][0]
      + '&c1=' + coords_pdf_row_column_ranges['start'][1]
      + '&r2=' + coords_pdf_row_column_ranges['end'][0]
      + '&c2=' + coords_pdf_row_column_ranges['end'][1]

  var params = {
    method: "GET",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()},
  }
  return UrlFetchApp.fetch(url, params).getBlob(); 
}

function displayErrorPanel(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

// Display a OK/Cancel pannel, returns true if OK was pressed.
function displayYesNoPanel(message) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(message, ui.ButtonSet.OK_CANCEL);
  return response == ui.Button.OK;
}

// Validate a cell at (x, y) whose value is set via a drop-down
// menu. We rely on the fact that a cell not yet set a proper value
// always has the same value.  When the value is valid, it is
// returned. Otherwise, '' is returned.
function validateAndReturnDropDownValue(coord, message) {
  var value = getStringAt(coord)
  if (value == 'Choix non renseigné' || value == '') {
    displayErrorPanel(message)
    return ''
  }
  return value
}

// Reformat the phone numbers
function formatPhoneNumbers() {
  function formatPhoneNumber(coords) {
    var phone = getStringAt(coords);
    if (phone != '') {
      // Compress the phone number removing all spaces
      phone = phone.replace(/\s/g, "") 
      // Insert replacing groups of two digits by the digits with a space
      var regex = new RegExp('([0-9]{2})', 'gi');
      phone = phone.replace(regex, '$1 ');
      setStringAt(coords, phone.replace(/\s$/, ""), "black");
    }
  }
  
  formatPhoneNumber(coord_family_phone1);
  formatPhoneNumber(coord_family_phone2);
}

// Verify that family members have a first name, 
// last name, a DoB and a sex assigned to them
function validateFamilyMembers() {
  var no_license = getNoLicenseString();
  for (var index in coords_identity_rows) {
    var first_name = getStringAt([coords_identity_rows[index], 2]);
    var last_name = getStringAt([coords_identity_rows[index], 3]);
    // Entry is empty, just skip it
    if (first_name == "" && last_name == "") {
      continue;
    }
    // We need both a first name and a last name
    if (! first_name) {
      return "Pas de nom de prénom fournit pour " + last_name;
    }
    if (! last_name) {
      return "Pas de nom de famille fournit pour " + first_name;
    }
    // Upcase the familly name and write it back
    last_name = last_name.toUpperCase();
    setStringAt([coords_identity_rows[index], 3], last_name, "black");
    // We need a DoB but only if a license has been requested.
    var dob = getDoB([coords_identity_rows[index], 4]);
    var license = getStringAt([coords_identity_rows[index], 7]);
    if (dob == undefined) {
      if (license != '' && license != no_license) {
        return ("Pas de date de naissance fournie pour " +
                first_name + " " + last_name +
                " ou date de naissance mal formatée (JJ/MM/AAAA)\n" +
                " ou année de naissance fantaisiste.");
      }
    }
    // We need a sex
    var sex = getStringAt([coords_identity_rows[index], 6]);
    if (sex != "Fille" && sex != "Garçon") {
      return "Pas de sexe défini pour " + first_name + " " + last_name;
    }
  }
  return "";
}

// Loosely cross checks attributed licenses with delivered subscriptions...
function validateLicenseSubscription(attributed_licenses) {
  var total_non_comp = 0;
  for (var index in coord_purchased_subscriptions_non_comp) {
    var row = coord_purchased_subscriptions_non_comp[index][0];
    var col = coord_purchased_subscriptions_non_comp[index][1];
    total_non_comp += Number(SpreadsheetApp.getActiveSheet().getRange(row, col).getValue());
  }
  // The number of non competition registrations should be equal to the 
  // number of non competition CNs purchased
  var juniors = attributed_licenses[getNonCompJuniorLicenseString()];
  var adults = attributed_licenses[getNonCompAdultLicenseString()];
  var family = attributed_licenses[getNonCompFamilyLicenseString()];
  var total_attributed_licenses = juniors + adults + family;
  if (total_attributed_licenses != total_non_comp) {
    detail = ("• \t" + getNonCompJuniorLicenseString() + ": " + juniors + "\n" +
              "• " + getNonCompAdultLicenseString() + ": " + adults + "\n" +
              "• " + getNonCompFamilyLicenseString() + ": " + family + "\n\n" +
              "• TOTAL: " + total_attributed_licenses) + "\n";
    return ("Le nombre total de license(s) loisir attribuée(s):\n\n" + detail + 
            "\nne correspond pas au nombre d'adhésion(s) loisr saisie(s) qui est de " +
            total_non_comp + ".");
  }
  return "";
}

// Cross check the attributed licenses with the ones selected for payment
function validateLicenseCrossCheck() {
  // First count how many licenses have been attributed to the
  // registered family. We use the values (after validation) directly

  function returnError(v) {
    return [v, {}];
  }

  var attributed_licenses = {}
  attributed_licenses_values.forEach(function(key) {
    attributed_licenses[key] = 0;
  });
  
  // Capture the no license and exec license string, we're going to use it a lot.
  var no_license = getNoLicenseString();
  var exec_license = getExecutiveLicenseString();
  var purchased_licenses = {}
  attributed_licenses_values.forEach(function(key) {
    // Entry indicating no license is skipped because it can't
    // be collected.
    if (key != no_license) {
      purchased_licenses[key] = 0;
    }
  });  
  
  // Collect the attributed licenses into a hash
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getStringAt([row, 7]);
    if (selected_license === '') {
      selected_license = no_license;
    }
    // You can't have no first/last name and an assigned license
    var first_name = getStringAt([row, 2]);
    var last_name = getStringAt([row, 3]);
    // If there's no name on that row, the only possible value is None
    if (first_name === '' && last_name === '') {
      if (selected_license != no_license) {
        return returnError("'" + selected_license + "' attribuée à un membre de famile inexistant!");
      }
      continue;
    }
    var found = false;
    attributed_licenses_values.forEach(function(key) {
      if (selected_license === key) {
        attributed_licenses[key] += 1;
        found = true;
        return;
      }
    });
    if (! found) {
      return returnError("'" + selected_license + "' n'est pas une license attribuée possible!");
    }
    // Executive license requires a city of birth
    if (selected_license == exec_license) {
      var city = getStringAt([row, 5]);
      if (city == '') {
        return returnError(first_name + " " + last_name + ": une license " + selected_license + "\n" +
                           "requiert de renseigner une ville et un pays de naissance");
      }
    }
    
    // If we don't have a license, we can stop now. No need to validate the DoB
    if (selected_license == no_license) {
      continue;
    }
    
    // Validate DoB against the type of license
    var dob = getDoB([row, 4]);
    dob = Number(new RegExp("[0-9]+/[0-9]+/([0-9]+)", "gi").exec(dob)[1]);
    dob_start = attributed_licenses_dob_validation[selected_license][0];
    dob_end = attributed_licenses_dob_validation[selected_license][1];
    if (dob < dob_start || dob > dob_end) {
      var date_range = '';
      if (dob_end >= 2050) {
        date_range = dob_start + " et après";
      } else if (dob_start <= 1900) {
        date_range = dob_end + " et avant";
      } else {
        data_range = "de " + dob_start + " à " + dob_end;
      }
      return returnError(first_name + " " + last_name + ": l'année de naissance " + dob + "\n" +
                         "ne correspond aux années de validité de la license choisie:\n'" +
                         selected_license + "': " + date_range);
    }
  }
  
  // Collect the amount of purchased licenses into a hash
  attributed_licenses_values.forEach(function(key) {
    // Entry indicating no license is skipped because it can't
    // be collected.
    if (key != no_license) {
      var row = coord_purchased_licenses[key][0];
      var col = coord_purchased_licenses[key][1];
      purchased_licenses[key] = Number(SpreadsheetApp.getActiveSheet().getRange(row, col).getValue());
    }
  });
  
  // Perform the verification
  
  // First verify the family non comp license because it's special.
  // 1- Family license requires at last four declared family members selecting it
  // 2- One family license should have been purchased
  // 3- Other checks are carried out still as they can be determined valid or not
  //    even with a family license purchased.
  var family_license = getNonCompFamilyLicenseString();
  if (attributed_licenses[family_license] != 0) {
    // A least four declared participants
    if (attributed_licenses[family_license] < 4) {
      return returnError("Il faut attribuer une license famille à au moins 4 membres " +
                         "d'une même famille. Seulement " + attributed_licenses[family_license] +
                         " ont été attribuées.");
    }
    // Check that one family license was purchased.
    if (purchased_licenses[family_license] != 1) {
      return returnError("Vous devez acheter une license loisir famille, vous en avez " +
                         "acheté pour l'instant " + purchased_licenses[family_license]);
    }
  }
  for (var index in attributed_licenses_values) {
    // Entry indicating no license is skipped because it can't
    // be collected. Entry indicating a family license skipped because
    // it's been already verified
    var key = attributed_licenses_values[index];
    if (key != no_license && key != family_license) {
      if (purchased_licenses[key] != attributed_licenses[key]) {
        return returnError("Le nombre de licences '" + key + "' selectionée(s) (au nombre de " +
                           attributed_licenses[key] + ")\n" +
                           "ne correspond pas au nombre de licences achetée(s) (au nombre de " +
                           purchased_licenses[key] + ")");        
      }
    }
  }
  return ["", attributed_licenses];
}

function isEmpty(obj) {
  return Object.keys(obj).length === 0;
}

function getOperator() {
  return SpreadsheetApp.getActiveSpreadsheet().getName().toString().split(':')[0]
}

function getFamilyName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName().toString().split(':')[1]
}

// Runs when the [secure authorization] button is pressed.
function GetAuthorization() {
  ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
}

function getInvoiceNumber() {
  var version = getStringAt(coord_version);
  if (version == "") {
    version = "version=0";
  }
  var extracted_num = Number(new RegExp("version=([0-9]+)", "gi").exec(version)[1]);
  if (extracted_num < 0) {
    displayErrorPanel("Problème lors de la génération du numéro de document\n" +
                      "Insérez version=99 en 79C et recommencez l'opération");
  }
  return extracted_num;
}

function getAndUpdateInvoiceNumber() {
  var extracted_num = getInvoiceNumber();
  extracted_num++;
  setStringAt(coord_version, "version=" + extracted_num, "black");
  SpreadsheetApp.flush();
  return extracted_num;
}

// Create the invoice as a PDF: first create a blob and then save
// the blob as a PDF and move it to the <db>/<OPERATOR:FAMILY>
// directory. Return the PDF file ID.  
function generatePDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var pdf_number = getAndUpdateInvoiceNumber();
  var blob = createPDF(spreadsheet)
  var pdf_filename = spreadsheet.getName() + '-' + pdf_number + '.pdf';
  var file = savePDF(blob, pdf_filename)
  
  var spreadsheet_folder_id =
    DriveApp.getFolderById(spreadsheet.getId()).getParents().next().getId()
  DriveApp.getFileById(file.getId()).moveTo(
    DriveApp.getFolderById(spreadsheet_folder_id))
  return file;
}

// Validate the invoice and return a dictionary of values
// to use during invoice generation.
function validateInvoice() {
  if (! isProd()) {
    Debug('Cette facture est en mode developpement. Aucun email ne sera envoyé, ' +
          'ni à la famile ni à ' + email_license + '.\n\n' +
          'Vous pouvez néamoins continuer et un dossier sera préparé et ' +
          'les mails serons envoyés à ' + email_dev + '.\n\n' +
          'Contacter ' + email_dev + ' pour obtenir plus d\'aide.')
          
  }
  
  // Make sure only an allowed user runs this.
  if (Session.getEffectiveUser() != allowed_user) {
    displayErrorPanel(
      "Vous n'utilisez pas cette feuille en tant que " +
      allowed_user + ".\n\n" +
      "Veuillez vous connecter d'abord à ce compte avant " +
      "d'utiliser cette feuille.");
    return {};
  }
  
  // Validation: proper civility
  var civility = validateAndReturnDropDownValue(
    coord_family_civility,
    "Vous n'avez pas renseigné de civilité")
  if (civility == '') {
    return {};
  }
  
  // Validation: a family name
  var family_name = getStringAt(coord_family_name)
  if (family_name == '') {
    displayErrorPanel(
      "Vous n'avez pas renseigné de nom de famille ou " +
      "vous avez oublié \n" +
      "de valider le nom de famille par [return] ou [enter]...")
    return {};
  }

  // Validation: proper email adress.
  var mail_to = getStringAt(coord_family_email)
  if (mail_to == '') {
    displayErrorPanel(
      "Vous n'avez pas saisi d'adresse email principale ou " +
      "vous avez oublié \n" +
      "de valider l'adresse email par [return] ou [enter]...")
    return {}
  }

  // Reformat the phone numbers  
  formatPhoneNumbers();

  // Validate all entered familly members
  var family_validation_error = validateFamilyMembers();
  if (family_validation_error) {
    displayErrorPanel(family_validation_error);
    return {};
  }

  // Validate requested licenses
  var ret = validateLicenseCrossCheck();
  var license_cross_check_error = ret[0]
  if (license_cross_check_error) {
    displayErrorPanel(license_cross_check_error);
    return {};
  }
  var attributed_licenses_values = ret[1];
  // Validate requeted licenses and subscriptions
  var subscription_validation_error = validateLicenseSubscription(attributed_licenses_values);
  if (subscription_validation_error) {
    subscription_validation_error += ("\n\nChoisissez 'OK' pour continuer à générer la facture.\n" +
                                      "Choisissez 'Annuler' pour ne pas générer la facture et vérifier les " + 
                                      "valeurs saisies...");
    if (! displayYesNoPanel(subscription_validation_error)) {
      return {};
    }
  }
  
  // Validate the parental consent.
  var consent = validateAndReturnDropDownValue(
    coord_parental_consent,
    "Vous n'avez pas renseigné la nécessitée ou non de devoir " +
    "recevoir l'autorisation parentale.");
  if (consent == '') {
    return {}
  }
  var consent = getStringAt(coord_parental_consent);
  if (consent == 'Non fournie') {
    displayErrorPanel("L'autorisation parentale doit être signée aujourd'hui pour " +
                      "valider le dossier et terminer l'inscription");
    return {};
  }
  
  // Update the timestamp. 
  setStringAt(coord_timestamp,
                    'Dernière MAJ le ' +
                    Utilities.formatDate(new Date(),
                                         Session.getScriptTimeZone(),
                                         "dd-MM-YY, HH:mm"), 'black')

  return {'civility': civility,
          'family_name': family_name,
          'mail_to': checkEmail(mail_to),
          'consent': consent};
}

function displayPDFLink(pdf_file, offset) {
  var link = createHyperLinkFromURL(pdf_file.getUrl(),
                                    "📁 Ouvrir " + pdf_file.getName())
  x = coord_generated_pdf[0]
  y = coord_generated_pdf[1]
  SpreadsheetApp.getActiveSheet().getRange(x, y).setFormula(link); 
}

// This is what the [generate] button runs
function GeneratePDFButton() {
  var validation = validateInvoice();
  if (isEmpty(validation)) {
    return;
  }
  setStringAt(coord_status, 
                    "⏳ Préparation de la facture...", "orange")  
  SpreadsheetApp.flush()
  displayPDFLink(generatePDF());
}

function maybeEmailLicenseSCA(invoice) {
  var operator = getOperator()
  var family_name = getFamilyName()
  var family_dict = getFamilyDictionary() 
  var string_family_members = "";
  for (var index in family_dict) {
    if (family_dict[index]['last'] == "") {
      continue
    }
    string_family_members += ("<tt>" +
                              "Nom: <b>" + family_dict[index]['last'].toUpperCase() + "</b><br>" +
                              "Prénom: " + family_dict[index]['first'] + "<br>" +
                              "Naissance: " + family_dict[index]['birth'] + "<br>" +
                              "Fille/Garçon: " + family_dict[index]['sex'] + "<br>" +
                              "Ville de Naissance: " + family_dict[index]['city'] + "<br>" +
                              "Licence: " + family_dict[index]['license'] + "<br>" +
                              "----------------------------------------------------</tt><br>\n");
  }
  if (string_family_members) {
    string_family_members = ("<p>Licence(s) nécessaire(s) pour:</p><blockquote>\n" +
                             string_family_members +
                             "</blockquote>\n");
  }
  
  var email_options = {
    name: family_name + ": nouvelle inscription",
    to: checkEmail(email_license),
    subject: family_name + ": nouvelle inscription",
    htmlBody:
      string_family_members +
      "<p>Dossier saisi par: " + operator + "</p>" +
      "<p>Facture (version " + getInvoiceNumber() + ") en attachement</p>" +
      "<p>Merci!</p>",
    attachments: invoice,
  } 
  MailApp.sendEmail(email_options)
}

// This is what the [generate and send folder] button runs.
function GeneratePDFAndSendEmailButton() {
  generatePDFAndMaybeSendEmail(/* send_email= */ true, /* just_the_invoice= */ false)
}

// This is what the [generate invoice] button runs.
function GeneratePDFButton() {
  generatePDFAndMaybeSendEmail(/* send_email= */ false, /* just_the_invoice= */ true)
}

// This is what the [generate and send only the invoice] button runs.
function GenerateJustPDFAndSendEmailButton() {
  generatePDFAndMaybeSendEmail(/* send_email= */ true, /* just_the_invoice= */ true)

}

function generatePDFAndMaybeSendEmail(send_email, just_the_invoice) {
  var validation = validateInvoice();
  if (isEmpty(validation)) {
    return;
  }
  setStringAt(coord_status, 
                    "⏳ Préparation de la facture...", "orange")
  SpreadsheetApp.flush()
  
  // Generate and prepare attaching the PDF to the email
  var pdf_file = generatePDF();
  var pdf = DriveApp.getFileById(pdf_file.getId());
  var attachments = [pdf.getAs(MimeType.PDF)]

  if (just_the_invoice) {
    setStringAt(coord_status,
                      "⏳ Génération " +
                      (send_email? "et envoit " : " ") + "de la facture...", "orange");
  } else {
    setStringAt(coord_status, 
                      "⏳ Génération " +
                      (send_email? "et envoit " : " ") + "du dossier...", "orange");
    SpreadsheetApp.flush()
  }
  
  var civility = validation['civility'];
  var family_name = validation['family_name'];
  var mail_to = validation['mail_to'];
  var consent = validation['consent'];
  
  // Determine whether parental consent needs to be generated. If
  // that's the case, we generate additional attachment content.
  var parental_consent_text = ''
  if (! just_the_invoice) {
    if (consent == 'À fournire') {
      attachments.push(DriveApp.getFileById(parental_consent_pdf).getAs(MimeType.PDF))
    
      parental_consent_text = (
        "<p>Il vous faut compléter, signer et nous retourner l'autorisation " +
        "parentale fournie en attachment, couvrant le droit à l'image, le " +
        "règlement intérieur et les interventions médicales.</p>");
    }

    // Insert the note and rules for the parents anyways
    attachments.push(DriveApp.getFileById(rules_pdf).getAs(MimeType.PDF))
    attachments.push(DriveApp.getFileById(parents_note_pdf).getAs(MimeType.PDF));
    parental_consent_text += (
      "<p>Vous trouverez également en attachement une note adressée aux " +
      "parents, ainsi que le règlement intérieur. Merci de lire ces deux " +
      "documents attentivement.</p>");
  }
  
  var subject = ("❄️ [Incription Ski Club Allevardin] " +
                 civility + ' ' + family_name +
		         ": Facture pour la saison " + season)

  // Collect the personal message and add it to the mail
  var personal_message_text = getStringAt(coord_personal_message)
  if (personal_message_text != '') {
      personal_message_text = ('<p><b>Message personnel:</b> ' +
                               personal_message_text + '</p>')
  }
  
  // Fetch a possible phone number
  var phone = getStringAt(coord_callme_phone)
  if (phone != 'Aucun') {
    phone = '<p>Merci de contacter ' + phone + '</p>'
  } else {
    phone = ''
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var email_options = {
    name: 'Ski Club Allevardin, gestion des inscriptions',
    // During validation, mail_to can be replaced by the dev email if not
    // running as prod.
    to: mail_to,
    subject: subject,
    htmlBody:
      "<h3>Bonjour " + civility + " " + family_name + "</h3>" +
    
      personal_message_text + 
      phone +

      "<p>Votre facture pour la saison " + season + " " +
      "est disponible en attachement. Veuillez contrôler " +
      "qu\'elle corresponde à vos besoins.</p>" +
    
      parental_consent_text +

      "<p>Des questions concernant cette facture? Contacter Aissam: " +
      "aissam.yaagoubi@sfr.fr (06-60-50-74-77) pour le ski loisir ou " +
      "Ludivine: tresorerie.sca@gmail.com pour le ski compétition.</p>" +
      "<p>Des questions concernant la saison " + season + " ? " +
      "Envoyez un mail à " + email_loisir + " (ski loisir) " +
      "ou à " + email_comp + " (ski compétition)</p>" +
    
      "<p>Nous vous remercions de la confiance que vous nous accordez " +
      "cette année.</p>" +
    
      "~SCA ❄️ 🏔️ ⛷️ 🏂",
    attachments: attachments
  }

  // Add CC if defined.
  var cc_to = getStringAt(coord_cc)
  if (cc_to != "") {
    email_options.cc = checkEmail(cc_to)
  }

  if (send_email) {
    // Send the email  
    MailApp.sendEmail(email_options)
    maybeEmailLicenseSCA([attachments[0]]);

    setStringAt(coord_status, 
                      "✅ Dossier envoyé", "green")  
  } else {
    setStringAt(coord_status, 
                      "✅ Dossier généré", "green")  
  }
  displayPDFLink(pdf_file)
  SpreadsheetApp.flush()  
}
