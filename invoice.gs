// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.
// This script runs with duplicates of the following shared doc: 
// shorturl.at/EJM58

// Dev or prod? "dev" sends all email to email_dev. Prod is the
// real thing: family will receive invoices, and so will email_license,
// unless the trigger in use is the TEST trigger.
dev_or_prod = "prod"

// Enable/disable new features - first entry set to false
// requires all following entries set to false. Note that
// a validation is carried out immediately after these
// definitions.
var advanced_verification_family_licenses = true;
var advanced_verification_subscriptions = true;
var advanced_verification_skipass = true;

// Boolean expression generated from a truth table.
if ((!advanced_verification_subscriptions &&
     advanced_verification_skipass) ||
    (!advanced_verification_family_licenses &&
    advanced_verification_subscriptions)) {
  displayErrorPanel(
    "advanced_verification_family_licenses: " +
    advanced_verification_family_licenses +
    "\nadvanced_verification_subscriptions: " +
    advanced_verification_subscriptions +
    "\nadvanced_verification_skipass: " +
    advanced_verification_skipass);
}

// Seasonal parameters - change for each season
// 
// - Name of the season
var season = "2022/2023"
// - Date at which we consider a licensee is adult
var adult_date = new Date("2007-01-01");
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
var coord_family_civility = [6,  3]
var coord_family_name =     [6,  4]
var coord_family_email =    [9,  3]
var coord_cc =              [9,  5]
var coord_family_phone1 =   [10, 3]
var coord_family_phone2 =   [10, 5]
//
// - Locations of various status line and collected input, located
//   a the bottom of the invoice.
// 
var coord_personal_message = [78, 3]
var coord_callme_phone =     [79, 7]
var coord_timestamp =        [79, 2]
var coord_version =          [79, 3]
var coord_parental_consent = [79, 5]
var coord_status =           [81, 4]
var coord_generated_pdf =    [81, 6]
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
//
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
  'CN Jeune (Compétition)':  [2008, 2050],
  'CN Adulte (Compétition)': [1900, 2007]};
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
// - Coordinates of where subscription purchases are indicated.
//
var coord_purchased_subscriptions_non_comp = [
  [38, 5],  // Riders
  [39, 5],  // Child 1
  [40, 5],  // Child 2
  [41, 5],  // Child 3
  [42, 5]   // Child 4
];
//
// Skipass section:
//
// - Skipass values
//
var ski_pass_values = [
  'Famille 4',
  'Famille 5',
  'Senior',
  'Adulte',
  'Étudiant',
  'Junior',
  'Enfant',
  'Peluche'];
//
// - Skipass DoB validations per type. change the
//   start of end of ranges not featuring a negative number
//
var ski_pass_dob_validation = {
  'Famille 4': [-1,   2051],
  'Famille 5': [-1,   2051],
  'Senior':    [1900, 1958],
  'Adulte':    [1959, 2004],
  'Étudiant':  [1993, 2004],
  'Junior':    [2005, 2013],
  'Enfant':    [2014, 2016],
  'Peluche':   [2017, 2050]};
//
// - Skipass purchase indicator coordinates
//
var coord_purchased_ski_pass = {
  'Famille 4': [23, 5],
  'Famille 5': [24, 5],
  'Senior':    [25, 5],
  'Adulte':    [26, 5],
  'Étudiant':  [27, 5],
  'Junior':    [28, 5],
  'Enfant':    [29, 5],
  'Peluche':   [30, 5]};

// Email configuration - these shouldn't change very often
var allowed_user = 'inscriptions.sca@gmail.com'
var email_loisir = 'sca.loisir@gmail.com'
var email_comp = 'skicluballevardin@gmail.com'
var email_dev = 'apbianco@gmail.com'
var email_license = 'licence.sca@gmail.com'

///////////////////////////////////////////////////////////////////////////////
// Accessors to the data defined above
///////////////////////////////////////////////////////////////////////////////

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

function getSkiPassFamily4String() {
  return ski_pass_values[0];
}

function getSkiPassFamily5String() {
  return ski_pass_values[1];
}

function getSkiPassStudentString() {
  return ski_pass_values[4];
}

//
// There's nothing else to configure past that point
//

///////////////////////////////////////////////////////////////////////////////
// Utility methods controlling the execution environment
///////////////////////////////////////////////////////////////////////////////

function isProd() {
  return dev_or_prod == 'prod'
}

function isDev() {
  return dev_or_prod == 'dev'
}

function isTest() {
  return getOperator() == 'TEST'
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

///////////////////////////////////////////////////////////////////////////////
// Debugging methods
///////////////////////////////////////////////////////////////////////////////

// Issue a key/value string followed by a carriage return
function hashToString(v) {
  var to_return = "";
  for (const [key, value] of Object.entries(v)) {
    to_return += (key + ": " + value + "\n");
  }
  return to_return;
}

function arrayToString(v) {
  var to_return = "";
  for (var index in v) {
    to_return += index + ": " + v[index] + "\n";
  }
  return to_return;
}

function dobRangeToString(dob_start, dob_end) {
  var date_range = '';
  if (dob_end >= 2050) {
    date_range = "en " + dob_start + " et après";
  } else if (dob_start <= 1900) {
    date_range = "en " + dob_end + " et avant";
  } else {
    date_range = "de " + dob_start + " à " + dob_end;
  }
  return date_range;
}

function isEmpty(obj) {
  return Object.keys(obj).length === 0;
}

function Debug(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

///////////////////////////////////////////////////////////////////////////////
// Link management
///////////////////////////////////////////////////////////////////////////////

function createHyperLinkFromURL(url, link_text) {
  return '=HYPERLINK("' + url + '"; "' + link_text + '")';
}

function displayPDFLink(pdf_file, offset) {
  var link = createHyperLinkFromURL(pdf_file.getUrl(),
                                    "📁 Ouvrir " + pdf_file.getName())
  x = coord_generated_pdf[0]
  y = coord_generated_pdf[1]
  SpreadsheetApp.getActiveSheet().getRange(x, y).setFormula(link); 
}

///////////////////////////////////////////////////////////////////////////////
// PDF generation methods
///////////////////////////////////////////////////////////////////////////////

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
      + '&top_margin=0.1' + '&bottom_margin=0.1'
      + '&left_margin=0.1' + '&right_margin=0.1'           
      + '&sheetnames=false' + '&printtitle=true'
      + '&pagenum=CENTER' + '&gridlines=false'
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

///////////////////////////////////////////////////////////////////////////////
// Spreadsheet data access methods
///////////////////////////////////////////////////////////////////////////////

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

function getNumberAt(coord) {
  var x = coord[0];
  var y = coord[1];
  return Number(SpreadsheetApp.getActiveSheet().getRange(x, y).getValue())
}

///////////////////////////////////////////////////////////////////////////////
// Date of Birth (DoB) management
///////////////////////////////////////////////////////////////////////////////

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
      var yob = getDoBYear(dob_string);
      if (yob >= 1900 && yob <= 2050) {
        return dob_string;
      }
    }
  }
  return undefined;
}

// Determine whether someone is an adult. adult_date is a global that needs to
// be adjusted for each season.
function isAdult(dob) {
  // Converting dob to a date and using getters doesn't work very well
  // so we're parsing the date instead.
  var res = new RegExp("([0-9]+)/([0-9]+)/([0-9]+)", "gi").exec(dob);
  var anniversary = new Date(res[3]+"-"+res[2]+"-"+res[1]);
  // Use millisecond since epoch to determine whether someone was born
  // before or at the adult date, which qualifies the person as being
  // an adult.
  return anniversary.valueOf() <= adult_date.valueOf();
}

// Given a list of DoBs, count the number of adults and kids in it.
function countAdultsAndKids(dobs) {
  var count_adults = 0;
  var count_children = 0;
  for (var index in dobs) {
    if (isAdult(dobs[index])) {
      count_adults += 1;
    } else {
      count_children += 1;
    }
  }
  return [count_adults, count_children];
}

// Given a dd/MM/YYYY dob, return the YoB.
function getDoBYear(dob) {
  return Number(new RegExp("[0-9]+/[0-9]+/([0-9]+)", "gi").exec(dob)[1]);
}

///////////////////////////////////////////////////////////////////////////////
// Input formating
///////////////////////////////////////////////////////////////////////////////

// Reformat the phone numbers
function formatPhoneNumbers() {
  function formatPhoneNumber(coords) {
    var phone = getStringAt(coords);
    if (phone != '') {
      // Compress the phone number removing all spaces
      phone = phone.replace(/\s/g, "") 
      // Compress the phone number removing all hyphens
      phone = phone.replace(/-/g, "")
      // Insert replacing groups of two digits by the digits with a space
      var regex = new RegExp('([0-9]{2})', 'gi');
      phone = phone.replace(regex, '$1 ');
      setStringAt(coords, phone.replace(/\s$/, ""), "black");
    }
  }
  
  formatPhoneNumber(coord_family_phone1);
  formatPhoneNumber(coord_family_phone2);
}

///////////////////////////////////////////////////////////////////////////////
// Alerting
///////////////////////////////////////////////////////////////////////////////

function displayErrorPanel(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert("❌ Erreur:\n\n" + message, ui.ButtonSet.OK);
}

// Display a OK/Cancel panel, returns true if OK was pressed.
function displayYesNoPanel(message) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("⚠️ Attention:\n\n" +
      message, ui.ButtonSet.OK_CANCEL);
  return response == ui.Button.OK;
}

function displayWarningPanel(message) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("⚠️ Attention:\n\n" +
      message, ui.ButtonSet.OK);
}

///////////////////////////////////////////////////////////////////////////////
// Validation methods
///////////////////////////////////////////////////////////////////////////////

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

// Verify that family members have a first name, 
// last name, a DoB and a sex assigned to them.
// Return an error and also a list of collected DoBs.
function validateFamilyMembers() {
  function returnError(m) {
    return [m, []]
  }

  dobs = [];
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
      return returnError("Pas de nom de prénom fournit pour " + last_name);
    }
    if (! last_name) {
      return returnError("Pas de nom de famille fournit pour " + first_name);
    }
    // Upcase the familly name and write it back
    last_name = last_name.toUpperCase();
    setStringAt([coords_identity_rows[index], 3], last_name, "black");
    // We need a DoB but only if a license has been requested.
    var dob = getDoB([coords_identity_rows[index], 4]);
    var license = getStringAt([coords_identity_rows[index], 7]);
    if (dob == undefined) {
      if (license != '' && license != no_license) {
        return returnError(
          "Pas de date de naissance fournie pour " +
          first_name + " " + last_name +
          " ou date de naissance mal formatée (JJ/MM/AAAA)\n" +
          " ou année de naissance fantaisiste.");
      }
    } else {
      dobs.push(dob);
    }
    // We need a sex
    var sex = getStringAt([coords_identity_rows[index], 6]);
    if (sex != "Fille" && sex != "Garçon") {
      return returnError("Pas de sexe défini pour " +
                         first_name + " " + last_name);
    }
  }
  return ["", dobs];
}

// Loosely cross checks attributed licenses with delivered subscriptions...
function validateLicenseSubscription(attributed_licenses) {
  var total_non_comp = 0;
  var reached_zero = false;
  for (var index in coord_purchased_subscriptions_non_comp) {
    value = getNumberAt(coord_purchased_subscriptions_non_comp[index]);
    // When we get past the rider category, that value can only be 1
    // and when it reaches 0, it can't be 1 again :)
    if (index > 0) {
      if (value != 1 && value != 0) {
        return ("La valeur du champ 'Adhésion / Stage / Transport - enfant " +
                index + "' ne peut prendre que la valeur 0 ou 1.");
      }
      if (value == 1 && reached_zero) {
        return (
          "La valeur du champ 'Adhésion / Stage / Transport - enfant " +
          index +
          "' ne peut pas prendre la valeur 1 alors que la valeur du champ " +
          "'Adhésion / Stage / Transport - enfant " +
          (index-1) + "' est de 0.");
      }
      if (value == 0) {
        reached_zero = true;
      }
    }
    total_non_comp += value;
  }
  // The number of non competition registrations should be equal to the 
  // number of non competition CNs purchased
  var juniors = attributed_licenses[getNonCompJuniorLicenseString()];
  var adults = attributed_licenses[getNonCompAdultLicenseString()];
  var family = attributed_licenses[getNonCompFamilyLicenseString()];
  var total_attributed_licenses = juniors + adults + family;
  if (total_attributed_licenses != total_non_comp) {
    detail = ("• " + getNonCompJuniorLicenseString() + ": " + juniors + "\n" +
              "• " + getNonCompAdultLicenseString() + ": " + adults + "\n" +
              "• " + getNonCompFamilyLicenseString() + ": " + family + "\n" +
              "• TOTAL: " + total_attributed_licenses) + "\n";
    return ("Le nombre total de license(s) Loisir attribuée(s):\n\n" + detail + 
            "\nne correspond pas au nombre d'adhésion(s) Loisir saisie(s) " +
            "qui est de " + total_non_comp + ".\n\n" +
            "Une ou plusieurs personnes (enfants, adultes) prennent une " +
            "license mais pas d'adhésion?");
  }
  return "";
}

function validateSkiPassPurchase(dobs) {
  function numberOfDoBsInRange(dobs, min, max) {
    var count = 0;
    for (var index in dobs) {
      var yob = getDoBYear(dobs[index]);
      if (yob >= min && yob <= max) {
        count += 1;
      }
    }
    return count;
  }
  
  function getSkiPassTypeQuantity(ski_pass_type) {
    return getNumberAt(coord_purchased_ski_pass[ski_pass_type]);
  }

  var family4 = getSkiPassFamily4String();
  var family5 = getSkiPassFamily5String();
  // Verify the family passes first
  // For a family pass of four, verify that we have two adults and two childs
  family4_quantity = getSkiPassTypeQuantity(family4);
  if (family4_quantity != 0) {
    if (family4_quantity != 1) {
      return (family4_quantity +
              " forfaits Annuel Famille 4 personnes sélectionnés.");
      // Verify that we have two adults and two child
    }
    var res = countAdultsAndKids(dobs);
    var count_adults = res[0];
    var count_children = res[1];
    Debug("adults: "+count_adults+" kids:"+count_children);
    if (count_adults != 2 || count_children != 2) {
      return ("Seulement " + count_adults + " adulte(s) déclaré(s) et " +
              count_children + " enfants(s) de moins de 17 ans déclaré(s) " +
              "pour le forfait Annuel Famille 4 personnes sélectionné.");
    }
  }
  // For a family pass of five, verify that we have two adults and three childs
  family5_quantity = getSkiPassTypeQuantity(family5);
  if (family5_quantity != 0) {
    if (family5_quantity != 1) {
      return (family5_quantity +
              " forfaits Annuel Famille 5 personnes sélectionné");
    }
    var res = countAdultsAndKids(dobs);
    var count_adults = res[0];
    var count_children = res[1];
    if (count_adults != 2 || count_children != 3) {
      return ("Seulement " + count_adults + " adulte(s) déclaré(s) et " +
              count_children + " enfants(s) de moins de 17 ans déclaré(s) " +
              "pour le forfait Annuel Famille 5 personnes sélectionné.");
    }
  }
  
  var error = ""
  // If we validated a family pass, return now. The rest of the validation is
  // going to be too complicated.
  if (ski_pass_type == family4 || ski_pass_type == family5) {
    return error;
  }
  
  var student = getSkiPassStudentString();
  // Match the number of reserved ski pass entries for a given age range with
  // the number of entries in that age range.
  for (var index in ski_pass_values) {
    var ski_pass_type = ski_pass_values[index];
    // Family passes have already been verified
    if (ski_pass_type == family4 || ski_pass_type == family5) {
      continue;
    }
    var ski_pass_purchased = getSkiPassTypeQuantity(ski_pass_type);
    var dob_start = ski_pass_dob_validation[ski_pass_type][0];
    var dob_end = ski_pass_dob_validation[ski_pass_type][1]
    var dobs_found = numberOfDoBsInRange(dobs, dob_start, dob_end);
    if (dobs_found != ski_pass_purchased) {
      error += (ski_pass_purchased + " forfait(s) " + ski_pass_type +
                " acheté(s) pour " + dobs_found + " individu(s) né(s) " + 
                dobRangeToString(dob_start, dob_end));
      if (ski_pass_type == student) {
        error += " (Avez vous un adulte étudiant?)";
      }
      error += "\n";
    }
  }
  return error;
}

// Cross check the attributed licenses with the ones selected for payment
function validateLicenseCrossCheck() {
  function returnError(v) {
    return [v, {}];
  }

  var attributed_licenses = {}
  attributed_licenses_values.forEach(function(key) {
    attributed_licenses[key] = 0;
  });
  
  // Capture the "no license" and exec license string, we're going to
  // use it a lot.
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
        return returnError("'" + selected_license +
                           "' attribuée à un membre de famile inexistant!");
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
      return returnError("'" + selected_license +
                         "' n'est pas une license attribuée possible!");
    }
    // Executive license requires a city of birth
    if (selected_license == exec_license) {
      var city = getStringAt([row, 5]);
      if (city == '') {
        return returnError(first_name + " " + last_name + ": une license " +
                           selected_license +
                           " requiert de renseigner une ville et un pays de " +
                           "naissance");
      }
    }
    
    // If we don't have a license, we can stop now. No need to
    // validate the DoB
    if (selected_license == no_license) {
      continue;
    }
    
    // Validate DoB against the type of license
    var yob = getDoBYear(getDoB([row, 4]));
    dob_start = attributed_licenses_dob_validation[selected_license][0];
    dob_end = attributed_licenses_dob_validation[selected_license][1];
    if (yob < dob_start || yob > dob_end) {
      var date_range = dobRangeToString(dob_start, dob_end);
      return returnError(first_name + " " + last_name +
                         ": l'année de naissance " + yob +
                         " ne correspond aux années de validité de la " +
                         "license choisie: '" + selected_license + 
                         "', années de validité: '" + date_range + "'");
    }
  }
  
  // Collect the amount of purchased licenses into a hash
  attributed_licenses_values.forEach(function(key) {
    // Entry indicating no license is skipped because it can't
    // be collected.
    if (key != no_license) {
      purchased_licenses[key] = getNumberAt(coord_purchased_licenses[key]);
    }
  });
  
  // Perform the verification
  
  // First verify the family non comp license because it's special.
  // 1- Family license requires at last four declared family members
  //    selecting it
  // 2- One family license should have been purchased
  // 3- Other checks are carried out still as they can be determined
  //    valid or not
  //    even with a family license purchased.
  var family_license = getNonCompFamilyLicenseString();
  if (attributed_licenses[family_license] != 0) {
    // A least four declared participants
    if (attributed_licenses[family_license] < 4) {
      return returnError(
        "Il faut attribuer une licence famille à au moins 4 membres " +
        "d'une même famille. Seulement " + attributed_licenses[family_license] +
        " ont été attribuées.");
    }
    // Check that one family license was purchased.
    if (purchased_licenses[family_license] != 1) {
      return returnError(
        "Vous devez acheter une licence loisir famille, vous en avez " +
        "acheté pour l'instant " + purchased_licenses[family_license]);
    }
    // Last verification: two adults and two+ kids
    var res = countAdultsAndKids(dobs);
    var count_adults = res[0];
    var count_children = res[1];
    if (count_adults != 2 || count_children < 2) {
      return returnError(
        "Seulement " + count_adults + " adulte(s) déclaré(s) et " +
        count_children + " enfants(s) de moins de 17 ans déclaré(s) " +
        "pour la license famille.");
    }
  }
  for (var index in attributed_licenses_values) {
    // Entry indicating no license is skipped because it can't
    // be collected. Entry indicating a family license skipped because
    // it's been already verified
    var key = attributed_licenses_values[index];
    if (key != no_license && key != family_license) {
      if (purchased_licenses[key] != attributed_licenses[key]) {
        return returnError(
          "Le nombre de licence(s) '" + key + "' attribuée(s) (au nombre de " +
          attributed_licenses[key] + ")\n" +
          "ne correspond pas au nombre de licence(s) achetée(s) (au nombre de " +
          purchased_licenses[key] + ")");        
      }
    }
  }
  return ["", attributed_licenses];
}

///////////////////////////////////////////////////////////////////////////////
// Invoice meta information management
///////////////////////////////////////////////////////////////////////////////

function getOperator() {
  return SpreadsheetApp.getActiveSpreadsheet().getName().
    toString().split(':')[0]
}

function getFamilyName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName().
    toString().split(':')[1]
}

function getInvoiceNumber() {
  var version = getStringAt(coord_version);
  if (version == "") {
    version = "version=0";
  }
  var extracted_num = Number(
    new RegExp("version=([0-9]+)", "gi").exec(version)[1]);
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

// Produce a dictionary of family members that has purchased a license.
// This assumes that some verification of first/last name, DoB and sex
// have already been performed.
function getDictionaryOfFamilyPurchasingALicense() {
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

///////////////////////////////////////////////////////////////////////////////
// Top level invoice validation
///////////////////////////////////////////////////////////////////////////////

// Validate the invoice and return a dictionary of values
// to use during invoice generation.
function validateInvoice() {
  // Reformat the phone numbers  
  formatPhoneNumbers();

  if (! isProd()) {
    Debug("Cette facture est en mode developpement. " +
          "Aucun email ne sera envoyé, " +
          "ni à la famile ni à " + email_license + ".\n\n" +
          "Vous pouvez néamoins continuer et un dossier sera préparé et " +
          "les mails serons envoyés à " + email_dev + ".\n\n" +
          "Contacter " + email_dev + " pour obtenir plus d\'aide.")
          
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
  
  // Validate all the entered family members
  
  var ret = validateFamilyMembers();
  var family_validation_error = ret[0];
  var dobs = ret[1];
  if (family_validation_error) {
    displayErrorPanel(family_validation_error);
    return {};
  }
  
  // Now performing the optional/advanced validations... 
  //
  // 1- Validate the licenses requested by this family
  if (advanced_verification_family_licenses) {
    var ret = validateLicenseCrossCheck(dobs);
    var license_cross_check_error = ret[0]
    if (license_cross_check_error) {
      displayErrorPanel(license_cross_check_error);
      return {};
    }
  }

  // 2- Verify the subscriptions. The operator may choose to continue
  //    as some situation are un-verifiable automatically.
  if (advanced_verification_family_licenses &&
      advanced_verification_subscriptions) {
    var attributed_licenses_values = ret[1];
    // Validate requeted licenses and subscriptions
    var subscription_validation_error = 
        validateLicenseSubscription(attributed_licenses_values);
    if (subscription_validation_error) {
      subscription_validation_error += (
        "\n\nChoisissez 'OK' pour continuer à générer la facture.\n" +
        "Choisissez 'Annuler' pour ne pas générer la facture et " +
        "vérifier les valeurs saisies...");
      if (! displayYesNoPanel(subscription_validation_error)) {
        return {};
      }
    }
  }

  // 3- Verify the ski pass purchases. The operator may choose to continue
  //    as some situation are un-verifiable automatically.
  if (advanced_verification_family_licenses &&
      advanced_verification_subscriptions &&
      advanced_verification_skipass) {
    var skipass_validation_error = validateSkiPassPurchase(dobs);
    if (skipass_validation_error) {
      skipass_validation_error += (
        "\n\nChoisissez 'OK' pour continuer à générer la facture.\n" +
        "Choisissez 'Annuler' pour ne pas générer la facture et " +
        "vérifier les valeurs saisies...");
      if (! displayYesNoPanel(skipass_validation_error)) {
        return {};
      }    
    }
  }
  
  // Finally, validate the parental consent.
  var consent = validateAndReturnDropDownValue(
    coord_parental_consent,
    "Vous n'avez pas renseigné la nécessitée ou non de devoir " +
    "recevoir l'autorisation parentale.");
  if (consent == '') {
    return {}
  }
  var consent = getStringAt(coord_parental_consent);
  if (consent == 'Non fournie') {
    displayErrorPanel(
      "L'autorisation parentale doit être signée aujourd'hui pour " +
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

///////////////////////////////////////////////////////////////////////////////
// Invoice emailing
///////////////////////////////////////////////////////////////////////////////

function maybeEmailLicenseSCA(invoice) {
  var operator = getOperator()
  var family_name = getFamilyName()
  var family_dict = getDictionaryOfFamilyPurchasingALicense() 
  var string_family_members = "";
  for (var index in family_dict) {
    if (family_dict[index]['last'] == "") {
      continue
    }
    string_family_members += (
      "<tt>" +
      "Nom: <b>" + family_dict[index]['last'].toUpperCase() + "</b><br>" +
      "Prénom: " + family_dict[index]['first'] + "<br>" +
      "Naissance: " + family_dict[index]['birth'] + "<br>" +
      "Fille/Garçon: " + family_dict[index]['sex'] + "<br>" +
      "Ville de Naissance: " + family_dict[index]['city'] + "<br>" +
      "Licence: " + family_dict[index]['license'] + "<br>" +
      "----------------------------------------------------</tt><br>\n");
  }
  if (string_family_members) {
    string_family_members = (
      "<p>Licence(s) nécessaire(s) pour:</p><blockquote>\n" +
      string_family_members +
      "</blockquote>\n");
  } else {
    return;
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

function generatePDFAndMaybeSendEmail(send_email, just_the_invoice) {
  setStringAt(coord_status, 
              "⏳ Validation de la facture...", "orange")      
  SpreadsheetApp.flush()
  var validation = validateInvoice();
  if (isEmpty(validation)) {
    setStringAt(coord_status, 
                "❌ La validation de la facture a échouée", "red")      
    SpreadsheetApp.flush()
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
      attachments.push(
        DriveApp.getFileById(parental_consent_pdf).getAs(MimeType.PDF))
    
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

  // For a more precise quota computation, we would need to be able
  // to tell here that we're going to send an email to the license
  // email address.
  var quota_threshold = 2 + (cc_to == "" ? 0 : 1);
  if (send_email) {
    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    if (emailQuotaRemaining < quota_threshold) {
      displayWarningPanel("Quota email insuffisant. Envoi " +
                          (just_the_invoice ? "de la facture" : "du dossier") +
                          " retardé. Quota restant: " + emailQuotaRemaining +
                          ". Nécessaire: " + quota_threshold);
      setStringAt(coord_status,
                  "⚠️ Quota email insuffisant. Envoi retardé", "orange");
    } else {
      // Send the email  
      MailApp.sendEmail(email_options)
      maybeEmailLicenseSCA([attachments[0]]);

      setStringAt(coord_status, 
                  "✅ " + (just_the_invoice ? 
                           "Facture envoyée" : "dossier envoyé"), "green")  
    }
  } else {
      setStringAt(coord_status, 
                  "✅ " + (just_the_invoice ?
                           "Facture générée" : "dossier généré"), "green")  
  }
  displayPDFLink(pdf_file)
  SpreadsheetApp.flush()  
}

///////////////////////////////////////////////////////////////////////////////
// Button callbacks
///////////////////////////////////////////////////////////////////////////////

// Runs when the [secure authorization] button is pressed.
function GetAuthorization() {
  ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
}

// This is what the [generate and send folder] button runs.
function GeneratePDFAndSendEmailButton() {
  generatePDFAndMaybeSendEmail(/* send_email= */ true,
                               /* just_the_invoice= */ false)
}

// This is what the [generate invoice] button runs.
function GeneratePDFButton() {
  generatePDFAndMaybeSendEmail(/* send_email= */ false,
                               /* just_the_invoice= */ true)
}

// This is what the [generate and send only the invoice] button runs.
function GenerateJustPDFAndSendEmailButton() {
  generatePDFAndMaybeSendEmail(/* send_email= */ true,
                               /* just_the_invoice= */ true)
}

