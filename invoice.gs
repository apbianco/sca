// Version: 2023-10-06
// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.

// Dev or prod? "dev" sends all email to email_dev. Prod is the
// real thing: family will receive invoices, and so will email_license,
// unless the trigger in use is the TEST trigger.
dev_or_prod = "prod"

// Enable/disable new features - first entry set to false
// requires all following entries set to false. Note that
// a validation of that constraint is carried out immediately
// after these definitions.
var advanced_verification_family_licenses = true;
var advanced_verification_subscriptions = true;
// FIXME: Too complicated for the 6th of October
var advanced_verification_skipass = false;

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
var season = "2023/2024"
//
// - Storage for the current season's database.
//
var db_folder = '1vTYVaHHs1oRvdbQ3mvmYfUvYGipXPaf3'
//
// - ID of attachements to be sent with the invoice - some may change
//   from one season to an other when they are refreshed.
//
// Level aggregation trix to update when a new entry is added
//
var license_trix = '13akc77rrPaI6g6mNDr13FrsXjkStShviTnBst78xSVY'
//
// PDF content to insert in a registration bundle.
// TODO: Might need to change for 2023/2024 if the content changes
//       after validation by the folks in charge.
//
var parental_consent_pdf = '1y68LVW5iZBSlRTEOIqM5umJFFi6o3ZCP'
var rules_pdf = '10zOpUgU0gt8qYpsLBJoJQT5RyBCdeAm1'
var parents_note_pdf = '1RewmJD4EvDJYUW0DXN0o36LTAA7r6n6-'
var information_leaflet_pdf = '1jpclCIoqu0eNh8fhwY0kklCNzH2d5EYO'
// The page in information_leaflet_pdf parents should sign
var information_leaflet_page = 16

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
var coord_rebate =           [76, 4]
var coord_personal_message = [85, 3]
var coord_timestamp =        [86, 2]
var coord_version =          [86, 3]
var coord_parental_consent = [86, 5]
var coord_medical_form =     [86, 7]
var coord_callme_phone =     [86, 9]
var coord_status =           [88, 4]
var coord_generated_pdf =    [88, 6]
//
// - Rows where the family names are entered
// 
var coords_identity_rows = [14, 15, 16, 17, 18, 19];
//
// - Columns where informationa about family members can be found
//
var coord_first_name_column = 2
var coord_last_name_column = 3
var coord_dob_column = 4
var coord_cob_column = 5
var coord_sex_column = 6
var coord_level_column = 7
var coord_license_column = 8
var coord_license_number_column = 9
//
// - Parameters defining the valid ranges to be retained during the
//   generation of the invoice's PDF
//
var coords_pdf_row_column_ranges = {'start': [1, 0], 'end': [86, 9]}

// Validation method use when creating license, skipasses and subscription
// maps.
function validateClassInstancesMap(map, map_name) {
  for (var key in map) {
    if (map[key].Name() != key) {
      displayErrorPanel('Erreur interne de validation pour ' + map_name + ': ' +
                        map_name + '[' + key + '].Name() = ' + map[key].Name())
    }
  }
}

//
// - Definition of all possible license values, the license class and
//   all possible instances.
//
function getNoLicenseString() {return 'Aucune'}
function getNonCompJuniorLicenseString() { return 'CN Jeune (Loisir)'}
function getNonCompAdultLicenseString() { return 'CN Adulte (Loisir)'}
function getNonCompFamilyLicenseString() {return 'CN Famille (Loisir)'}
function getExecutiveLicenseString() {return 'CN Dirigeant'}
function getCompJuniorLicenseString() {return 'CN Jeune (Compétition)'}
function getCompAdultLicenseString() {return 'CN Adulte (Compétition)'}

// Get a license at coord and normalize it's value. It's important to normalize
// the license string as it can be used as a key to a license_map.
function getLicenseAt(coord) {
  var license = getStringAt(coord)
  if(isLicenseNotDefined(license)) {
    license = getNoLicenseString()
  }
  return license
}

class License {
  constructor(name, purchase_range, dob_validation_method, valid_dob_range_message) {
    // The name of the ski pass type
    this.name = name
    // The range at which the number of licenses of the same kind can be found
    this.purchase_range = purchase_range
    // The DoB validation method
    this.dob_validation_method = dob_validation_method
    this.valid_dob_range_message = valid_dob_range_message
    this.occurence_count = 0
    this.purchased_amount = 0
  }

  Name() { return this.name }

  // When this method run, we capture the amount licenses of that given type
  // the operator entered.
    UpdatePurchasedLicenseAmount() {
    if (this.purchase_range != null) {
      this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                           this.purchase_range.getColumn()])
    }
  }
  PurchasedLicenseAmount() { return this.purchased_amount }

  IncrementAttributedLicenseCount() {
    this.occurence_count += 1
  }
  AttributedLicenseCount() { return this.occurence_count }

  ValidateDoB(dob) {
    if (this.dob_validation_method != null) {
      return this.dob_validation_method(dob)
    } else {
      return false
    }
  }

  ValidDoBRangeMessage() {
    return this.valid_dob_range_message
  }
}
  
function createLicensesMap(sheet) {
  var to_return = {
    'Aucune': new License(
      getNoLicenseString(), null, null),
    'CN Jeune (Loisir)': new License(
      getNonCompJuniorLicenseString(),
      sheet.getRange(41, 5),
      // ageVerificationBornAfter is inclusive - so born after 2009 is born on 1/1/2009 or later.
      (dob) => {return ageVerificationBornAfter(dob, new Date("January 1, 2009"))},
      "2009 et après"),
    'CN Adulte (Loisir)': new License(
      getNonCompAdultLicenseString(),
      sheet.getRange(42, 5),
      // ageVerificationBornBefore is inclusive - so born on 2008 and before means 12/1/2008 or before.
      (dob) => {return ageVerificationBornBefore(dob, new Date("December 31, 2008"))},
      "2008 et avant"),
    'CN Famille (Loisir)': new License(
      getNonCompFamilyLicenseString(),
      sheet.getRange(43, 5),
      (dob) => {return true},
      ""),
    'CN Dirigeant': new License(
      getExecutiveLicenseString(),
      sheet.getRange(44, 5),
      (dob) => {return isAdult(dob)},
      "adulte"),
    'CN Jeune (Compétition)': new License(
      getCompJuniorLicenseString(),
      sheet.getRange(51, 5),
      (dob) => {return ageVerificationBornAfter(dob, new Date("January 1, 2009"))},
      "2009 et après"),
    'CN Adulte (Compétition)': new License(
      getCompAdultLicenseString(),
      sheet.getRange(52, 5),
      (dob) => {return ageVerificationBornBefore(dob, new Date("December 31, 2008"))},
      "2008 et avant"),
  }
  validateClassInstancesMap(to_return, 'license_map')
  return to_return
}
//
// - Definition of the subscription class
//
class Subscription {
  constructor(name, purchase_range, dob_validation_method) {
    // The name of the subscription
    this.name = name
    // The range at which the number of subscriptions the same kind can be found
    this.purchase_range = purchase_range
    // The DoB validation method
    this.dob_validation_method = dob_validation_method
    // The amount entered by the operation for purchase
    this.purchased_amount = 0
    // The number of occurences that should exist based in the population
    this.occurence_count = 0
  }

  Name() { return this.name }

  // When this method run, we capture the amount of subscription of that nature
  // the operator entered.
  UpdatePurchasedSubscriptionAmount() {
    this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                         this.purchase_range.getColumn()])
  }
  PurchasedSubscriptionAmount() { return this.purchased_amount }

  // When that method run, we verify that the DoB matches the offer and if it does
  // the occurence count is incremented. In the end, this tracks how many purchased
  // subscription count (purchased_amount) should have been entered by the operator.
  IncrementAttributedSubscriptionCountIfDoB(dob) {
    if (this.ValidateDoB(dob)) {
      this.occurence_count += 1
    }
  }
  AttributedSubscriptionCount() {
    return this.occurence_count
  }

  ValidateDoB(dob) {
    return this.dob_validation_method(dob)
  }
}
//
// Categories used to from a range to establish the subscription
// ranges.
var comp_subscription_categories = [
  'U8', 'U10', 'U12+'
]

function createCompSubscriptionMap(sheet) {
  var row = 53
  var to_return = {}
  for (var rank = 1; rank <= 4; rank +=1) {
    var label = rank + comp_subscription_categories[0]
    to_return[label] = new Subscription(
      label,
      sheet.getRange(row, 5),
      (dob) => {return ageVerificationBornBetweenYearsIncluded(dob, 2016, 2017)})
    row += 1;
    label = rank + comp_subscription_categories[1]
    to_return[label] = new Subscription(
      label,
      sheet.getRange(row, 5),
      (dob) => {return ageVerificationBornBetweenYearsIncluded(dob, 2014, 2015)})
    row += 1
    label = rank + comp_subscription_categories[2]
    to_return[label] = new Subscription(
      label,
      sheet.getRange(row, 5),
      (dob) => {return ageVerificationBornBeforeYearIncluded(dob, 2013)})
    row += 1
  }
  validateClassInstancesMap(to_return, 'subscription_map')
  return to_return
}

// FIXME: This can be migrated to a validation similar to the thing we do
// for competitors.
var coord_purchased_subscriptions_non_comp = [
  [45, 5],  // First Rider. We can have more than one rider registered
  [46, 5],  // Child 1
  [47, 5],  // Child 2
  [48, 5],  // Child 3
  [49, 5]   // Child 4
];
//
// Skipass section:
//
// - Definition of the ski pass class and all possible instances.
//
class SkiPass {
  constructor(name, purchase_range, dob_validation_method, valid_dob_range_message) {
    // The name of the ski pass type
    this.name = name
    // The range at which the number of ski pass of the same kind can be found
    this.purchase_range = purchase_range
    // The DoB validation method
    this.dob_validation_method = dob_validation_method
    this.purchased_amount = 0
    this.occurence_count = 0
  }

  Name() { return this.name }

  // When this method run, we capture the amount skippasses of that type
  // the operator entered.
  UpdatePurchasedSkiPassAmount() {
    this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                         this.purchase_range.getColumn()])
  }
  PurchasedSkiPassAmount() { return this.purchased_amount }

  IncrementAttributedSkiPassCountIfDoB(dob) {
    if (this.ValidateDoB(dob)) {
      this.occurence_count += 1
    }
  }
  AttributedSkiPassCount() {
    return this.occurence_count
  }

  ValidateDoB(dob) {
    return this.dob_validation_method(dob)
  }

  ValidDoBRangeMessage() {
    return this.valid_dob_range_message
  }
}

function createSkipassMap(sheet) {
  var to_return = {
    'Collet Senior': new SkiPass(
      'Collet Senior',
      sheet.getRange(25, 5),
      (dob) => {return ageVerificationRange(dob, 70, 75)},
      "70 à 75 ans révolus"),
    'Collet Vermeil': new SkiPass(
      'Collet Vermeil',
      sheet.getRange(26, 5),
      (dob) => {return ageVerificationOlder(dob, 75)},
      '+75 ans'),
    'Collet Adulte': new SkiPass(
      'Collet Adulte',
      sheet.getRange(27, 5),
      (dob) => {return isAdult(dob) && ageVerificationYounger(dob, 70)},
      "Adulte de moins de 70 ans"),
    'Collet Étudiant': new SkiPass(
      'Collet Étudiant',
      sheet.getRange(28, 5),
      (dob) => {return ageVerificationBornBetween(dob, new Date("January 1, 1993"), new Date("December 31, 2004"))},
      '1er janvier 1993 et le 31 décembre 2004'),
    'Collet Junior': new SkiPass(
      'Collet Junior',
      sheet.getRange(29, 5),
      (dob) => {return ageVerificationBornBetween(dob, new Date("January 1, 2005"), new Date("December 31, 2012"))},
      '1er janvier 2005 et le 31 décembre 2012'),
    'Collet Enfant': new SkiPass(
      'Collet Enfant',
      sheet.getRange(30, 5),
      (dob) => {return ageVerificationBornBetween(dob, new Date("January 1, 2013"), new Date("December 31, 2017"))},
      "1er janvier 2013 et le 31 décembre 2017"),
    'Collet Bambin': new SkiPass(
      'Collet Bambin',
      sheet.getRange(31, 5),
      (dob) => {return ageVerificationBornBefore(dob, new Date("December 31, 2017"))}),

    '3D Senior': new SkiPass(
      '3D Senior',
      sheet.getRange(33, 5),
      (dob) => {return ageVerificationRange(dob, 70, 75)},
      "70 à 75 ans révolus"),
    '3D Vermeil': new SkiPass(
      '3D Vermeil',
      sheet.getRange(34, 5),
      (dob) => {return ageVerificationOlder(dob, 75)},
      '+75 ans'),
    '3D Adulte': new SkiPass(
      '3D Adulte',
      sheet.getRange(35, 5),
      (dob) => {return isAdult(dob) && ageVerificationYounger(dob, 70)},
      "Adulte de moins de 70 ans"),
    '3D Étudiant': new SkiPass(
      '3D Étudiant',
      sheet.getRange(36, 5),
      (dob) => {return ageVerificationBornBetween(dob, new Date("January 1, 1993"), new Date("December 31, 2004"))},
      '1er janvier 1993 et le 31 décembre 2004'),
    '3D Junior': new SkiPass(
      '3D Junior',
      sheet.getRange(37, 5),
      (dob) => {return ageVerificationBornBetween(dob, new Date("January 1, 2005"), new Date("December 31, 2012"))},
      '1er janvier 2005 et le 31 décembre 2012'),
    '3D Enfant': new SkiPass(
      '3D Enfant',
      sheet.getRange(38, 5),
      (dob) => {return ageVerificationBornBetween(dob, new Date("January 1, 2013"), new Date("December 31, 2017"))},
      "1er janvier 2013 et le 31 décembre 2017"),
    '3D Bambin': new SkiPass(
      '3D Bambin',
      sheet.getRange(39, 5),
      (dob) => {return ageVerificationBornBefore(dob, new Date("December 31, 2017"))}),
  }
  for (var license in to_return) {
    if (to_return[license].Name() != license) {
      displayErrorPanel('Erreur validation license_map: ' +
                        'to_return[' + license + '].Name() = ' + to_return[license].Name())
    }
  }
  validateClassInstancesMap(to_return, 'skipass_map')
  return to_return
}

// Identifier we retain as ski passes valid for competitor
var competitor_ski_passes = [
  'Collet Adulte',
  'Collet Étudiant',
  'Collet Junior',
  'Collet Enfant',
]

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
  'Senior-70-75':    [1900, 1958],
  'Senior-75+': [1900, 1958],
  'Adulte':    [1959, 2004],
  'Étudiant':  [1993, 2004],
  'Junior':    [2005, 2012],
  'Enfant':    [2013, 2017],
  'Peluche':   [2018, 2050]};
//
// - Skipass purchase indicator coordinates
//
var coord_purchased_ski_pass = {
  'Famille 4': [23, 5],
  'Famille 5': [24, 5],
  'Senior':    [25, 5],
  'Senior-70-75':    [25, 5],
  'Adulte-75+':    [26, 5],
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

function getSkiPassFamily4String() {
  return ski_pass_values[0];
}

function getSkiPassFamily5String() {
  return ski_pass_values[1];
}

function getSkiPassStudentString() {
  return ski_pass_values[4];
}

function getRiderSubscriptionCoordinates() {
  return coord_purchased_subscriptions_non_comp[0];
}

function getFirstNonRiderSubscriptionCoordinate() {
  return coord_purchased_subscriptions_non_comp[1];
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
  function setResetRebate(coord, color) {
    var row = coord[0]
    var column = coord[1]
    while(column > 0) {
      setStringAt([row, column], getStringAt([row, column]), color)
      column -= 1
    }
    SpreadsheetApp.flush() 
  }
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var pdf_number = getAndUpdateInvoiceNumber();
  // If no rebate has been entered, mask the information entirely 
  // before creating the PDF and revert that. Otherwise, don't change
  // anything.
  if (getNumberAt(coord_rebate) == 0) {
    setResetRebate(coord_rebate, "white")
  }
  var blob = createPDF(spreadsheet) 
  var pdf_filename = spreadsheet.getName() + '-' + pdf_number + '.pdf';
  var file = savePDF(blob, pdf_filename)

  // Always for the rebate area to be back in black 🤘.
  setResetRebate(coord_rebate, "black")

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
  var row = coord[0];
  var column = coord[1];
  sheet.getRange(row, column).setValue(text);
  sheet.getRange(row, column).setFontColor(color);
}

function getStringAt(coord) {
  var row = coord[0];
  var column = coord[1];
  return SpreadsheetApp.getActiveSheet().getRange(row, column).getValue().toString()
}

function getNumberAt(coord) {
  var row = coord[0];
  var column = coord[1];
  return Number(SpreadsheetApp.getActiveSheet().getRange(row, column).getValue())
}

function nthString(value) {
  var nth_strings = ["??ème", "1er", "2ème", "3ème", "4ème", "5ème"]
  if (value < 0 || value >= 5) {
    return "??ème"
  } else {
    return nth_strings[value]
  }
}

///////////////////////////////////////////////////////////////////////////////
// Date of Birth (DoB) management
///////////////////////////////////////////////////////////////////////////////

// Obtain a DoB at coords. Return an empty string if no DoB exists or if it
// does not represent a reasonable DoB
function getDoB(coords) {
  var dob = getStringAt(coords);
  if (dob != "") {
    // Verify the format
    dob = new Date(dob);
    if (dob != undefined) {
      // Verify the date is reasonable
      var yob = getDoBYear(dob);
      if (yob >= 1900 && yob <= 2050) {
        return dob;
      }
    }
  }
  return undefined;
}

// Return true if DoB happened after date
function ageVerificationBornAfter(dob, date) {
  return dob.valueOf() >= date.valueOf()
}

// Return true if DoB happened before date
function ageVerificationBornBefore(dob, date) {
  return dob.valueOf() <= date.valueOf()
}

// Return true if DoB happened between first and last date included.
function ageVerificationBornBetween(dob, first, last) {
  return (dob.valueOf() >= first.valueOf() && dob.valueOf() <= last.valueOf())
}

function ageVerificationBornBetweenYearsIncluded(dob, first, last) {
  var dob_year = getDoBYear(dob)
  return dob_year >= first && dob_year <= last
}

function ageVerificationBornBeforeYearIncluded(dob, year) {
  return getDoBYear(dob) <= year
}

// Return and age from DoB with respect to now
function ageFromDoB(dob) {
  // Calculate month difference from current date in time
  var month_diff = Date.now() - dob.getTime();
  // Convert the calculated difference in date format
  var age_dt = new Date(month_diff); 
  // Extract year from date    
  var year = age_dt.getUTCFullYear();
  // Now calculate the age of the user and return it
  return Math.abs(year - 1970);
}

// Return True if at the current time, someone with dob is strictly older than age.
function ageVerificationOlder(dob, age){
  return ageFromDoB(dob) > age
}

// Return True if at the current time, someone with dob is strictly younger than age.
function ageVerificationYounger(dob, age){
  return ageFromDoB(dob) < age
}

// Return True if at current tie, someone with dob is between age1 and age2 years old included.
function ageVerificationRange(dob, age1, age2) {
  var age = ageFromDoB(dob)
  return (age >= age1 && age <= age2);
}

// Determine whether someone is a legal adult. An adult is someone who's
// age today is 18 year old
function isAdult(dob) {
  return ageFromDoB(dob) >= 18;
}

function isMinor(dob) {
  return ! isAdult(dob)
}

// Given a list of DoBs, count the number of adults and kids in it.
function countAdultsAndKids(dobs) {
  var count_adults = 0;
  var count_children = 0;
  for (var index in dobs) {
    // Fixme: need to convert to date?
    if (isAdult(dobs[index])) {
      count_adults += 1;
    } else {
      count_children += 1;
    }
  }
  return [count_adults, count_children];
}

// Given a dd/MM/YYYY dob or a Date object, return the YoB.
function getDoBYear(dob) {
  if (typeof dob === 'string') {
    return new Date(dob).getFullYear()
  } else {
    return dob.getFullYear()
  }
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

// Normalize a name:
// - Remove leading/trailing spaces
// - Replace diacritics by their accented counterpart (for instance, é becomes e).
// - Other caracters transformed or removes
// - Optionally, the output can be upcased if required. Default is not to upcase.
function normalizeName(str, to_upper_case=false) {
  to_return = str.trim().normalize("NFD").replace(/\p{Diacritic}/gu, "").
      replace(/\s/g, "-").  // No spaces in the middle
      replace(/\d+/g, "").  // No numbers
      replace(/\//g, "-").  // / into -
      replace(/\./g, "-").  // . into -
      replace(/_/g, "-").   // _ into -
      replace(/:/g, "-").   // : into -
      replace(/-+/g, "-");  // Many - into a single one
  if (to_upper_case) {
    to_return = to_return.toUpperCase()
  }
  return to_return
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

function isLicenseDefined(license) {
  // Fixme: extra paranoid: check it's a key of the license hash.
  return license != '' && license != getNoLicenseString()
}

function isLicenseNotDefined(license) {
  return !isLicenseDefined(license)
}

function isLicenseNonComp(license) {
  return (license == getNonCompJuniorLicenseString() ||
          license == getNonCompFamilyLicenseString())
}

function isLicenseComp(license) {
  return (license == getCompAdultLicenseString() ||
          license == getCompJuniorLicenseString())
}

function isLicenseCompAdult(license) {
  return license == getCompAdultLicenseString()
}

function isLicenseCompJunior(license) {
  return license == getCompJuniorLicenseString()
}

function isExecLicense(license) {
  return license == getExecutiveLicenseString()
}

function isLicenseFamily(license) {
  return license == getNonCompFamilyLicenseString()
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

// Verify that family members have a first name, 
// last name, a DoB and a sex assigned to them.
// Return an error and also a list of collected DoBs.
function validateFamilyMembers() {
  function returnError(m) {
    return [m, []]
  }

  dobs = [];
  for (var index in coords_identity_rows) {
    var first_name = getStringAt([coords_identity_rows[index], coord_first_name_column]);
    var last_name = getStringAt([coords_identity_rows[index], coord_last_name_column]);
    // Normalize the first name, normalize/upcase the last name
    first_name = normalizeName(first_name, false)
    last_name = normalizeName(last_name, true)

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
    // Write the first and last name back as it's been normalized
    setStringAt([coords_identity_rows[index], coord_first_name_column], first_name, "black");
    setStringAt([coords_identity_rows[index], coord_last_name_column], last_name, "black");

    var license = getLicenseAt([coords_identity_rows[index], coord_license_column]);
    // We need a DoB but only if a license has been requested. If a DoB is
    // defined, we save it.
    var dob = getDoB([coords_identity_rows[index], coord_dob_column]);
    if (dob == undefined) {
      if (isLicenseDefined(license)) {
        return returnError(
          "Pas de date de naissance fournie pour " +
          first_name + " " + last_name +
          " ou date de naissance mal formatée (JJ/MM/AAAA)\n" +
          " ou année de naissance fantaisiste.");
      }
    } else {
      dobs.push(dob);
    }
    // We need a level but only if a non competitor license has been requested
    // We exclude adults because a non competitor license can be a family license.
    var level = getStringAt([coords_identity_rows[index], coord_level_column]);
    if (level == '' && isMinor(dob) && isLicenseNonComp(license)) {
      return returnError(
          "Pas de niveau fourni pour " + first_name + " " + last_name         
      )
    }
    // We need a sex
    var sex = getStringAt([coords_identity_rows[index], coord_sex_column]);
    if (sex != "Fille" && sex != "Garçon") {
      return returnError("Pas de sexe défini pour " +
                         first_name + " " + last_name);
    }
  }
  return ["", dobs];
}

// Loosely cross checks attributed licenses with delivered subscriptions...
function validateLicenseSubscription(attributed_licenses) {

  function errorMessageBadSubscriptionValue(index) {
    return ("La valeur du champ 'Adhésion / Stage / Transport - enfant " +
            index + "' ne peut prendre que la valeur 0 ou 1.")
  }

  // Validation of the selection of the subscriptions:
  //   - A rider counts as a first subscription. The value can only be 1.
  //     This is a limitation but that's OK.
  //   - We can't have a rider and a first subscription.
  //   - Past a rider/first subscription, the value can for a subscription
  //     can only be 1 or 0 and when it reaches 0, it can't be 1 again.

  // Rider cell value validation
  var rider = getNumberAt(getRiderSubscriptionCoordinates())
  var first_subscription = getNumberAt(getFirstNonRiderSubscriptionCoordinate())
  if (rider < 0) {
    return ("La valeur du champ 'Adhésion Rider / Stage / Sortie hors station / Transport'" +
            " ne peut prendre la valeur " + rider);
  }
  // First subscription cell validation
  if (first_subscription != 0 && first_subscription != 1) {
    return errorMessageBadSubscriptionValue(1)
  }

  // Can't have a rider and a first cell
  if (rider >= 1 && first_subscription == 1) {
    return ("L'adhésion rider compte comme une première Adhésion / Stage / Transport")
  }

  var total_non_comp = 0;
  var reached_zero = false;
  if (rider >= 1 || first_subscription == 1) {
    total_non_comp = rider + first_subscription;
  }
  if (rider == 0 && first_subscription == 0) {
    reached_zero = true;
  }

  // Validate the subscriptions past the rider/first subscription.
  for (var index in coord_purchased_subscriptions_non_comp) {
    // We verified the rider and the first subscription, we are only verifying
    // the subscriptions past those - hence the test index > 1 (rider is index 0,
    // first subscription is index 1)
    if (index > 1) {
      value = getNumberAt(coord_purchased_subscriptions_non_comp[index]);
      if (value != 1 && value != 0) {
        return errorMessageBadSubscriptionValue(index)
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
      total_non_comp += value;
    }
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

  if (family4_quantity == 1 && family5_quantity == 1) {
    return ("Deux forfaits annuels Famille 4 personnes et 5 personnes sélectionnés.")
  }
  
  var error = ""
  // If we validated a family pass, return now. The rest of the validation is
  // going to be too complicated.
  if (family4_quantity != 0 || family5_quantity != 0) {
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

function validateSkiPassComp() {
  var ski_passes_map = createSkipassMap(SpreadsheetApp.getActiveSheet())
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getLicenseAt([row, coord_license_column]);
    if (! isLicenseComp(selected_license)) {
      continue
    }
    var dob = getDoB([row, coord_dob_column])
    // Increment the ski pass count that validates for a DoB. This will tell us
    // how many ski passes suitable for competitor we can expect to be
    // purchased.
    for (var skipass in ski_passes_map) {
      ski_passes_map[skipass].IncrementAttributedSkiPassCountIfDoB(dob)
    }
  }

  // Collect the amount of skipasses that have be declared for purchase.
  for (var skipass in ski_passes_map) {
    ski_passes_map[skipass].UpdatePurchasedSkiPassAmount()
  }

  // Go over all the skipass that apply to a competitor and perform the verification:
  // If the occurence count of a skipass that can apply to a competitor
  // is above than the purchased amount, we have an error (some competitor are not
  // buying a ski pass). If it's below it means that the amount of purchased ski passes
  // includes competitors and non competitors which is fine.
  for (var index in competitor_ski_passes) {
    var skipass_name = competitor_ski_passes[index]
    var skipass = ski_passes_map[skipass_name]
    if (skipass.AttributedSkiPassCount() > skipass.PurchasedSkiPassAmount()) {
      return (skipass.PurchasedSkiPassAmount() + ' forfait(s) ' + skipass_name +
              ' acheté(s) pour ' + skipass.AttributedSkiPassCount() +
              ' license(s) compétiteur dans cette tranche d\'âge')
    }
  }
  return ''
}

function validateCompetitionSubscriptions() {
  var comp_subscription_map = createCompSubscriptionMap(SpreadsheetApp.getActiveSheet())
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getLicenseAt([row, coord_license_column]);
    if (! isLicenseComp(selected_license)) {
      continue
    }
    var dob = getDoB([row, coord_dob_column])
    // Increment the subscription count that validates for a DoB. This will tell us
    // how many subscription suitable for competitor we can expect to be
    // purchased.
    for (var subscription in comp_subscription_map) {
      comp_subscription_map[subscription].IncrementAttributedSubscriptionCountIfDoB(dob)
    }
  }

  // Collect the amount of subscriptions that have be declared for purchase.
  for (var subscription in comp_subscription_map) {
    comp_subscription_map[subscription].UpdatePurchasedSubscriptionAmount()
  }

  // Verification:
  // 1- For each category, the number of licenses purchased matches the number of
  //    licensees.
  // 2- The number of licensee in one category can only be either 0 or 1
  // 3- A category rank (1st, 2nd, etc...) can only be filled if all the previous
  //    ranks have been filled.
  for (index in comp_subscription_categories) {
    var category = comp_subscription_categories[index]
    var total_existing = 0
    var total_purchased = 0
    for (var rank = 1; rank <= 4; rank += 1) {
      var indexed_category = rank + category
      // The number of existing competitor in an age range has already been accumulated.
      // just capture it. The number of purchased competitor needs to be accumulated as it's
      // spread over several cells
      total_existing = comp_subscription_map[indexed_category].AttributedSubscriptionCount()
      var current_purchased = comp_subscription_map[indexed_category].PurchasedSubscriptionAmount()
      if (current_purchased > 1 || current_purchased < 0 || ~~current_purchased != current_purchased) {
        return ('Le nombre d\'adhésion(s) ' + category + ' achetée(s) (' + 
                current_purchased + ') n\'est pas valide.')
      }
      // If we have a current purchased subscription past rank 1, we should have a purchased
      // subscription in the previous ranks for all categories
      if (rank > 1 && current_purchased == 1) {
        var subscription_other_category = 0
        for (var reversed_rank = rank - 1; reversed_rank >= 1; reversed_rank -= 1) {
          for (var reversed_category_index in comp_subscription_categories) {
            var reversed_indexed_category = (reversed_rank +
                                             comp_subscription_categories[reversed_category_index])
            subscription_other_category += 
              comp_subscription_map[reversed_indexed_category].PurchasedSubscriptionAmount()
          }
          if (subscription_other_category == 0) {
            return ('Une adhésion compétition existe pour un(e) ' + nthString(rank) + ' ' + category +
                    ' sans adhésion déclarée pour un ' + nthString(rank-1) +
                    ' enfant')
          }
        }
      }
      total_purchased += current_purchased
    }
    // For that category, the total number of purchased licenses must match
    // the number of accumulated purchases accross all ranks.
    if (total_existing != total_purchased) {
      return (total_purchased + ' adhésion(s) compétition ' + category +
              ' achetée(s) pour ' + total_existing +
              ' license(s) compétiteur dans cette tranche d\'âge')
    }
  }
  // 4- Only one category can be filled per rank. That loops needs
  //    to start for each rank so the loop above can not be used.
  for (var rank = 1; rank <= 4; rank += 1) {
    var total_purchased_for_rank = 0
    for (index in comp_subscription_categories) {
      var category = comp_subscription_categories[index] 
      var indexed_category = rank + category
      total_purchased_for_rank += comp_subscription_map[indexed_category].PurchasedSubscriptionAmount()
      if (total_purchased_for_rank > 1) {
        return (total_purchased_for_rank + ' adhésions compétition achetées pour un ' + 
                nthString(rank) + ' enfant. Ce nombre ne peut dépasser 1')
      }
    }
  }
  return ''
}

// Cross check the attributed licenses with the ones selected for payment
function validateLicenseCrossCheck(license_map, dobs) {
  function returnError(v) {
    return [v, {}];
  }
  
  // Collect the attributed licenses
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getLicenseAt([row, coord_license_column]);
    // You can't have no first/last name and an assigned license
    var first_name = getStringAt([row, coord_first_name_column]);
    var last_name = getStringAt([row, coord_last_name_column]);
    // If there's no name on that row, the only possible value is None
    if (first_name === '' && last_name === '') {
      if (isLicenseDefined(selected_license)) {
        return returnError("'" + selected_license +
                           "' attribuée à un membre de famile inexistant!");
      }
      continue;
    }

    // Sanity check: the selected license must exist and increment
    // the count number for the license at hand.
    if (!license_map.hasOwnProperty(selected_license)) {
      return returnError("'" + selected_license +
                         "' n'est pas une license attribuée possible!");
    }
    license_map[selected_license].IncrementAttributedLicenseCount()

    // Executive license requires a city of birth
    if (isExecLicense(selected_license)) {
      var city = getStringAt([row, coord_cob_column]);
      if (city == '') {
        return returnError(first_name + " " + last_name + ": une license " +
                           selected_license +
                           " requiert de renseigner une ville et un pays de " +
                           "naissance");
      }
    }
    
    // If we don't have a license, we can stop now. No need to
    // validate the DoB
    if (isLicenseNotDefined(selected_license)) {
      continue;
    }
    
    // Validate DoB against the type of license
    var dob = getDoB([row, coord_dob_column])
    var yob = getDoBYear(dob)
    if (! license_map[selected_license].ValidateDoB(dob)) {
      return returnError(first_name + " " + last_name +
                         ": l'année de naissance " + yob +
                         " ne correspond aux années de validité de la " +
                         "license choisie: '" + selected_license + " " +
                         license_map[selected_license].ValidDoBRangeMessage())
    }
  }

  // Collect the amount of purchased licenses for all licenses
  for (var index in license_map) {
    license_map[index].UpdatePurchasedLicenseAmount()
  }
  
  // Perform the verification
  
  // First verify the family non comp license because it's special.
  // 1- Family license requires at last four declared family members
  //    selecting it
  // 2- One family license should have been purchased
  // 3- Other checks are carried out still as they can be determined
  //    valid or not
  //    even with a family license purchased.
  var family_license_attributed_count = license_map[getNonCompFamilyLicenseString()].AttributedLicenseCount()
  if (family_license_attributed_count != 0) {
    // A least four declared participants
    if (family_license_attributed_count < 4) {
      return returnError(
        "Il faut attribuer une licence famille à au moins 4 membres " +
        "d'une même famille. Seulement " + 
        family_license_attributed_count +
        " ont été attribuées.");
    }
    // Check that one family license was purchased.
    var family_license_purchased = license_map[getNonCompFamilyLicenseString()].PurchasedLicenseAmount()
    if (family_license_purchased != 1) {
      return returnError(
        "Vous devez acheter une licence loisir famille, vous en avez " +
        "acheté pour l'instant " + family_license_purchased);
    }
    // Last verification: two adults and two+ kids
    var res = countAdultsAndKids(dobs);
    var count_adults = res[0];
    var count_children = res[1];
    if (count_adults < 2 || count_adults + count_children < 4) {
      return returnError(
        "Seulement " + count_adults + " adulte(s) déclaré(s) et " +
        count_children + " enfants(s) de moins de 17 ans déclaré(s) " +
        "pour la license famille.");
    }
  }
  
var attributed_licenses = {}
for (var index in license_map) {
    attributed_licenses[index] = license_map[index].AttributedLicenseCount()
    // Entry indicating no license is skipped because it can't
    // be collected. Entry indicating a family license skipped because
    // it's been already verified
    if (isLicenseNotDefined(index) || isLicenseFamily(index)) {
      continue
    }
    // What you attributed must match what you're purchasing...
    if (license_map[index].PurchasedLicenseAmount() != license_map[index].AttributedLicenseCount()) {
      return returnError(
        "Le nombre de licence(s) '" + index + "' attribuée(s) (au nombre de " +
        license_map[index].AttributedLicenseCount() + ")\n" +
        "ne correspond pas au nombre de licence(s) achetée(s) (au nombre de " +
        license_map[index].PurchasedLicenseAmount() + ")")
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

class FamilyMember {
  constructor(first_name, last_name, dob, sex, city, license_type, license_number, level) {
    this.first_name = first_name
    this.last_name = last_name
    this.dob = dob
    this.sex = sex
    this.city = city
    this.license_type = license_type
    this.license_number = license_number
    this.level = level
  }
}

// Produce a list of family members that has purchased a license.
// This assumes that some verification of first/last name, DoB, sex and level
// have already been performed. Each family member is an instance of the
// FamilyMember class.
function getListOfFamilyPurchasingALicense() {
  var family = []
  for (var index in coords_identity_rows) {    
    var first_name = getStringAt([coords_identity_rows[index], coord_first_name_column]);
    var last_name = getStringAt([coords_identity_rows[index], coord_last_name_column]);    
    // After validation of the family entries, having a first name
    // guarantees a last name. Just check the existence of a
    // first name in order to skip that entry
    if (first_name == "") {
      continue;
    }
    // We can skip that entry if no license is required. That familly
    // member doesn't need to be reported in this dictionary.
    var license = getLicenseAt([coords_identity_rows[index], coord_license_column]);
    if (isLicenseNotDefined(license)) {
      continue
    }
    var license_number = getStringAt([coords_identity_rows[index], coord_license_number_column])

    // DoB is guaranteed to be there if a license was requested
    var birth = getDoB([coords_identity_rows[index], coord_dob_column]);
    var city = getStringAt([coords_identity_rows[index], coord_cob_column])
    if (city == "") {
      city = "\\";
    }
    // Sex is guaranteed to be there
    var sex = getStringAt([coords_identity_rows[index], coord_sex_column])

    // Level is guaranteed to be there, it should have been verified
    // in validateAndReturnDropDownValue()
    var level = getStringAt([coords_identity_rows[index], coord_level_column])

    family.push(new FamilyMember(first_name, last_name, birth,
                                 sex, city, license, license_number, level))
  }
  return family
}

///////////////////////////////////////////////////////////////////////////////
// Updating the FFS tab of the aggregation trix with entered data (only once)
///////////////////////////////////////////////////////////////////////////////

// Find the first empty slot in range and return the range it corresponds to.
// Note that parameter range must be a column
function findFirstEmptySlot(sheet, range) {
    // Range must be a column
    var values = range.getValues();
    var ct = 0;
    while ( values[ct] && values[ct][0] != "" ) {
      ct++;
    }
    return sheet.getRange(range.getRow() + ct,range.getColumn())
}

function doUpdateAggregationTrix(data, allow_overwrite) {
  
  // Search for elements in data in sheet over range. If we can't find data,
  // return a range on the first empty slot we find. If we can find data
  // return its range unless allow_overwrite is false, in which case we
  // return null: this will be used to avoid overwriting existing data.
  function SearchEntry(sheet, range, data, allow_overwrite) {
    var finder = range.createTextFinder(data.last_name)
    while (true) {
      var current_range = finder.findNext()
      if (current_range == null) {
        return findFirstEmptySlot(sheet, range)
      }
      var row = current_range.getRow()
      var col = current_range.getColumn()
      // This assumes that first_name will be found at col+1 relative to last_name.
      if (sheet.getRange(row, col+1).getValue().toString() == data.first_name) {
        if (allow_overwrite) {
          return sheet.getRange(row, col)
        }
        return null
      }
    }
  }

  // Update the row at range in sheet with data
  function UpdateRow(sheet, range, data) {
    var row = range.getRow()
    var column = range.getColumn()
    sheet.getRange(row,column).setValue(data.last_name)
    sheet.getRange(row,column+1).setValue(data.first_name)
    sheet.getRange(row,column+2).setValue(data.license_number)
    sheet.getRange(row,column+3).setValue(data.sex)
    sheet.getRange(row,column+4).setValue(data.dob)
    sheet.getRange(row,column+5).setValue(getDoBYear(data.dob))
    sheet.getRange(row,column+6).setValue(data.level)
  }

  // Open the spread sheet, insert the name if the operation is possible. Sync
  // the spreadsheet.
  var sheet = SpreadsheetApp.openById(license_trix).getSheetByName('FFS');
  var last_name_range = sheet.getRange('B7:B')
  var entire_range = sheet.getRange('B7:N')

  for (var index in data) {
    var res = SearchEntry(sheet, last_name_range, data[index], allow_overwrite)
    // res can be null if allow_overwrite is false and the entry was found
    if (res == null) {
      continue
    }
    UpdateRow(sheet, res, data[index])
  }
  // Sort the spread sheet by last name and then first name and sync the spreadsheet.
  entire_range.sort([{column: entire_range.getColumn()}, {column: entire_range.getColumn()+5}])
  SpreadsheetApp.flush()
}

function updateAggregationTrix() {
  // We computed this already possible in the caller, maybe pass this as an argument?
  var family_dict = getListOfFamilyPurchasingALicense()
  var family = []
  // Skip over an empty name
  for (var index in family_dict) {
    var family_member = family_dict[index]
    if (family_dict[index].last_name == "") {
      continue
    }

    // Retain kids with a non comp license (can be a junior license or a family license)
    if (isMinor(family_member.dob) && isLicenseNonComp(family_member.license_type)) {
      family.push(family_member)
    }
  }
  doUpdateAggregationTrix(family, false)
}

///////////////////////////////////////////////////////////////////////////////
// Updating the tab tracking registration with issues in the aggregation trix
///////////////////////////////////////////////////////////////////////////////

function updateProblematicRegistration(link, context) {
  updateStatusBar('⏳ Enregistrement de la notification de problème...', 'orange')
  var sheet = SpreadsheetApp.openById(license_trix).getSheetByName('Dossiers problématiques')
  var insertion_range = findFirstEmptySlot(sheet, sheet.getRange('A2:A'))
  var entire_range = sheet.getRange('A2:D')
  var row = insertion_range.getRow()
  var column = insertion_range.getColumn()
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-YYYY, HH:mm')
  sheet.getRange(row, column).setValue(date)
  sheet.getRange(row, column+1).setValue(link)
  sheet.getRange(row, column+2).setValue(context)
  entire_range.sort([{column: entire_range.getColumn(), ascending: false}])
  SpreadsheetApp.flush()
}

///////////////////////////////////////////////////////////////////////////////
// Top level invoice validation
///////////////////////////////////////////////////////////////////////////////

// Validate the invoice and return a dictionary of values
// to use during invoice generation.
function validateInvoice() {
  function augmentEscapeHatch(source) {
    return source + ("\n\nChoisissez 'OK' pour continuer à générer la facture.\n" +
                    "Choisissez 'Annuler' pour ne pas générer la facture et " +
                    "vérifier les valeurs saisies...");
  }
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

  // Validation: first phone number
  var phone_number = getStringAt(coord_family_phone1)
  if (phone_number == '') {
    displayErrorPanel(
      "Vous n'avez pas renseigné de nom de numéro de telephone ou " +
      "vous avez oublié \n" +
      "de valider le numéro de téléphone par [return] ou [enter]...")
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
  if (family_validation_error) {
    displayErrorPanel(family_validation_error);
    return {};
  }
  var dobs = ret[1];

  // Validate de competitor ski passes
  var validate_ski_pass_comp_error = validateSkiPassComp()
  if (validate_ski_pass_comp_error) {
    if (! displayYesNoPanel(augmentEscapeHatch(validate_ski_pass_comp_error))) {
      return {};
    }      
  }

  var validate_subscription_comp_error = validateCompetitionSubscriptions()
  if (validate_subscription_comp_error) {
    if (! displayYesNoPanel(augmentEscapeHatch(validate_subscription_comp_error))) {
      return {};
    }      
  }  
  // Now performing the optional/advanced validations... 
  //
  // 1- Validate the licenses requested by this family
  if (advanced_verification_family_licenses) {
    var license_map = createLicensesMap(SpreadsheetApp.getActiveSheet())
    var ret = validateLicenseCrossCheck(license_map, dobs);
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
    var collected_attributed_licenses_values = ret[1];
    // Validate requested licenses and subscriptions
    var subscription_validation_error = 
        validateLicenseSubscription(collected_attributed_licenses_values);
    if (subscription_validation_error) {
      if (! displayYesNoPanel(augmentEscapeHatch(subscription_validation_error))) {
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
      if (! displayYesNoPanel(augmentEscapeHatch(skipass_validation_error))) {
        return {};
      }    
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
    displayErrorPanel(
      "L'autorisation parentale doit être signée aujourd'hui pour " +
      "valider le dossier et terminer l'inscription");
    return {};
  }

  // Validate the medical form
  var medical_form = validateAndReturnDropDownValue(
    coord_medical_form,
    "Vous n'avez pas renseigné votre réponse (OUI/NON) au questionaire médicale.")
  if (consent == '') {
    return {}
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
          'consent': consent,
          'medical_form': medical_form};
}

///////////////////////////////////////////////////////////////////////////////
// Invoice emailing
///////////////////////////////////////////////////////////////////////////////

function maybeEmailLicenseSCA(invoice) {
  var operator = getOperator()
  var family_name = getFamilyName()
  var family_dict = getListOfFamilyPurchasingALicense() 
  var string_family_members = "";
  for (var index in family_dict) {
    if (family_dict[index].last_name == "") {
      continue
    }
    string_family_members += (
      "<tt>" +
      "Nom: <b>" + family_dict[index].last_name.toUpperCase() + "</b><br>" +
      "Prénom: " + family_dict[index].first_name + "<br>" +
      "Naissance: " + Utilities.formatDate(family_dict[index].dob, Session.getScriptTimeZone(), 'dd-MM-YYYY') + "<br>" +
      "Fille/Garçon: " + family_dict[index].sex + "<br>" +
      "Ville de Naissance: " + family_dict[index].city + "<br>" +
      "Licence: " + family_dict[index].license_type + "<br>" +
      "Numéro License: " + family_dict[index].license_number + "<br>" +
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

function updateStatusBar(message, color) {
  setStringAt(coord_status, message, color)
  SpreadsheetApp.flush()
}

function generatePDFAndMaybeSendEmail(send_email, just_the_invoice) {
  updateStatusBar("⏳ Validation de la facture...", "orange")      
  var validation = validateInvoice();
  if (isEmpty(validation)) {
    updateStatusBar("❌ La validation de la facture a échouée", "red")      
    return;
  }
  updateStatusBar("⏳ Préparation de la facture...", "orange")
  
  // Generate and prepare attaching the PDF to the email
  var pdf_file = generatePDF();
  var pdf = DriveApp.getFileById(pdf_file.getId());
  var attachments = [pdf.getAs(MimeType.PDF)]

  if (just_the_invoice) {
    updateStatusBar("⏳ Génération " +
                    (send_email? "et envoit " : " ") + "de la facture...", "orange")
  } else {
    updateStatusBar("⏳ Génération " +
                    (send_email? "et envoit " : " ") + "du dossier...", "orange")
  }
  
  var civility = validation['civility'];
  var family_name = validation['family_name'];
  var mail_to = validation['mail_to'];
  var consent = validation['consent'];
  var medical_form = validation['medical_form']
  
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
  
  // Take a look at the medical form answer:
  // 1- Yes: we need to tell that a medical certificate needs to be provided
  // 2- No: a new attachment need to be added.
  var medical_form_text = ''
  if (medical_form == 'Une réponse OUI') {
    medical_form_text = ('<p><b><font color="red">' +
                         'Les réponses que vous avez porté au questionaire médicale vous ' +
                         'obligent à transmettre dans au SCA (inscriptions.sca@gmail.com) les plus ' +
                         'brefs délais un certificat médical en cours de validité' +
                         '</font></b>')
  } else if (medical_form == 'Toutes réponses NON') {
    medical_form_text = ('<p><b><font color="red">' +
                         'Les réponses que vous avez porté au questionaire médical vous ' +
                         'obligent à signer la page ' + information_leaflet_page +
                         ' de la notice d\'informations ' + season + ' fournie en attachement' +
                         '</font></b>')
    attachments.push(DriveApp.getFileById(information_leaflet_pdf).getAs(MimeType.PDF))
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
      medical_form_text +

      "<p>Des questions concernant cette facture? Contacter Marlène: " +
      "marlene.czajka@gmail.com (06-60-69-75-39) / Aurélien: " +
      "armand.aurelien@gmail.com (07-69-62-84-29) pour le ski loisir ou " +
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
  // The final status to display is captured in this variable and
  // the status bar is updated after the aggregation trix has been
  // updated.
  var final_status = ['', 'green']
  if (send_email) {
    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    if (emailQuotaRemaining < quota_threshold) {
      displayWarningPanel("Quota email insuffisant. Envoi " +
                          (just_the_invoice ? "de la facture" : "du dossier") +
                          " retardé. Quota restant: " + emailQuotaRemaining +
                          ". Nécessaire: " + quota_threshold);
      final_status[0] = "⚠️ Quota email insuffisant. Envoi retardé"
      final_status[1] = "orange";
      // Insert a reference to the file in a trix
      var link = '=HYPERLINK("' + SpreadsheetApp.getActive().getUrl() +
                             '"; "' + SpreadsheetApp.getActive().getName() + '")'
      var context = (just_the_invoice ? 'Facture seule à renvoyer' : 'Dossier complet à renvoyer')
      updateProblematicRegistration(link, context)
    } else {
      // Send the email  
      MailApp.sendEmail(email_options)
      maybeEmailLicenseSCA([attachments[0]]);
      final_status[0] = "✅ " + (just_the_invoice ? 
                                "Facture envoyée" : "dossier envoyé")
    }
  } else {
      final_status[0] = "✅ " + (just_the_invoice ?
                        "Facture générée" : "dossier généré")
  }

  // Now we can update the level aggregation trix with all the folks that
  // where declared as not competitors
  updateStatusBar("⏳ Enregistrement des niveaux...", "orange")
  updateAggregationTrix()
  // And deliver the final status.
  updateStatusBar(final_status[0], final_status[1])

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

// This is what the [signal problem] button runs
function SignalProblem() {
  var link = '=HYPERLINK("' + SpreadsheetApp.getActive().getUrl() +
                         '"; "' + SpreadsheetApp.getActive().getName() + '")'
  updateProblematicRegistration(link, getStringAt(coord_status))
  updateStatusBar('⚠️ Facture signalée commme problématique', 'red')
}
