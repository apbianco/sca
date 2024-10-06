// Version: 2024-10-24 BAS
// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.
//
// Important note on development and tests:
// - Development: use Firefox only while inscriptions.sca@gmail.com is
//   the only account that is connected
// - Test: use Safari only while inscriptions.sca@gmail.com is
//   the only account that is connected. Once an invoice has been generated,
//   it's possible to debug its code in Safari.

// Dev or prod? "dev" sends all email to email_dev. Prod is the
// real thing: family will receive invoices, and so will email_license,
// unless the trigger in use is the TEST trigger.
dev_or_prod = "dev"

// Enable/disable new features - first entry set to false
// requires all following entries set to false. Note that
// a validation of that constraint is carried out immediately
// after these definitions.
//
// The coresponding truth table is:
//
// a     b     c     result
// ------------------------
// 0     0     0     1
// 1     0     0     1
// 1     1     0     1
// 1     1     1     1
// *     *     *     0
//
// result = !b!c + ab
class AdvancedFeatures {
  constructor() {
    // a
    this.verification_family_licenses = false
    // b
    this.verification_subscriptions = false
    // c
    this.verification_skipass = false
  }

  Validate() {
    // !b!c
    var not_b_not_c = ! this.verification_subscriptions && ! this.verification_skipass
    // ab
    var ab = this.verification_family_licenses && this.verification_subscriptions
    // !b!c + ab
    var result = not_b_not_c || ab
    if (!result) {
      displayErrorPanel(
        "verification_family_licenses: " + this.verification_family_licenses +
        "\nverification_subscriptions: " + this.verification_subscriptions +
        "\nverification_skipass: " + this.verification_skipass)
    }
  }

  SetAdvancedVerificationFamilyLicenses() {
    this.verification_family_licenses = true
    this.Validate()
  }
  AdvancedVerificationFamilyLicenses() {
    return this.verification_family_licenses
  }

  SetAdvancedVerificationSubscriptions() {
    this.verification_subscriptions = true
    this.Validate()
  }
  AdvancedVerificationSubscription() {
    return this.verification_subscriptions
  }

  SetAdvancedVerificationSkipass() {
    this.verification_skipass = true
    this.Validate()
  }
  AdvancedVerificationSkipass() {
    return this.verification_skipass
  }
}

// Create the advanced validation global instance and turn on the advanced
// validation features. The verification happens as soon as the feature is set.
var advanced_validation = new AdvancedFeatures()
advanced_validation.SetAdvancedVerificationFamilyLicenses()
advanced_validation.SetAdvancedVerificationSubscriptions()
advanced_validation.SetAdvancedVerificationSkipass()

///////////////////////////////////////////////////////////////////////////////
// Seasonal parameters - change for each season
///////////////////////////////////////////////////////////////////////////////
//
// - Name of the season
var season = "2024/2025"
//
// - Storage for the current season's database.
//
var db_folder = '1GmOdaWlEwH1V9xGx3pTp1L3Z4zXQdjjn'
//
//
// Level aggregation trix to update when a new entry is added
//
var license_trix = '1tR3HvdpXWwjziRziBkeVxr4CIp10rHWyfTosv20dG_I'
//
// - ID of attachements to be sent with the invoice - some may change
//   from one season to an other when they are refreshed.
//
// PDF content to insert in a registration bundle.
//
var legal_disclaimer_pdf = '18jFQWTmLnmBa9HGmPkFS58xr0GjNqERu'
var rules_pdf = '1U-eeiEFelWN4aHMwjHJ9IQRH3h2mZJoW'
var parents_note_pdf = '1fVb5J3o8YikPcn4DDAplt9X-XtP9QdYS'
var ffs_information_leaflet_pdf = '1zxP502NjvVo0AKFt_6FCxs1lQeJnNxmV'
// The page in ffs_information_leaflet_pdf parents should sign
var ffs_information_leaflet_page = 16

// Spreadsheet parameters (row, columns, etc...). Adjust as necessary
// when the master invoice is modified.
// 
// - Locations of family details:
//
var coord_family_civility = [6,  3]
var coord_family_name =     [6,  4]
var coord_family_street =   [8,  3]
var coord_family_zip =      [8,  4]
var coord_family_city =     [8,  5]
var coord_family_email =    [9,  3]
var coord_cc =              [9,  5]
var coord_family_phone1 =   [10, 3]
var coord_family_phone2 =   [10, 5]
//
// - Locations of various status line and collected input, located
//   a the bottom of the invoice.
// 
var coord_rebate =           [76, 4]
var coord_charge =           [77, 4]
var coord_personal_message = [85, 3]
var coord_timestamp =        [86, 2]
var coord_version =          [86, 3]
var coord_legal_disclaimer = [86, 5]
var coord_medical_form =     [86, 7]
var coord_callme_phone =     [86, 9]
var coord_yolo =             [87, 3]
var coord_status =           [88, 4]
var coord_generated_pdf =    [88, 6]
//
// - Rows where the family names are entered
// 
var coords_identity_rows = [14, 15, 16, 17, 18, 19];
//
// - Columns where information about family members can be found
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
      getNoLicenseString(),
      null,
      (dob) => {return true}),
    'CN Jeune (Loisir)': new License(
      getNonCompJuniorLicenseString(),
      sheet.getRange(41, 5),
      // Born after 2010 is born on 1/1/2010 or later.
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, new Date("January 1, 2010"))},
      "être né en 2010 et après"),
    'CN Adulte (Loisir)': new License(
      getNonCompAdultLicenseString(),
      sheet.getRange(42, 5),
      // Born on 2009 and before means 12/1/2009 or before.
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, new Date("December 31, 2009"))},
      "être né en 2009 et avant"),
    'CN Famille (Loisir)': new License(
      getNonCompFamilyLicenseString(),
      sheet.getRange(43, 5),
      (dob) => {return true},
      ""),
    'CN Dirigeant': new License(
      getExecutiveLicenseString(),
      sheet.getRange(44, 5),
      (dob) => {return isAdult(dob)},
      "être adulte (18 ans ou plus)"),
    'CN Jeune (Compétition)': new License(
      getCompJuniorLicenseString(),
      sheet.getRange(51, 5),
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, new Date("January 1, 2010"))},
      "être né en 2010 et après"),
    'CN Adulte (Compétition)': new License(
      getCompAdultLicenseString(),
      sheet.getRange(52, 5),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, new Date("December 31, 2009"))},
      "être né 2009 et avant"),
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
  validateClassInstancesMap(to_return, 'createCompSubscriptionMap')
  return to_return
}

var noncomp_subscription_categories = ['Rider', '1er enfant', '2ème enfant', '3ème enfant', '4ème enfant']

function createNonCompSubscriptionMap(sheet) {
  var to_return = {}
  var row = 45
  for (index in noncomp_subscription_categories) {
    var label = noncomp_subscription_categories[index]
    to_return[label] = new Subscription(
      label,
      sheet.getRange(row, 5),
      (dob) => {return true}),
    row += 1
  }
  validateClassInstancesMap(to_return, 'createNonCompSubscriptionMap')
  return to_return
}

// This defines a level: the level is undetermined
function getNoLevelString() {
  return 'Non déterminé'
}

// This DOES NOT define a level. It's equivalent to not
// setting anything.
function getNALevelString() {
  return 'Pas Concerné'
}

// This is almost useless - a syntatic sugar
function getLevelAt(coord) {
  return getStringAt(coord)
}

// A level is not adjusted when it starts with "⚠️ "
function isLevelNotAdjusted(level) {
  return level.substring(0, 3) == "⚠️ ";
}

// NOTE: A level is not defined when it has not been entered or
// when it has been set to getNALevelString()
// NOTABLY:
//  - A level of getNoLevelString() value *DEFINES* a level.
//  - A level not yet adjusted *DEFINES* a level.
function isLevelNotDefined(level) {
  return level == '' || level == getNALevelString()
}

function getRiderLevelString() {
  // FIXME: Fragile
  return noncomp_subscription_categories[0]
}

function getFirstKid() {
  // FIXME: fragile?
  return noncomp_subscription_categories[1]
}

function isFirstKid(subscription) {
  return isLevelDefined(subscription) && subscription == getFirstKid()
}

function isLevelDefined(level) {
  return ! isLevelNotDefined(level)
}

function isLevelRider(level) {
  return level == getRiderLevelString()
}

function isLevelNotRider(level) {
  return isLevelDefined(level) && ! isLevelRider(level)
}

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
    this.valid_dob_range_message = valid_dob_range_message
    this.purchased_amount = 0
    this.occurence_count = 0
    this.student = isSkipassStudent(name)
    this.adult = isSkipassAdult(name)
  }

  Name() { return this.name }

  // When this method run, we capture the amount skippasses of that type
  // the operator entered.
  UpdatePurchasedSkiPassAmount() {
    this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                         this.purchase_range.getColumn()])
  }
  PurchasedSkiPassAmount() { return this.purchased_amount }

  // When this method run, we return the total amount paid a given skipass kind 
  // that we fetch at the row for this ski pass but in column G so two columns
  // to the right of the purchased amount column
  GetTotalPrice() {
    return getNumberAt([this.purchase_range.getRow(),
                        this.purchase_range.getColumn() + 2])
  }

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

  IsStudent() { return this.student}
  IsAdult() { return this.adult }
}

// FIXME: All date intervals need to be adjustable globals. Include the YoB for an adult.
function createSkipassMap(sheet) {
  var to_return = {
    'Collet Senior': new SkiPass(
      'Collet Senior',
      sheet.getRange(25, 5),
      (dob) => {return ageVerificationRangeIncluded(dob, 70, 74)},
      "70 à 74 ans révolus"),
    'Collet Vermeil': new SkiPass(
      'Collet Vermeil',
      sheet.getRange(26, 5),
      (dob) => {return ageVerificationStrictlyOldOrOlder(dob, 75)},
      'plus de 75 ans'),
    'Collet Adulte': new SkiPass(
      'Collet Adulte',
      sheet.getRange(27, 5),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, new Date("December 31, 2005")) && ageVerificationStrictlyYounger(dob, 70)},
      "Adulte non étudiant de moins de 70 ans"),
    'Collet Étudiant': new SkiPass(
      'Collet Étudiant',
      sheet.getRange(28, 5),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, new Date("January 1, 1994"), new Date("December 31, 2005"))},
      '1er janvier 1994 et le 31 décembre 2005'),
    'Collet Junior': new SkiPass(
      'Collet Junior',
      sheet.getRange(29, 5),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, new Date("January 1, 2006"), new Date("December 31, 2013"))},
      '1er janvier 2006 et le 31 décembre 2013'),
    'Collet Enfant': new SkiPass(
      'Collet Enfant',
      sheet.getRange(30, 5),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, new Date("January 1, 2014"), new Date("December 31, 2018"))},
      "1er janvier 2014 et le 31 décembre 2018"),
    'Collet Bambin': new SkiPass(
      'Collet Bambin',
      sheet.getRange(31, 5),
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, new Date("January 1, 2019"))},
      'A partir du 1er Janvier 2019 et après'),

    '3 Domaines Senior': new SkiPass(
      '3 Domaines Senior',
      sheet.getRange(33, 5),
      (dob) => {return ageVerificationRangeIncluded(dob, 70, 74)},
      "70 à 74 ans révolus"),
    '3 Domaines Vermeil': new SkiPass(
      '3 Domaines Vermeil',
      sheet.getRange(34, 5),
      (dob) => {return ageVerificationStrictlyOldOrOlder(dob, 75)},
      'plus de 75 ans'),
    '3 Domaines Adulte': new SkiPass(
      '3 Domaines Adulte',
      sheet.getRange(35, 5),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, new Date("December 31, 2005")) && ageVerificationStrictlyYounger(dob, 70)},
      "Adulte non étudiant de moins de 70 ans"),
    '3 Domaines Étudiant': new SkiPass(
      '3 Domaines Étudiant',
      sheet.getRange(36, 5),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, new Date("January 1, 1994"), new Date("December 31, 2005"))},
      '1er janvier 1994 et le 31 décembre 2005'),
    '3 Domaines Junior': new SkiPass(
      '3 Domaines Junior',
      sheet.getRange(37, 5),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, new Date("January 1, 2006"), new Date("December 31, 2013"))},
      '1er janvier 2006 et le 31 décembre 2013'),
    '3 Domaines Enfant': new SkiPass(
      '3 Domaines Enfant',
      sheet.getRange(38, 5),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, new Date("January 1, 2014"), new Date("December 31, 2018"))},
      "1er janvier 2014 et le 31 décembre 2018"),
    '3 Domaines Bambin': new SkiPass(
      '3 Domaines Bambin',
      sheet.getRange(39, 5),
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, new Date("January 1, 2019"))},
      'A partir du 1er Janvier 2019 et après'),
  }
  validateClassInstancesMap(to_return, 'skipass_map')
  return to_return
}

var coord_rebate_family_of_4_amount = [23, 4]
var coord_rebate_family_of_4_count = [23, 5]
var coord_rebate_family_of_5_amount = [24, 4]
var coord_rebate_family_of_5_count = [24, 5]

function isSkipassStudent(ski_pass) {
  return (ski_pass == '3 Domaines Étudiant' || ski_pass == 'Collet Étudiant')
}

function isSkipassAdult(ski_pass) {
  return (ski_pass == '3 Domaines Adulte' || ski_pass == 'Collet Adulte')
}

function getSkiPassesPairs() {
  var to_return = []
  var entries = ['Senior', 'Vermeil', 'Adulte', 'Étudiant', 'Junior', 'Enfant', 'Bambin']
  entries.forEach(function(entry) {
    to_return.push(['Collet ' + entry, '3 Domaines ' + entry])
  })
  return to_return
}

// Identifier we retain as ski passes valid for competitor. Ideally, we should be adding
// the 3 domains
var competitor_ski_passes = [
  'Collet Adulte',
  'Collet Étudiant',
  'Collet Junior',
  'Collet Enfant',
]

// Email configuration - these shouldn't change very often
var allowed_user = 'inscriptions.sca@gmail.com'
var email_loisir = 'sca.loisir@gmail.com'
var email_comp = 'skicluballevardin@gmail.com'
var email_dev = 'apbianco@gmail.com'
var email_license = 'licence.sca@gmail.com'

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

function yolo() {
    if (getStringAt(coord_yolo) == "#yolo") {
      updateStatusBar("🍀 #yolo", "red")
      setStringAt(coord_yolo, "")
      return true
    }
    return false
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
    // Do not change the color for the number - it shouldn't be there
    // anyways (doing so would write back a string that would invalidate
    // formulas that dependend on the value of that cell)
    var column = coord[1] - 1
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
  if (getNumberAt(coord_charge) == 0) {
    setResetRebate(coord_charge, "white")
  }
  var blob = createPDF(spreadsheet) 
  var pdf_filename = spreadsheet.getName() + '-' + pdf_number + '.pdf';
  var file = savePDF(blob, pdf_filename)

  // Always for the rebate/charge area to be back in black 🤘.
  setResetRebate(coord_rebate, "black")
  setResetRebate(coord_charge, "black")

  var spreadsheet_folder_id =
    DriveApp.getFolderById(spreadsheet.getId()).getParents().next().getId()
  DriveApp.getFileById(file.getId()).moveTo(
    DriveApp.getFolderById(spreadsheet_folder_id))
  return file;
}

///////////////////////////////////////////////////////////////////////////////
// Spreadsheet data access methods
///////////////////////////////////////////////////////////////////////////////

function clearRange(coord) {
  SpreadsheetApp.getActiveSheet().getRange(coord[0],coord[1]).clearContent();
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
function ageVerificationBornAfterDateIncluded(dob, date) {
  // Apparently, the date formating in a cell sets the date to the first
  // hours of the day: 31/12/2008 will in fact be 31/12/2008, 1:00am. So we
  // remove one hour to get the exact date.
  var dob_valueof = dob.valueOf() - 3600000
  return dob_valueof >= date.valueOf()
}

// Return true if DoB happened before date
function ageVerificationBornBeforeDateIncluded(dob, date) {
  var dob_valueof = dob.valueOf() - 3600000
  return dob_valueof <= date.valueOf()
}

// Return true if DoB happened between first and last date included.
function ageVerificationBornBetweenDatesIncluded(dob, first, last) {
  var dob_valueof = dob.valueOf() - 3600000
  return (dob_valueof >= first.valueOf() && dob_valueof <= last.valueOf())
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
function ageVerificationStrictlyOldOrOlder(dob, age){
  return ageFromDoB(dob) >= age
}

// Return True if at the current time, someone with dob is strictly younger than age.
function ageVerificationStrictlyYounger(dob, age){
  return ageFromDoB(dob) < age
}

// Return True if at current tie, someone with dob is between age1 and age2 years old included.
function ageVerificationRangeIncluded(dob, age1, age2) {
  var age = ageFromDoB(dob)
  return (age >= age1 && age <= age2);
}

// Determine whether someone is a legal adult. An adult is someone who's
// age today is 18 year old (that includes turning 18 today)
function isAdult(dob) {
  return ageFromDoB(dob) >= 18;
}

// Determine whether is someone is an adult by year of birth only. For 2023/2024 season, someone is
// adult if they were born in 2005.
function isAdultByYoB(dob) {
  var yob = getDoBYear(dob)
  return yob <= 2005
}

function isMinor(dob) {
  return ! isAdult(dob)
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
// Input/output formating
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
  var to_return = str.trim().normalize("NFD").replace(/\p{Diacritic}/gu, "").
      replace(/\s/g, "-").  // No spaces in the middle
      replace(/\d+/g, "").  // No numbers
      replace(/\//g, "-").  // / into -
      replace(/\./g, "-").  // . into -
      replace(/_/g, "-").   // _ into -
      replace(/:/g, "-").   // : into -
      replace(/-+/g, "-").  // Many - into a single one
      replace(/-$/, "").    // No trailing
      replace(/\s$/, "")    // No trailing space
  if (to_upper_case) {
    to_return = to_return.toUpperCase()
  }
  return to_return
}

// Return a plural based on the value of number
function Plural(number, message) {
  if (number > 1) {
    message += 's'
  }
  return message
}

///////////////////////////////////////////////////////////////////////////////
// Alerting
///////////////////////////////////////////////////////////////////////////////

function displayErrorPanel(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert("❌ Erreur:\n\n" + message, ui.ButtonSet.OK);
}

// Display a OK/Cancel panel, returns true if OK was pressed. If the message
// carries the value indicating it's been previously answered with No, return
// the value associated to NO.
function displayYesNoPanel(message) {
  if (message == 'BUTTON_NO') {
    return false
  }
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
  if (value == '⚠️ Choix non renseigné' || value == '') {
    displayErrorPanel(message)
    return ''
  }
  return value
}

// Validation routines table. They must be running in a specific order defined here after.
// For some, failing is a  show stopper. For others, it's possible to decline accepting
// there's an error and continue (escape hatch). Most can be skipped when #yolo is active.
//
// ------------------------------+---------------------------------------------+----------------
// Name                          |What                                         |Esc. Hatch
//                               |                                             |RETURN:
// ------------------------------+---------------------------------------------+----------------
// validateFamilyMembers         | Validation of all family members, including | NO
//                               | some invariants like: junion non comp       | string if error
//                               | must have a level. If anything other than   | 
//                               | first/last name are defined, there must be  | 
//                               | a first/last name defined                   |
// ------------------------------+---------------------------------------------+----------------
// validateLicenses              | Validate selected licenses with the one     | NO
//                               | selected for payment.                       | string if error
// ------------------------------+---------------------------------------------+----------------
// validateNonCompSubscriptions  | Validate the non competitor subscriptions.  | YES
//                               |                                             | string if error
// ------------------------------+---------------------------------------------+----------------
// validateSkiPasses             | Validation of ski pass purchases.           | YES.
//                               |                                             | String if error
// ------------------------------+---------------------------------------------+----------------
// validateCompSubscriptions     | Competitor subscription validation          | NO
//                               | This code is complete                       | String if error
// ------------------------------+---------------------------------------------+----------------
//
// Execution order (desired, not actual: FIXME)
//
// - validateFamilyMembers()      | if error, bail
// - validateLicenses()           | if error, bail
// - validateCompSubscriptions    | if error, Yes/No
// - validateNonCompSubscriptions | if error, Yes/No
// - validateSkiPasses            | if error, Yes/No.

function TESTValidation() {
  function test(f) {
    var error = f()
    if (error) {
      displayErrorPannel(error)
    }
  }
  test(validateFamilyMembers)
  test(validateLicenses)
  test(validateNonCompSubscriptions)
  test(validateSkiPasses)
}

function TESTValidateInvoice() {
  function test(f) {
    var result = f()
    if (result == {}) {
      displayErrorPannel("Error during test")
    }
  }
  test(validateInvoice)
}

// Verify that family members are properly defined, especially with regard to
// what they are claiming to be associated to (level, type of license, etc...)
// When this succeeds, the family members should be defined in such a way that
// other validations are easier to implement.
function validateFamilyMembers() {
  updateStatusBar("Validation des données de la famille...", "grey", add=true)
  // Get a licene map so that we can immediately verify the age picked
  // for a license is correct.
  var license_map = createLicensesMap(SpreadsheetApp.getActiveSheet())
  for (var index in coords_identity_rows) {
    var first_name = getStringAt([coords_identity_rows[index], coord_first_name_column]);
    var last_name = getStringAt([coords_identity_rows[index], coord_last_name_column]);
    // Normalize the first name, normalize/upcase the last name
    first_name = normalizeName(first_name, false)
    last_name = normalizeName(last_name, true)

    // Get these early as we're going to verify that if they are defined, they are
    // attached to a name
    var license = getLicenseAt([coords_identity_rows[index], coord_license_column]);
    var dob = getDoB([coords_identity_rows[index], coord_dob_column]);
    var level = getLevelAt([coords_identity_rows[index], coord_level_column]);
    var sex = getStringAt([coords_identity_rows[index], coord_sex_column]);
    var city = getStringAt([coords_identity_rows[index], coord_cob_column]);
    var license_number = getStringAt([coords_identity_rows[index], coord_license_number_column ])

    // Entry is empty, just skip it unless there's information that has been 
    // entered
    if (first_name == "" && last_name == "") {
      if (isLicenseDefined(license) || dob != undefined || isLevelDefined(level) ||
          sex != '' || city != '' || license_number != '') {
        return (nthString(index) + " membre de famille: information définie (naissance, sexe, etc...) " +
                "sans prénom on nom de famille renseigné")        
      }
      continue;
    }
    // We need both a first name and a last name
    if (! first_name) {
      return ("Pas de nom de prénom fournit pour " + last_name);
    }
    if (! last_name) {
      return ("Pas de nom de famille fournit pour " + first_name);
    }
    // Write the first and last name back as it's been normalized
    setStringAt([coords_identity_rows[index], coord_first_name_column], first_name, "black");
    setStringAt([coords_identity_rows[index], coord_last_name_column], last_name, "black");

    // We need a DoB but only if a license or a level has been requested.
    if (dob == undefined) {
      if (isLicenseDefined(license) || isLevelDefined(level)) {
        return (
          "Licence ou niveau défini pour " + first_name + " " + last_name +
          " mais pas de date de naissance fournie ou date de naissance mal formatée (JJ/MM/AAAA)" +
          " ou année de naissance fantaisiste.");
      }
      continue
    }

    // We need a properly defined sex
    if (sex != "Fille" && sex != "Garçon") {
      return ("Pas de sexe défini pour " + first_name + " " + last_name);
    }

    // Level validation. We establishing here is that a junior non comp entry must have
    // a level set to something. The rest of the code assumes that a level set to something
    // is an other way of identifying a non comp license.
    //
    //   1- When a level is defined, it must have been validated
    //   2- When a level is defined, a license must be defined.
    //   3- A level must be defined for non competitor minor
    //   4- A competitor can not declare a level, it will confuse the rest of the validation
    //   5- An executive can not declare a level, it will confuse the rest of the validation
    //

    if (isLevelDefined(level) && isLevelNotAdjusted(level)) {
      return ("Ajuster le niveau de " + first_name + " " + last_name + " en précisant " +
              "le niveau de pratique pour la saison " + season)
    }

    if (isLevelDefined(level) && isLicenseNotDefined(license)) {
      return (first_name + " " + last_name + " est un loisir junior avec un niveau de ski " +
              "défini ce qui traduit l'intention de prendre une adhésion au club. Choisissez " +
              "une license appropriée (CN Jeune Loisir ou Famille) pour cette personne")
    }
    if (isLevelNotDefined(level) && isMinor(dob) && isLicenseNonComp(license)) {
      return (
        "Vous devez fournir un niveau pour le junior non compétiteur " + 
        first_name + " " + last_name + " qui prend une license " + license         
      )
    }
    if (isLevelDefined(level) && isLicenseComp(license)) {
      return (first_name + " " + last_name + " est un compétiteur. Ne définissez pas " +
              "de niveau pour un compétiteur")
    }
    if (isLevelDefined(level) && isExecLicense(license)) {
      return (first_name + " " + last_name + " est un cadre/dirigeant. Ne définissez pas " +
              "de niveau pour un cadre/dirigeant")      
    }

    // License validation:
    //
    // 1- A license must match the age of the person it's attributed to
    // 2- An exec license requires a city of birth
    // 3- An existing non exec license doesn't require a city of birth
    if (!license_map.hasOwnProperty(license)) {
      return (first_name + " " + last_name + "La licence attribuée '" + license +
              "' n'est pas une license existante !")
    }   
    // When a license for a family member, it must match the dob requirement
    if (isLicenseDefined(license) && ! license_map[license].ValidateDoB(dob)) {
      return (first_name + " " + last_name +
              ": l'année de naissance " + getDoBYear(dob) +
              " ne correspond critère de validité de la " +
              "license choisie: (" + license + "): " +
              license_map[license].ValidDoBRangeMessage() + '.')
    }
    if (isExecLicense(license) && city == '') {
      return (first_name + " " + last_name + ": la licence attribuée (" +
              license + ") requiert de renseigner une ville et un pays de naissance");
    }
    if (isLicenseDefined(license) && !isExecLicense(license) && city != '') {
      return (first_name + " " + last_name + ": la licence attribuée (" +
              license + ") ne requiert pas de renseigner une ville et un pays de naissance. " +
              "Supprimez cette donnée")
    }
  }
  return ''
}

// Cross check the attributed licenses with the ones selected for payment
function validateLicenses() {
  updateStatusBar("Validation du choix des licenses...", "grey", add=true)
  var license_map = createLicensesMap(SpreadsheetApp.getActiveSheet())
  // Collect the attributed licenses
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getLicenseAt([row, coord_license_column]);
    license_map[selected_license].IncrementAttributedLicenseCount()
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
  //    valid or not even with a family license purchased.
  var family_license_attributed_count = license_map[getNonCompFamilyLicenseString()].AttributedLicenseCount()
  if (family_license_attributed_count != 0) {
    // A least four declared participants
    if (family_license_attributed_count < 4) {
      return (
        "Il faut attribuer une licence famille à au moins 4 membres " +
        "d'une même famille. Seulement " + 
        family_license_attributed_count +
        " ont été attribuées.");
    }
    // Check that one family license was purchased.
    var family_license_purchased = license_map[getNonCompFamilyLicenseString()].PurchasedLicenseAmount()
    if (family_license_purchased != 1) {
      return (
        "Vous devez acheter une licence loisir famille, vous en avez " +
        "acheté pour l'instant " + family_license_purchased);
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
      return (
        "Le nombre de licence(s) '" + index + "' attribuée(s) (au nombre de " +
        license_map[index].AttributedLicenseCount() + ")\n" +
        "ne correspond pas au nombre de licence(s) achetée(s) (au nombre de " +
        license_map[index].PurchasedLicenseAmount() + ")")
    }
  }
  return ''
}

function validateCompSubscriptions() {
  updateStatusBar("Validation des adhésions compétition...", "grey", add=true)
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

// Validation of non competitor subscriptions. This method is invoked after all
// family members have been validated, which means that we are guaranteed here
// that people declared to have a non comp license also have a level entered.
function validateNonCompSubscriptions() {
  // Return true if:
  //   1- level_or_subscription is Rider
  //   2- level_or_subscription is FirstKid and there's one or more rider defined
  function skipRiderOrFirstKidIfRider(level_or_subscription) {
    return (isLevelRider(level_or_subscription) || 
            (isFirstKid(level_or_subscription) &&
             subscription_map[getRiderLevelString()].PurchasedSubscriptionAmount() > 0))
  }

  updateStatusBar("Validation des adhésions loisir...", "grey", add=true)
  var subscription_map = createNonCompSubscriptionMap(SpreadsheetApp.getActiveSheet())

  // Update the number of noncomp subscription registered
  for (var subscription in subscription_map) {
    subscription_map[subscription].UpdatePurchasedSubscriptionAmount()
  }

  // Compute the number of people with a level, making the distinction between riders/non riders.
  // These are the people that need a subscription.
  var rider_number = 0
  var non_rider_number = 0
  coords_identity_rows.forEach(function(row) {
    var level = getLevelAt([row, coord_level_column])
    if (isLevelRider(level)) {
      rider_number += 1
    }
    if (isLevelNotRider(level)) {
      non_rider_number += 1
    }
  })

  // The number of riders must be equal to the number of riders we found. First count them all
  // and the perform the verification
  var subscribed_rider_number = subscription_map[getRiderLevelString()].PurchasedSubscriptionAmount()
  if (rider_number != subscribed_rider_number) {
    return ("Le nombre d'adhésion(s) rider souscrite(s) (" + subscribed_rider_number + ") ne correspond pas au nombre " +
            "de rider(s) renseigné(s) (" + rider_number + ")")
  }

  // If we have a rider, the first subscription can not exist, we jump directly to the second kid
  var first_kid = subscription_map[getFirstKid()].PurchasedSubscriptionAmount()
  if (first_kid != 0 && rider_number > 0) {
    return ("L'adhésion rider compte comme une première Adhésion / Stage / Transport - 1er enfant. " +
            "Veuillez saisir les adhésion à partir du deuxièmme enfant.")
  }

  // 2- Go over the 1st to 4th subscription and make sure that N is matched by N-1 for
  //    N > 1. This verification is adjusted in case we have a rider as a first kid.
  var subscribed_non_rider_number_accumulator = 0
  for (var index in noncomp_subscription_categories) {
    // Skip Rider and skip first kid if we had a rider registered
    var subscription = noncomp_subscription_categories[index]
    if (skipRiderOrFirstKidIfRider(subscription)) {
      // If we're skipping first kid because there's one or more riders declared
      // we adjust the accumulator so that the rest of the verification can
      // happen.
      if (isFirstKid(subscription)) {
        subscribed_non_rider_number_accumulator = 1
      }
      continue
    }
    var current_purchased = subscription_map[subscription].PurchasedSubscriptionAmount()
    if (current_purchased == 1 && subscribed_non_rider_number_accumulator != (index - 1)) {
      return ('Une adhésion existe pour un ' + subscription +
              ' sans adhésion déclarée pour un ' + noncomp_subscription_categories[index-1])
    }
    subscribed_non_rider_number_accumulator += subscription_map[subscription].PurchasedSubscriptionAmount()
  }

  // 3- Go over 1st to 4th subscription and accumulate the number of subscriptions entered.
  //    that number must match the number of non rider member entered.
  var subscribed_non_rider_number = 0
  for (var subscription in subscription_map) {
    // Skip riders, we already verified them. This works because there's a subscription
    // that bears the level it matches ('Rider')
    if(skipRiderOrFirstKidIfRider(subscription)) {
      continue
    }
    var found_non_rider_number = subscription_map[subscription].PurchasedSubscriptionAmount()
    if (found_non_rider_number < 0 || found_non_rider_number > 1) {
      return ("La valeur du champ Adhésion / Stage / Transport pour le " +
              subscription + " ne peut prendre que la valeur 0 ou 1.")
    }
    subscribed_non_rider_number += found_non_rider_number
  }
  if (subscribed_non_rider_number != non_rider_number) {
    return ("Le nombre d'adhésion(s) NON rider souscrite(s) (" + subscribed_non_rider_number + ") " +
            "ne correspond pas au nombre " +
            "de NON rider(s) renseigné(s) (" + non_rider_number + ")")    
  }
  return ''
}

function validateSkiPasses() {
  updateStatusBar("Validation des forfaits loisir...", "grey", add=true)
  // Always clear the rebate section before eventually recomputing it at the end
  clearRange(coord_rebate_family_of_4_count)
  clearRange(coord_rebate_family_of_4_amount)
  clearRange(coord_rebate_family_of_5_count)
  clearRange(coord_rebate_family_of_5_amount)
  SpreadsheetApp.flush();

  var ski_passes_map = createSkipassMap(SpreadsheetApp.getActiveSheet())

  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var dob = getDoB([row, coord_dob_column])
    // Undefined DoB marks a non existing entry that we skip
    if (dob == undefined) {
      continue
    }
    var selected_license = getLicenseAt([row, coord_license_column]);
    // Increment the ski pass count that validates for a DoB. This will tell us
    // how many ski passes suitable for a non competitor we can expect to be
    // purchased.
    for (var skipass in ski_passes_map) {
      ski_passes_map[skipass].IncrementAttributedSkiPassCountIfDoB(dob)    
    }
  }

  // Collect the amount of skipasses that have be declared for purchase.
  // We also use this loop to compute the size of the familly to see if a rebate is
  // going to apply. A family member for a rebate is anyone who is not a student.
  var family_size = 0
  var number_of_non_student_adults = 0
  var total_paid_skipass = 0
  for (var skipass_name in ski_passes_map) {
    skipass = ski_passes_map[skipass_name]
    skipass.UpdatePurchasedSkiPassAmount()
    var purchased_amount = skipass.PurchasedSkiPassAmount()
    if (purchased_amount < 0 || isNaN(purchased_amount)) {
      return (purchased_amount + ' forfait ' + skipass_name + 
              ' acheté n\'est pas un nombre valide')
    }
    // Family size MUST exclude students
    if (! skipass.IsStudent() && purchased_amount > 0) {
      if (skipass.IsAdult()) {
        number_of_non_student_adults += purchased_amount
      }
      family_size += purchased_amount
      total_paid_skipass += skipass.GetTotalPrice() 
    }
  }

  var ski_passes_pairs = getSkiPassesPairs()
  for (index in ski_passes_pairs) {
    var pair = ski_passes_pairs[index]
    var zone1 = pair[0]
    var zone2 = pair[1]
    // Skip students in this verification, we'll do it separately
    if (isSkipassStudent(zone1)) {
      if (! isSkipassStudent(zone2)) {
        return ('Error interne: zone1=' + zone1 + ', zone2=' + zone2)
      }
      continue
    }
    var zone1_count = ski_passes_map[zone1].AttributedSkiPassCount()
    var zone2_count = ski_passes_map[zone2].AttributedSkiPassCount()
    // zone1_count should be exactly zone2_count. Verify this here.    
    if (zone1_count != zone2_count) {
      return ('Error interne: ' + zone1 + '=' + zone1_count + ', ' +
              zone2 + '=' + zone2_count)
    }
    var total_purchased = (ski_passes_map[zone1].PurchasedSkiPassAmount() + 
                           ski_passes_map[zone2].PurchasedSkiPassAmount())
    // There's a discrepancy between the number of qualified person in a ski pass 
    // DoB validity range and the amount purchased. Offer here to continue the verification
    // or return indicating that the entered value must be checked. This can happen for
    // instance when one adult purchases a license but doesn't purchase a ski pass...
    if (zone1_count != total_purchased) {
      var message = (total_purchased + Plural(total_purchased, ' forfait') + ' ' +
                     zone1 + '/' + zone2 + Plural(total_purchased, ' acheté') + ' pour ' +
                     zone1_count + Plural(zone1_count, ' personne') +
                     ' dans cette tranche d\'âge (' + ski_passes_map[zone1].ValidDoBRangeMessage() + ")")
      message += ("\n\nIl ce peut que ce choix de forfait soit valide mais pas " +
                  "automatiquement vérifiable. C'est le cas lorsque plusieurs options  " +
                  "correctes existent comment par exemple un membre étudiant qui est aussi " +
                  "d'âge adulte ou un adulte qui prend un license mais pas de forfait.")                     
      if (! displayYesNoPanel(augmentEscapeHatch(message))) {
        return 'BUTTON_NO'
      }
    }
  }

  // FIXME: Validation of the students

  if (number_of_non_student_adults == 2) {
    if (family_size == 4) {
      var rebate = -(total_paid_skipass * 0.1)
      setStringAt(coord_rebate_family_of_4_count , "1")
      setStringAt(coord_rebate_family_of_4_amount, rebate)
    }
    if (family_size >= 5) {
      var rebate = -(total_paid_skipass * 0.15)
      setStringAt(coord_rebate_family_of_5_count , "1")
      setStringAt(coord_rebate_family_of_5_amount, rebate)    
    }
    SpreadsheetApp.flush();
  }
  return ''
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
    var human_readable_coord = SpreadsheetApp.getActiveSheet().getRange(coord_version[0],
                                                                    coord_version[1]).getA1Notation()
    displayErrorPanel("Problème lors de la génération du numéro de document\n" +
                      "Insérez 'version=99' en " + human_readable_coord + " et recommencez l'opération");
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

// Family member description class returned by getListOfFamilyPurchasingALicense.
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
// with a link and some context.
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

function augmentEscapeHatch(source) {
  // If the source indicates the error has already acknowledged as one, return
  // the same error. Otherwise, wrap it around a message.
  if (source == 'BUTTON_NO') {
    return source
  }
  return source + ("\n\nChoisissez 'OK' pour continuer à générer la facture.\n" +
                  "Choisissez 'Annuler' pour ne pas générer la facture et " +
                  "vérifier les valeurs saisies...");
}

// Invoice validation data returned by validateInvoice() after a successful
// invoice validation.
class InvoiceValidationData {
  constructor(error, civility, family_name, mail_to, legal_disclaimer, medical_form) {
    this.error = error
    this.civility = civility
    this.family_name = family_name
    this.mail_to = mail_to
    this.legal_disclaimer = legal_disclaimer
    this.medical_form = medical_form
  }
}

// Validate the invoice and return a dictionary of values
// to use during invoice generation.
function validateInvoice() {
  function validatationDataError() {
    return new InvoiceValidationData(true, '', '', '', '', '')
  }
  if (! isProd()) {
    Debug("Cette facture est en mode developpement. " +
          "Aucun email ne sera envoyé, " +
          "ni à la famile ni à " + email_license + ".\n\n" +
          "Vous pouvez néamoins continuer et un dossier sera préparé et " +
          "les mails serons envoyés à " + email_dev + ".\n\n" +
          "Contacter " + email_dev + " pour obtenir plus d\'aide.")
  }
  
  // Reformat the phone numbers  
  formatPhoneNumbers();

  // Make sure only an allowed user runs this.
  if (Session.getEffectiveUser() != allowed_user) {
    displayErrorPanel(
      "Vous n'utilisez pas cette feuille en tant que " +
      allowed_user + ".\n\n" +
      "Veuillez vous connecter d'abord à ce compte avant " +
      "d'utiliser cette feuille.");
    return validatationDataError();
  }
  
  updateStatusBar("Validation des coordonées...", "grey")
  // Validation: proper civility
  var civility = validateAndReturnDropDownValue(
    coord_family_civility,
    "Vous n'avez pas renseigné de civilité")
  if (civility == '') {
    return validatationDataError()
  }
  
  // Validation: a family name
  var family_name = getStringAt(coord_family_name)
  if (family_name == '') {
    displayErrorPanel(
      "Vous n'avez pas renseigné de nom de famille ou " +
      "vous avez oublié \n" +
      "de valider le nom de famille par [return] ou [enter]...")
    return validatationDataError()
  }

  // Validation: a correct address
  var street_address = getStringAt(coord_family_street)
  var zip_number = getStringAt(coord_family_zip)
  var city = getStringAt(coord_family_city)
  if (street_address == "" || zip_number == "" || city == "") {
    displayErrorPanel(
      "Vous n'avez pas correctement renseigné une adresse de facturation:\n" +
      "numéro de rue, code postale ou commune - ou " +
      "vous avez oublié \n" +
      "de valider une valeur entrée par [return] ou [enter]...")
    return validatationDataError()
  }

  // Validation: first phone number
  var phone_number = getStringAt(coord_family_phone1)
  if (phone_number == '') {
    displayErrorPanel(
      "Vous n'avez pas renseigné de nom de numéro de telephone ou " +
      "vous avez oublié \n" +
      "de valider le numéro de téléphone par [return] ou [enter]...")
    return validatationDataError()   
  }

  // Validation: proper email adress.
  var mail_to = getStringAt(coord_family_email)
  if (mail_to == '' || mail_to == '@') {
    displayErrorPanel(
      "Vous n'avez pas saisi d'adresse email principale ou " +
      "vous avez oublié \n" +
      "de valider l'adresse email par [return] ou [enter]...")
    return validatationDataError()
  }
  
  var validation_error = ''
  var is_yolo = yolo()
  // Validation phase that can be entirely skipped if #yolo
  if (! is_yolo) {
    // Validate all the entered family members. This will make sure that
    // a non comp license is matched by a level, something that the validation
    // of non comp subscription depends on.
    validation_error = validateFamilyMembers();
    if (validation_error) {
      displayErrorPanel(validation_error);
      return validatationDataError()
    }

    // Now performing the optional/advanced validations.
    if (advanced_validation.AdvancedVerificationFamilyLicenses()) {
      validation_error = validateLicenses();
      if (validation_error) {
        displayErrorPanel(validation_error);
        return validatationDataError()
      }
    }

    // Validate the competitor subscriptions
    validation_error = validateCompSubscriptions()
    if (validation_error) {
      if (! displayYesNoPanel(augmentEscapeHatch(validation_error))) {
        return validatationDataError()
      }      
    }

    // 2- Verify the subscriptions. The operator may choose to continue
    //    as some situation are un-verifiable automatically.
    if (advanced_validation.AdvancedVerificationSubscription()) {
      // Validate requested licenses and subscriptions
      validation_error = validateNonCompSubscriptions()
      if (validation_error) {
        if (! displayYesNoPanel(augmentEscapeHatch(validation_error))) {
          return validatationDataError()
        }
      }
    }

    // 3- Verify the ski pass purchases. The operator may choose to continue
    //    as some situation are un-verifiable automatically.
    if (advanced_validation.AdvancedVerificationSkipass()) {
      validation_error = validateSkiPasses()
      if (validation_error) {
        if (! displayYesNoPanel(augmentEscapeHatch(validation_error))) {
          return validatationDataError()
        }    
      }
    }
 
    // Validate the legal disclaimer.
    updateStatusBar("Validation autorisation/questionaire...", "grey", add=true)
    var legal_disclaimer_validation = validateAndReturnDropDownValue(
      coord_legal_disclaimer,
      "Vous n'avez pas renseigné la nécessitée ou non de devoir " +
      "recevoir la mention légale.");
    if (legal_disclaimer_validation == '') {
      return validatationDataError()
    }
    var legal_disclaimer_value = getStringAt(coord_legal_disclaimer);
    if (legal_disclaimer_value == 'Non fournie') {
      displayErrorPanel(
        "La mention légale doit être signée aujourd'hui pour " +
        "valider le dossier et terminer l'inscription");
      return validatationDataError()
    }

    // Validate the medical form
    var medical_form_validation = validateAndReturnDropDownValue(
      coord_medical_form,
      "Vous n'avez pas renseigné votre réponse (OUI/NON) au questionaire médicale.")
    if (medical_form_validation == '') {
      return validatationDataError()
    }
  }

  // Update the timestamp. 
  updateStatusBar("Mise à jour de la version...", "grey", add=true)
  setStringAt(coord_timestamp,
              'Dernière MAJ le ' +
              Utilities.formatDate(new Date(),
                                   Session.getScriptTimeZone(),
                                   "dd-MM-YY, HH:mm") + (is_yolo ? "\n#yolo" : ""), 'black')


  return new InvoiceValidationData(false, civility, family_name, checkEmail(mail_to),
                                   legal_disclaimer_validation, medical_form_validation)    
}

///////////////////////////////////////////////////////////////////////////////
// Invoice emailing
///////////////////////////////////////////////////////////////////////////////

function maybeEmailLicenseSCA(invoice) {
  var operator = getOperator()
  var family_name = getFamilyName()
  var family_dict = getListOfFamilyPurchasingALicense() 
  var string_family_members = "";
  var license_count = 0
  for (var index in family_dict) {
    if (family_dict[index].last_name == "") {
      continue
    }
    license_number = family_dict[index].license_number
    if (license_number == '') {
      license_number = '<font color="red"><b>À créer</b></font>'
    }
    string_family_members += (
      "<tt>" +
      "Nom: <b>" + family_dict[index].last_name.toUpperCase() + "</b><br>" +
      "Prénom: " + family_dict[index].first_name + "<br>" +
      "Naissance: " + Utilities.formatDate(family_dict[index].dob, Session.getScriptTimeZone(), 'dd-MM-YYYY') + "<br>" +
      "Fille/Garçon: " + family_dict[index].sex + "<br>" +
      "Ville de Naissance: " + family_dict[index].city + "<br>" +
      "Licence: " + family_dict[index].license_type + "<br>" +
      "Numéro License: " + license_number + "<br>" +
      "----------------------------------------------------</tt><br>\n");
      license_count += 1;
  }
  if (string_family_members) {
    string_family_members = (
      "<p>" + license_count + Plural(license_count, " licence")
      + Plural(license_count, " nécessaire") + " pour:</p><blockquote>\n" +
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

function updateStatusBar(message, color, add=false) {
  var content = message
  if (add) {
    message = getStringAt(coord_status) + ' ✔\n' + message
    var messages = message.split("\n")
    if (messages.length > 3) {
      messages.splice(0, 1)
      message = messages.join("\n")
    }
  }
  setStringAt(coord_status, message, color)
  SpreadsheetApp.flush()
}

function generatePDFAndMaybeSendEmail(send_email, just_the_invoice) {
  updateStatusBar("⏳ Validation de la facture...", "orange")      
  var validation = validateInvoice();
  if (validation.error) {
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
  
  var civility = validation.civility
  var family_name = validation.family_name
  var mail_to = validation.mail_to
  var legal_disclaimer_validation = validation.legal_disclaimer
  var medical_form_validation = validation.medical_form
  
  // Determine whether parental consent needs to be generated. If
  // that's the case, we generate additional attachment content.
  var legal_disclaimer_text = ''
  if (! just_the_invoice) {
    if (legal_disclaimer_validation == 'À fournire signée') {
      attachments.push(
        DriveApp.getFileById(legal_disclaimer_pdf).getAs(MimeType.PDF))
    
      legal_disclaimer_text = (
        "<p>Il vous faut compléter, signer et nous retourner la mention " +
        "légale fournie en attachment, couvrant le droit à l'image, le " +
        "règlement intérieur, les interventions médicales et la GDPD.</p>");
    }

    // Insert the note and rules for the parents anyways
    attachments.push(DriveApp.getFileById(rules_pdf).getAs(MimeType.PDF))
    attachments.push(DriveApp.getFileById(parents_note_pdf).getAs(MimeType.PDF));
    legal_disclaimer_text += (
      "<p>Vous trouverez également en attachement une note adressée aux " +
      "parents, ainsi que le règlement intérieur. Merci de lire ces deux " +
      "documents attentivement.</p>");
  }
  
  // Take a look at the medical form answer:
  // 1- Yes: we need to tell that a medical certificate needs to be provided
  // 2- No: a new attachment need to be added.
  var medical_form_text = ''
  if (medical_form_validation == 'Une réponse OUI') {
    medical_form_text = ('<p><b><font color="red">' +
                         'La ou les réponses positives que vous avez porté au questionaire médical vous ' +
                         'obligent à transmettre au SCA (inscriptions.sca@gmail.com) dans les plus ' +
                         'brefs délais <u>un certificat médical en cours de validité</u>.' +
                         '</font></b>')
  } else if (medical_form_validation == 'Toutes réponses NON') {
    medical_form_text = ('<p><b><font color="red">' +
                         'Les réponses négative que vous avez porté au questionaire médical vous ' +
                         'dispense de fournir un certificat médical mais vous obligent à <u>signer la page ' + ffs_information_leaflet_page +
                         ' de la notice d\'informations FFS ' + season + '</u> fournie en attachement.' +
                         '</font></b>')
    attachments.push(DriveApp.getFileById(ffs_information_leaflet_pdf).getAs(MimeType.PDF))
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
    
      legal_disclaimer_text +
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

  // Add CC if defined - do not collect a CC in test/dev mode
  if (!(isTest() || isDev())) {
    var cc_to = getStringAt(coord_cc)
    if (cc_to != "") {
      email_options.cc = checkEmail(cc_to)
    }
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
