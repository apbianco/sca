// Version: 2025-06-04 - Preparing 2025/2026
// https://github.com/apbianco/sca/commit/<FIXME>
//
// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.
//
// Important note on development and tests:
// - Development: use Firefox only while inscriptions.sca@gmail.com is
//   the only account that is connected
// - Test: use Safari only while inscriptions.sca@gmail.com is
//   the only account that is connected. Once an invoice has been generated,
//   it's possible to debug its code in Safari.
// - Automated tests exists in invoice_test.gs
// - Seasonal configuration points are in the file config.ps
//
// Dev or prod? "dev" sends all email to email_dev. Prod is the
// real thing: family will receive invoices, and so will email_license,
// unless the trigger in use is the TEST trigger.
dev_or_prod = "prod"

///////////////////////////////////////////////////////////////////////////////
// Advanced functionality and features management
///////////////////////////////////////////////////////////////////////////////

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

// Validation method use when creating license, skipasses and subscription
// maps.
// FIXME: Move this somewhere elsea
function validateClassInstancesMap(map, map_name) {
  for (var key in map) {
    if (map[key].Name() != key) {
      displayErrorPanel('Erreur interne de validation pour ' + map_name + ': ' +
                        map_name + '[' + key + '].Name() = ' + map[key].Name())
    }
  }
}

///////////////////////////////////////////////////////////////////////////////
// License management code
///////////////////////////////////////////////////////////////////////////////

// Definition of all possible license values
function getNoLicenseString() {return 'Aucune'}
function getNonCompJuniorLicenseString() { return 'CN Jeune (Loisir)'}
function getNonCompAdultLicenseString() { return 'CN Adulte (Loisir)'}
function getNonCompFamilyLicenseString() {return 'CN Famille (Loisir)'}
function getExecutiveLicenseString() {return 'CN Dirigeant'}
function getCompJuniorLicenseString() {return 'CN Jeune (Compétition)'}
function getCompAdultLicenseString() {return 'CN Adulte (Compétition)'}

function isLicenseDefined(license) {
  return license != getNoLicenseString() && license in createLicensesMap(SpreadsheetApp.getActiveSheet())
}

function isLicenseNoLicense(license) {
  return license == getNoLicenseString()
}

function isLicenseNotDefined(license) {
  return !isLicenseDefined(license)
}

function isLicenseNonComp(license) {
  // FIXME: Why is it not !isLicenseComp() ?
  return (license == getNonCompJuniorLicenseString() ||
          license == getNonCompFamilyLicenseString() ||
          license == getNonCompAdultLicenseString())
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

function isLicenseExec(license) {
  return license == getExecutiveLicenseString()
}

function isLicenseFamily(license) {
  return license == getNonCompFamilyLicenseString()
}

// A License class to create an object that has a name, a range in the trix at
// which it can be marked as purchased, a validation method that takes a DoB
// as input and a message to issue when the validation failed.
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
  UpdatePurchasedLicenseAmountFromTrix() {
    if (this.purchase_range != null) {
      this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                           this.purchase_range.getColumn()])
    }
  }
  SetPurchasedLicenseAmount() {
    // Update only if we have a number and if the data is updatable (no license for
    // instance can be counted but can't be updated at a given cell)
    if (this.purchase_range == null) {
      return
    }
    // If the license is a familly license, always adjust the non zero count
    // back to zero as there can only be one attributed.
    var count = this.AttributedLicenseCount()
    var coord = [this.purchase_range.getRow(), this.purchase_range.getColumn()]
    if (count > 0 && isLicenseFamily(this.Name())) {
      count = 1
    }
    if (count > 0) {
      this.purchased_amount = count
      setStringAt(coord, count)
    } else {
      setStringAt(coord, "")
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

// Get a license at coord and DONT normalize it's value. Only use this
// method to fetch a license if you know what you're doing :)
function getRawLicenseAt(coord) {
  return getStringAt(coord)
}

// Get a license at coord and normalize it's value. It's important to normalize
// the license string as it can be used as a key to a license_map. The normalized
// value for a license is 'Aucune'.
function getLicenseAt(coord) {
  var license = getStringAt(coord)
  if(isLicenseNotDefined(license)) {
    license = getNoLicenseString()
  }
  return license
}

// Create a map of all existing licenses
function createLicensesMap(sheet) {
  function getYear(license) {
    if (! license in licenses_configuration_map) {
      displayErrorPanel(license + " n'est pas dans licenses_configuration_map!")
      return
    } 
    return licenses_configuration_map[license][0]
  }
  function createDate(license, day) {
    var date = day + ", " + getYear(license)
    return new Date(date)
  }
  function getRange(license) {
    return sheet.getRange(licenses_configuration_map[license][1],
                          licenses_configuration_map[license][2])
  }
  var to_return = {
    'Aucune': new License(
      getNoLicenseString(),
      null,
      (dob) => {return true}),
    'CN Jeune (Loisir)': new License(
      getNonCompJuniorLicenseString(),
      getRange(getNonCompJuniorLicenseString()),
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, createDate(getNonCompJuniorLicenseString(), "January 1"))},
      "requiert d'être né en " + getYear(getNonCompJuniorLicenseString()) + " et après"),
    'CN Adulte (Loisir)': new License(
      getNonCompAdultLicenseString(),
      getRange(getNonCompAdultLicenseString()),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, createDate(getNonCompAdultLicenseString(), "December 31"))},
      "requiert d'être né en " + getYear(getNonCompAdultLicenseString()) + " et avant"),
    'CN Famille (Loisir)': new License(
      getNonCompFamilyLicenseString(),
      getRange(getNonCompFamilyLicenseString()),
      (dob) => {return true},
      ""),
    'CN Dirigeant': new License(
      getExecutiveLicenseString(),
      getRange(getExecutiveLicenseString()),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, createDate(getExecutiveLicenseString(), "December 31"))},
      "requiert d'être né en " + getYear(getNonCompAdultLicenseString()) + " et avant"),
    'CN Jeune (Compétition)': new License(
      getCompJuniorLicenseString(),
      getRange(getCompJuniorLicenseString()),
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, createDate(getCompJuniorLicenseString(), "January 1"))},
      "requiert être né en " + getYear(getCompJuniorLicenseString()) + " et après"),
    'CN Adulte (Compétition)': new License(
      getCompAdultLicenseString(),
      getRange(getCompAdultLicenseString()),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, createDate(getCompAdultLicenseString(), "December 31"))},
      "requiert d'être né en " + getYear(getCompAdultLicenseString()) + " et avant"),
  }
  validateClassInstancesMap(to_return, 'license_map')
  return to_return
}

///////////////////////////////////////////////////////////////////////////////
// Competitor categories management
///////////////////////////////////////////////////////////////////////////////

// Competitors: categories definition (used to establish the subscription
// range). Also define an array with these values as it's needed everywhere.
// Also define a categories sorting function, used to sort arrays of categories,
// verifying that comp_subscription_categories.sort() == comp_subscription_categories
function getU8String() { return 'U8' }
function getU10String() { return 'U10'}
function getU12PlusString() { return 'U12+' }
var comp_subscription_categories = [
  getU8String(), getU10String(), getU12PlusString()
]

function categoriesAscendingOrder(cat1, cat2) {
  // No order change
  if (cat1 == cat2) {
    return 0
  }
  // U8 is always the smallest category
  if (cat1 == getU8String()) {
    return -1
  }
  // U10 will only rank lower than U12+
  if (cat1 == getU10String()) {
    if (cat2 == getU12PlusString()) {
      return -1
    }
    return 1
  }
  // cat1 is U12, it's always last unless it's 
  // compared against U12 which has already been done
  return 1
}

///////////////////////////////////////////////////////////////////////////////
// Non competitor levels and subscription values management
///////////////////////////////////////////////////////////////////////////////

// Note: Undertermined level is not absence of level. Absence of level is
// getNALevelString or ''
function getNoLevelString() { return 'Non déterminé' } // DOES define an undetermined
function getNALevelString() { return 'Pas Concerné' }  // DOEST NOT define a level
function getLevelCompString() {return "Compétiteur" }
function getRiderLevelString() { return 'Rider' }
function getFirstKidString() { return '1er enfant' }
function getSecondKidString() { return '2ème enfant' }
function getThirdKidString() { return '3ème enfant' }
function getFourthKidString() { return '4ème enfant' }

var noncomp_subscription_categories = [
  getRiderLevelString(), getFirstKidString(), getSecondKidString(),
  getThirdKidString(), getFourthKidString()
]

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

function isLevelDefined(level) {
  return ! isLevelNotDefined(level)
}

function isLevelComp(level) {
  return level == getLevelCompString()
}

function isLevelNotComp(level) {
  return isLevelDefined(level) && ! isLevelComp(level);
}

function isLevelRider(level) {
  return level == getRiderLevelString()
}

function isLevelNotRider(level) {
  return isLevelDefined(level) && ! isLevelRider(level)
}

function isFirstKid(subscription) {
  return isLevelDefined(subscription) && subscription == getFirstKidString()
}

///////////////////////////////////////////////////////////////////////////////
// Subscription management code
///////////////////////////////////////////////////////////////////////////////

// A Subscription class to create an object that has a name, a range in the trix at
// which it can be marked as purchased, a validation method that takes a DoB
// as input and can keep track of occurences and purchases.
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
  UpdatePurchasedSubscriptionAmountFromTrix() {
    this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                         this.purchase_range.getColumn()])
  }
  PurchasedSubscriptionAmount() { return this.purchased_amount }
  SetPurchasedSubscriptionAmount(amount) {
    // Clear a cell with no information
    var coord = [this.purchase_range.getRow(), this.purchase_range.getColumn()]
    if (amount <= 0) {
      setStringAt(coord ,'')
      return
    }
    this.purchased_amount = amount
    setStringAt(coord, amount)
  }

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

// Definition of all possible subscription values
function createCompSubscriptionMap(sheet) {
  function getFirstYear(label) {
    if (! label in comp_subscription_map) {
      displayErrorPanel(label + " n'est pas dans comp_subscription_map!")
      return
    }
    return comp_subscription_map[label][0]
  }
  function getLastYear(label) {
    return comp_subscription_map[label][1]
  }
  var row = coord_comp_start_row
  var to_return = {}
  for (var rank = 1; rank <= comp_kids_per_family; rank +=1) {
    var label = rank + getU8String()
    to_return[label] = new Subscription(
      label,
      // FIXME: Should be attached to comp_subscription_map
      sheet.getRange(row, 5),
      (dob) => {return ageVerificationBornBetweenYearsIncluded(dob, getFirstYear(getU8String()), getLastYear(getU8String()))})
    row += 1;
    label = rank + getU10String()
    to_return[label] = new Subscription(
      label,
      sheet.getRange(row, 5),
      (dob) => {return ageVerificationBornBetweenYearsIncluded(dob, getFirstYear(getU10String()), getLastYear(getU10String()))})
    row += 1
    label = rank + getU12PlusString()
    to_return[label] = new Subscription(
      label,
      sheet.getRange(row, 5),
      (dob) => {return ageVerificationBornBeforeYearIncluded(dob, getFirstYear(getU12PlusString()))})
    row += 1
  }
  validateClassInstancesMap(to_return, 'createCompSubscriptionMap')
  return to_return
}

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

///////////////////////////////////////////////////////////////////////////////
// Skipass management code
///////////////////////////////////////////////////////////////////////////////

// Definition of all possible skipass values
function getSkiPassSenior() { return 'Senior' }
function getSkiPassSuperSenior() { return 'Vermeil' }
function getSkiPassAdult() { return 'Adulte' }
function getSkiPassStudent() { return 'Étudiant' }
function getSkiPassJunior() { return 'Junior' }
function getSkiPassKid() { return 'Enfant' }
function getSkiPassToddler() { return 'Bambin' }
function localizeSkiPassCollet(pass) { return 'Collet ' + pass}
function localizeSkiPass3D(pass) { return '3 Domaines ' + pass}
function isSkipPassLocalizedCollet(pass) {
  return pass.substring(0, 7) == 'Collet '
}

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
  UpdatePurchasedSkiPassAmountFromTrix() {
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
  SetPurchasedSkiPassAmount() {
    this.purchased_amount = this.occurence_count
    var coord = [this.purchase_range.getRow(),
                 this.purchase_range.getColumn()]
    if (this.purchased_amount > 0) {
      setStringAt(coord, this.purchased_amount)
    } else {
      setStringAt(coord, '')
    }
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

function createSkipassMap(sheet) {
  function getFirstValue(label) {
    if (! label in skipass_configuration_map) {
      displayErrorPanel(label + " n'est pas dans comp_subscription_map!")
      return
    }
    return skipass_configuration_map[label][0]
  }
  function getSecondValue(label) {
    return skipass_configuration_map[label][1]
  }
  // Early dates are starting January 1st,
  function getEarlyDate(label) {
    if (! label in skipass_configuration_map) {
      displayErrorPanel(label + " n'est pas dans comp_subscription_map!")
      return
    }
    return new Date("January 1, " + skipass_configuration_map[label][0])
  }
  // Late dates are ending December 31st.
  function getLateDate(label) {
    return new Date("December 31, " + skipass_configuration_map[label][1])
  }

  function getSkipassConfigRow(label, d3=false) {
    if (!(label in skipass_configuration_map)) {
      displayErrorPanel(label + " n'est pas dans comp_subscription_map!")
      return
    }
    return skipass_configuration_map[label][2] + (d3 ? skipass_configuration_map_3d_row_offset : 0)
  }

  function getSkipassConfigCol(label) {
    return skipass_configuration_map[label][3];
  }

  var to_return = {
    'Collet Senior': new SkiPass(
      localizeSkiPassCollet(getSkiPassSenior()),
      sheet.getRange(getSkipassConfigRow(getSkiPassSenior()), getSkipassConfigCol(getSkiPassSenior())),
      // Remember: no local variable capture possible in functor, use function calls only
      (dob) => {return ageVerificationRangeIncluded(dob, getFirstValue(getSkiPassSenior()), getSecondValue(getSkiPassSenior()))},
      getFirstValue(getSkiPassSenior()) + " à " + getSecondValue(getSkiPassSenior()) + " ans révolus"),
    'Collet Vermeil': new SkiPass(
      localizeSkiPassCollet(getSkiPassSuperSenior()),
      sheet.getRange(getSkipassConfigRow(getSkiPassSuperSenior()), getSkipassConfigCol(getSkiPassSuperSenior())),
      (dob) => {return ageVerificationStrictlyOldOrOlder(dob, getFirstValue(getSkiPassSuperSenior()))},
      "plus de "  + getFirstValue(getSkiPassSuperSenior()) + " ans"),
    'Collet Adulte': new SkiPass(
      localizeSkiPassCollet(getSkiPassAdult()),
      sheet.getRange(getSkipassConfigRow(getSkiPassAdult()), getSkipassConfigCol(getSkiPassAdult())),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, getEarlyDate(getSkiPassAdult())) &&
                       ageVerificationStrictlyYounger(dob, getSecondValue(getSkiPassAdult()))},
      "Adulte non étudiant de moins de " + getSecondValue(getSkiPassAdult()) + " ans"),
    'Collet Étudiant': new SkiPass(
      localizeSkiPassCollet(getSkiPassStudent()),
      sheet.getRange(getSkipassConfigRow(getSkiPassStudent()), getSkipassConfigCol(getSkiPassStudent())),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, getEarlyDate(getSkiPassStudent()), getLateDate(getSkiPassStudent()))},
      "1er janvier " + getFirstValue(getSkiPassStudent()) + " et le 31 décembre " + getSecondValue(getSkiPassStudent())),
    'Collet Junior': new SkiPass(
      localizeSkiPassCollet(getSkiPassJunior()),
      sheet.getRange(getSkipassConfigRow(getSkiPassJunior()), getSkipassConfigCol(getSkiPassJunior())),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, getEarlyDate(getSkiPassJunior()), getLateDate(getSkiPassJunior()))},
      "1er janvier " + getFirstValue(getSkiPassJunior()) + " et le 31 décembre " + getSecondValue(getSkiPassJunior())),
    'Collet Enfant': new SkiPass(
      localizeSkiPassCollet(getSkiPassKid()),
      sheet.getRange(getSkipassConfigRow(getSkiPassKid()), getSkipassConfigCol(getSkiPassKid())),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, getEarlyDate(getSkiPassKid()), getLateDate(getSkiPassKid()))},
      "1er janvier " + getFirstValue(getSkiPassKid()) + " et le 31 décembre " + getSecondValue(getSkiPassKid())),
    'Collet Bambin': new SkiPass(
      localizeSkiPassCollet(getSkiPassToddler()),
      sheet.getRange(getSkipassConfigRow(getSkiPassToddler()), getSkipassConfigCol(getSkiPassToddler())),
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, getEarlyDate(getSkiPassToddler()))},
      "A partir du 1er Janvier " + getFirstValue(getSkiPassKid()) + " et après"),

    '3 Domaines Senior': new SkiPass(
      localizeSkiPass3D(getSkiPassSenior()),
      sheet.getRange(getSkipassConfigRow(getSkiPassSenior(), true), getSkipassConfigCol(getSkiPassSenior())),
      // Remember: no local variable capture possible in functor, use function calls only
      (dob) => {return ageVerificationRangeIncluded(dob, getFirstValue(getSkiPassSenior()), getSecondValue(getSkiPassSenior()))},
      getFirstValue(getSkiPassSenior()) + " à " + getSecondValue(getSkiPassSenior()) + " ans révolus"),
    '3 Domaines Vermeil': new SkiPass(
      localizeSkiPass3D(getSkiPassSuperSenior()),
      sheet.getRange(getSkipassConfigRow(getSkiPassSuperSenior(), true), getSkipassConfigCol(getSkiPassSuperSenior(), true)),
      (dob) => {return ageVerificationStrictlyOldOrOlder(dob, getFirstValue(getSkiPassSuperSenior()))},
      "plus de "  + getFirstValue(getSkiPassSuperSenior()) + " ans"),
    '3 Domaines Adulte': new SkiPass(
      localizeSkiPass3D(getSkiPassAdult()),
      sheet.getRange(getSkipassConfigRow(getSkiPassAdult(), true), getSkipassConfigCol(getSkiPassAdult())),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, getEarlyDate(getSkiPassAdult())) &&
                       ageVerificationStrictlyYounger(dob, getSecondValue(getSkiPassAdult()))},
      "Adulte non étudiant de moins de " + getSecondValue(getSkiPassAdult()) + " ans"),
    '3 Domaines Étudiant': new SkiPass(
      localizeSkiPass3D(getSkiPassStudent()),
      sheet.getRange(getSkipassConfigRow(getSkiPassStudent(), true), getSkipassConfigCol(getSkiPassStudent())),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, getEarlyDate(getSkiPassStudent()), getLateDate(getSkiPassStudent()))},
      "1er janvier " + getFirstValue(getSkiPassStudent()) + " et le 31 décembre " + getSecondValue(getSkiPassStudent())),
    '3 Domaines Junior': new SkiPass(
      localizeSkiPass3D(getSkiPassJunior()),
      sheet.getRange(getSkipassConfigRow(getSkiPassJunior(), true), getSkipassConfigCol(getSkiPassJunior())),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, getEarlyDate(getSkiPassJunior()), getLateDate(getSkiPassJunior()))},
      "1er janvier " + getFirstValue(getSkiPassJunior()) + " et le 31 décembre " + getSecondValue(getSkiPassJunior())),
    '3 Domaines Enfant': new SkiPass(
      localizeSkiPass3D(getSkiPassKid()),
      sheet.getRange(getSkipassConfigRow(getSkiPassKid(), true), getSkipassConfigCol(getSkiPassKid())),
      (dob) => {return ageVerificationBornBetweenDatesIncluded(dob, getEarlyDate(getSkiPassKid()), getLateDate(getSkiPassKid()))},
      "1er janvier " + getFirstValue(getSkiPassKid()) + " et le 31 décembre " + getSecondValue(getSkiPassKid())),
    '3 Domaines Bambin': new SkiPass(
      localizeSkiPass3D(getSkiPassToddler()),
      sheet.getRange(getSkipassConfigRow(getSkiPassToddler(), true), getSkipassConfigCol(getSkiPassToddler())),
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, getEarlyDate(getSkiPassToddler()))},
      "A partir du 1er Janvier " + getFirstValue(getSkiPassKid()) + " et après"),
  }
  validateClassInstancesMap(to_return, 'skipass_map')
  return to_return
}

function isSkipassStudent(ski_pass) {
  return (ski_pass == localizeSkiPass3D(getSkiPassStudent()) ||
          ski_pass == localizeSkiPassCollet(getSkiPassStudent()))
}

function isSkipassAdult(ski_pass) {
  return (ski_pass == localizeSkiPass3D(getSkiPassAdult()) ||
          ski_pass == localizeSkiPassCollet(getSkiPassAdult()))
}

function getSkiPassesPairs() {
  var to_return = []
  var entries = [getSkiPassSenior(), getSkiPassSuperSenior(), getSkiPassAdult(),
                 getSkiPassStudent(), getSkiPassJunior(), getSkiPassKid(), getSkiPassToddler()]
  entries.forEach(function(entry) {
    to_return.push([localizeSkiPassCollet(entry), localizeSkiPass3D(entry)])
  })
  return to_return
}

///////////////////////////////////////////////////////////////////////////////
// Email address configuration points.
///////////////////////////////////////////////////////////////////////////////

var allowed_user = 'inscriptions.sca@gmail.com'
var email_loisir = 'sca.loisir@gmail.com'
var email_comp = 'skicluballevardin@gmail.com'
var email_dev = 'apbianco@gmail.com'
var email_license = 'licence.sca@gmail.com'

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

// Helper function with the core phone number formatting logic
function formatPhoneNumberString(phoneString) {
  if (phoneString == null || phoneString == '') {
    return '';
  }
  // Ensure it's a string before calling string methods
  var phone = String(phoneString);

  // Compress the phone number removing all spaces
  phone = phone.replace(/\s/g, "");
  // Compress the phone number removing all hyphens
  phone = phone.replace(/-/g, "");
  // Insert replacing groups of two digits by the digits with a space
  var regex = new RegExp('([0-9]{2})', 'gi');
  phone = phone.replace(regex, '$1 ');
  // Remove trailing space if any
  return phone.replace(/\s$/, "");
}

// Reformat the phone numbers
function formatPhoneNumbers() {
  function formatPhoneNumber(coords) {
    var phone = getStringAt(coords);
    // Only format if not empty, and let formatPhoneNumberString handle actual empty string logic
    if (phone != '') {
      var formattedPhone = formatPhoneNumberString(phone);
      setStringAt(coords, formattedPhone, "black");
    }
    // If phone is initially empty, do nothing, consistent with original behavior.
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

// Return a plural based on the value of number:
// "a" -> "as", "a b" -> "as bs", " a b" -> " as bs", " a b " -> " as bs "
function Plural(number, message) {
  if (number <= 1) {
    return message
  }
  if (!/\S/.test(message)) {
    return message
  }  
  var prefix = '', postfix = ''
  if (message[0] == ' ') {
    prefix = ' '
    message = message.substring(1)
  }
  if (message[message.length-1] == ' ') {
    postfix = ' '
    message = message.substring(0, message.length-1)
  }
  return prefix + message.split(" ").join("s ") + "s" + postfix
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
// Auto-fill AKA Magic Wand.
///////////////////////////////////////////////////////////////////////////////

function computeLicense(license_map, dob, level) {
  var is_competitor = isLevelComp(level)
  for (var license in license_map) {
    if (isLicenseNoLicense(license)) {
      continue
    }
    if (license_map[license].ValidateDoB(dob)) {
      if (is_competitor) {
        if (license == getCompJuniorLicenseString() ||
            license == getCompAdultLicenseString()) {
          return license
        }
      } else {
        return license
      }
    }
  }
  return getNoLicenseString()
}

function autoComputeLicenses() {
  // Return true when a license is elligible to a familly license upgrade: it 
  // must be a non comp kid/adult license.
  function EligibleToFamillyLicenseUpgrade(license) {
    return (license == getNonCompJuniorLicenseString() ||
            license == getNonCompAdultLicenseString())
  }
  updateStatusBar("Attribution automatique des licenses...", "grey")
  var license_map = createLicensesMap(SpreadsheetApp.getActiveSheet())
  // We're going to store what will be computed in these two array: one
  // to remember the row we changed and an other one to remember the
  // value we want to write back.
  var computed_licenses_rows = []
  var computed_licenses = []
  // We also will as we go count adults and kids to establish whether
  // it's fit to automatically attribute a familly license
  var number_of_adults = 0
  var number_of_kids = 0
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index]
    // Only if the row is valid
    var first_name = getStringAt([row, coord_first_name_column])
    var last_name = getStringAt([row, coord_last_name_column])
    if (first_name == "" || last_name == "") {
      continue
    }
    // Retrieve the UNNORMALIZED proposed license. We don't second-guess
    // the user only if the license has been knowingly set to Exec license
    // or no license.
    var selected_license = getRawLicenseAt([row, coord_license_column])
    if (isLicenseExec(selected_license) || isLicenseNoLicense(selected_license)) {
      continue
    }
    var dob = getDoB([row, coord_dob_column])
    // Verify we have a DoB other issue an error
    if (dob == undefined) {
        displayErrorPanel(
          "Date de naissance non fournie ou date de naissance mal formatée (JJ/MM/AAAA)" +
          " ou année de naissance fantaisiste pour " + first_name + " " + last_name)
        return false     
    }
    var level = getStringAt([row, coord_level_column])
    if (isLevelNotAdjusted(level)) {
      displayErrorPanel("Niveau pour " + first_name + " " + last_name + " non adjusté")
      return false
    }
    var license = computeLicense(license_map, dob, level)
    if (isLicenseNoLicense(license)) {
      displayErrorPanel("Pas d'attribution de licence possible pour " + first_name + " " + last_name)
      return false;
    }
    // We have successfully automatically computed the license. 
    // - Remember the license we want to write back and where.
    // - If that's a non comp adjult or kid license, adjust the number
    //   of adults/kids that got this license. We will use this data later
    //   to decide whether we want to assign a familly license
    computed_licenses.push(license)
    computed_licenses_rows.push(row)
    if (EligibleToFamillyLicenseUpgrade(license)) {
      if (isAdult(dob)) {
        number_of_adults += 1
      } else {
        number_of_kids += 1
      }
    }
  }
  // Take a look at the licenses we have added so far and determine whether we want
  // to assign a familly license 
  var attribute_familly_license = number_of_adults >= 2 && number_of_kids >= 2
  // Go over the licenses we want to write back and write them back, inserting a
  // familly license where we can.
  for (var index in computed_licenses_rows) {
    var license = computed_licenses[index]
    if (attribute_familly_license && EligibleToFamillyLicenseUpgrade(license)) {
      license = getNonCompFamilyLicenseString()
    }
    setStringAt([computed_licenses_rows[index], coord_license_column], license)
  }
  return true
}

function autoFillLicensePurchases() {
  updateStatusBar("Achat automatique des licenses...", "grey", add=true)
  var license_map = createLicensesMap(SpreadsheetApp.getActiveSheet())
  // Collect the attributed licenses
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getLicenseAt([row, coord_license_column]);
    license_map[selected_license].IncrementAttributedLicenseCount()
  }
  for (var license in license_map) {
    license_map[license].SetPurchasedLicenseAmount()
  }
  return true
}

function autoFillNonCompSubscriptions() {
  updateStatusBar("Achat automatique des adhésions loisir...", "grey", add=true)
  var subscription_map = createNonCompSubscriptionMap(SpreadsheetApp.getActiveSheet())
  var subscription_slots = [0, 0, 0, 0, 0]
  var current_non_rider_slot = 1
  // Collect the licenses and the levels. We assume someone wants 
  // a subscription when they have a adult/kid non comp license and that
  // their level has a value that isn't "not in scope" (non concernée)
  // As usual, we start looping over the identity section. We don't do any verification
  // of names/DoB because we assume that the auto-filler has previously flagged any
  // problems
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getLicenseAt([row, coord_license_column])
    var level = getStringAt([row, coord_level_column])
    // Non competitor license and only if level indicates interest in a non competitor
    // subscription.
    if (isLicenseNonComp(selected_license) && isLevelDefined(level)) {
      // Riders are accumulated
      if (isLevelRider(level)) {
        subscription_slots[0] += 1
      } else {
        // Non riders are dispatched. We stop filling things past 4
        // FIXME: Issue a warning?
        if (current_non_rider_slot < 5) {
          subscription_slots[current_non_rider_slot] = 1
          current_non_rider_slot += 1
        }
      }
    }
  }
  // A rider subscription counts as an occupied non-rider subscription slot:
  // [1, 1, 0, 0, 0] becomes [1, 0, 1, 0, 0] and [2, 1, 1, 0, 0] becomes
  // [1, 0, 0, 1, 1]. This adjustment can not happen as we built 
  // subscrition_slots because a rider can happen at any time in the list of
  // registered familly members.
  // Save the number of rider, use it as an iteration counter
  var number_of_riders = subscription_slots.shift()
  var operation_count = number_of_riders
  // Insert as many zeros at the beginning of the array as there are riders
  while (operation_count != 0) {
    // Insert a zero
    subscription_slots.splice(0, 0, 0)
    operation_count -= 1
  }
  // Insert the number of riders back
  subscription_slots.splice(0, 0, number_of_riders)
  // Truncate the array by as many 0s we initially inserted
  subscription_slots.splice(-number_of_riders, number_of_riders)

  for (var index in noncomp_subscription_categories) {
    var subcription = noncomp_subscription_categories[index]
    subscription_map[subcription].SetPurchasedSubscriptionAmount(subscription_slots[index])
  }
  return true
}

function autoFillCompSubscriptions() {
  updateStatusBar("Achat automatique des adhésions compétition...", "grey", add=true)
  var subscription_map = createCompSubscriptionMap(SpreadsheetApp.getActiveSheet())
  var comp_categories = []
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var selected_license = getLicenseAt([row, coord_license_column])
    var level = getStringAt([row, coord_level_column])
    // Competitor license and only if level indicates interest in a non competitor
    // subscription.
    if (isLicenseComp(selected_license) && isLevelComp(level)) {
      var dob = getDoB([row, coord_dob_column])
      // First - find the level this competitor is at.
      // NOTE: assumption is made here that DoB has been minimally verified already
      // when autoComputeLicenses was previously invoked.
      for (var comp_category_index in comp_subscription_categories) {
        // Remember: categories in the map are labelled with an ranking index.
        // For identification, using a fixed ranking index of 1 is fine.
        var comp_category = 1 + comp_subscription_categories[comp_category_index]
        if (subscription_map[comp_category].ValidateDoB(dob)) {
          // But what we store is the non prefixed category.
          comp_categories.push(comp_subscription_categories[comp_category_index])
        }
      }
    }        
  }
  // Clear existing choices
  for (var category in subscription_map) {
    subscription_map[category].SetPurchasedSubscriptionAmount(0)
  }
  // And set the new values
  comp_categories = comp_categories.sort(categoriesAscendingOrder)
  var rank = 1
  for (var index in comp_categories) {
    // Can't use index+1 as rank.
    var category = rank + comp_categories[index]
    subscription_map[category].SetPurchasedSubscriptionAmount(1)
    rank += 1
    // If we go past our capacity, just stop
    if (rank > 4) {
      break
    }
  }
  return true
}

function autoFillSkiPassPurchases() {
  // Clear ski pass rebates as they aren't computed by the magic wand for now.
  clearSkiPassesRebates()
  updateStatusBar("Achat automatique des forfaits...", "grey", add=true)
  var ski_pass_map = createSkipassMap(SpreadsheetApp.getActiveSheet())
  // Collect the attributed licenses
  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index]
    // No name verification as this methods runs after this has been
    // done
    var dob = getDoB([row, coord_dob_column])

    if (dob == undefined) {
      continue
    }
    for (var ski_pass in ski_pass_map) {
      if (!isSkipPassLocalizedCollet(ski_pass)) {
        continue
      }
      ski_pass_map[ski_pass].IncrementAttributedSkiPassCountIfDoB(dob)
    }
  }
  for (var ski_pass in ski_pass_map) {
      ski_pass_map[ski_pass].SetPurchasedSkiPassAmount()
  }
  return true
}

///////////////////////////////////////////////////////////////////////////////
// Drop down value validation
///////////////////////////////////////////////////////////////////////////////

// Validate a cell at coord whose value is set via a drop-down
// menu. We rely on the fact that a cell not yet set a proper value
// can only take a couple values.  When the value is valid, it is
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
      displayErrorPanel(error)
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
      displayErrorPanel("Error during test")
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
    var level = getStringAt([coords_identity_rows[index], coord_level_column]);
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
    if (isLevelNotComp(level) && isLicenseComp(license)) {
      return (
        "Vous devez utiliser le niveau 'Compétiteur' pour " + 
        first_name + " " + last_name + " qui prend une license " + license +
        " ou choisir une autre license."         
      )      
    }
    if (isLevelDefined(level) && isLicenseExec(license)) {
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
              "license choisie. Une " + license + ' ' +
              license_map[license].ValidDoBRangeMessage() + '.')
    }
    if (isLicenseExec(license) && city == '') {
      return (first_name + " " + last_name + ": la licence attribuée (" +
              license + ") requiert de renseigner une ville et un pays de naissance");
    }
    if (isLicenseDefined(license) && !isLicenseExec(license) && city != '') {
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
    license_map[index].UpdatePurchasedLicenseAmountFromTrix()
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
        var alc = license_map[index].AttributedLicenseCount()
        var pla = license_map[index].PurchasedLicenseAmount()
      return (
        "Le nombre de " + Plural(alc, "licence") + " '" + index + "' "+
        Plural(alc, "attribuée") + " (au nombre de " + alc + ")\n" +
        "ne correspond pas au nombre de " +
        Plural(pla, "licence achetée") + " (au nombre de " + pla  + ")")
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
    comp_subscription_map[subscription].UpdatePurchasedSubscriptionAmountFromTrix()
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
    for (var rank = 1; rank <= comp_kids_per_family; rank += 1) {
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
  for (var rank = 1; rank <= comp_kids_per_family; rank += 1) {
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
    subscription_map[subscription].UpdatePurchasedSubscriptionAmountFromTrix()
  }

  // Compute the number of people with a level, making the distinction between riders/non riders.
  // These are the people that need a subscription. Note that a level indicating a competitor
  // is not taken into account.
  var rider_number = 0
  var non_rider_number = 0
  coords_identity_rows.forEach(function(row) {
    var level = getStringAt([row, coord_level_column])
    if (isLevelRider(level)) {
      rider_number += 1
    }
    if (isLevelNotComp(level) && isLevelNotRider(level)) {
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
  var first_kid = subscription_map[getFirstKidString()].PurchasedSubscriptionAmount()
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

function clearSkiPassesRebates() {
  clearRange(coord_rebate_family_of_4_count)
  clearRange(coord_rebate_family_of_4_amount)
  clearRange(coord_rebate_family_of_5_count)
  clearRange(coord_rebate_family_of_5_amount)
  SpreadsheetApp.flush();
}

function getTotalSkiPasses() {
  return getNumberAt(coord_total_ski_pass)
}

function validateSkiPasses() {
  updateStatusBar("Validation des forfaits loisir...", "grey", add=true)
  // Always clear the rebate section before eventually recomputing it at the end
  clearSkiPassesRebates()
  var ski_passes_map = createSkipassMap(SpreadsheetApp.getActiveSheet())

  for (var index in coords_identity_rows) {
    var row = coords_identity_rows[index];
    var dob = getDoB([row, coord_dob_column])
    // Undefined DoB marks a non existing entry that we skip
    if (dob == undefined) {
      continue
    }
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
    skipass.UpdatePurchasedSkiPassAmountFromTrix()
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
  constructor(error, civility, family_name, mail_to, legal_disclaimer, ffs_medical_form) {
    this.error = error
    this.civility = civility
    this.family_name = family_name
    this.mail_to = mail_to
    this.legal_disclaimer = legal_disclaimer
    this.ffs_medical_form = ffs_medical_form
  }
}

function validateEmailAddress(email_address, mandatory=true) {
  if (!mandatory && email_address == '') {
    return true
  }
  if (mandatory && (email_address == '@' || email_address == '')) {
    return false
  }
  if (! email_address.match('@')) {
    return false
  }
  return true
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
  if (!validateEmailAddress(mail_to)) {
    displayErrorPanel(
      "Vous n'avez pas saisi d'adresse email principale ou " +
      "vous avez oublié \n" +
      "de valider l'adresse email par [return] ou [enter]...")
    return validatationDataError()
  }

  var mail_cc = getStringAt(coord_cc)
  if (mail_cc != '' ) {
    if (!validateEmailAddress(mail_cc, mandatory=false)) {
     displayErrorPanel(
      "Addresse email secondaire incorrecte")
    return validatationDataError()     
    }
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
    updateStatusBar("Validation règlement/autorisation/questionaire...", "grey", add=true)
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

    // Validate the ffs/medical form
    var ffs_medical_form_validation = validateAndReturnDropDownValue(
      coord_ffs_medical_form,
      "Vous n'avez pas renseigné de réponse à la question Notice FFS / Questionaire Médical.")
    if (ffs_medical_form_validation == '') {
      return validatationDataError()
    }

    // Validate the invoice payment
    var invoice_payment_validation = validateAndReturnDropDownValue(
      coord_payment_validation_form,
      "Vous n'avez pas validé le règlement de la facture.")
    if (invoice_payment_validation == '') {
      return validatationDataError()
    }
    // Verify that what's set matches the tally
    var owed = getNumberAt(coord_owed)
    if (owed < 0) {
      displayErrorPanel('Le montnant dû (' + owed + '€) ne peux pas être négatif')
      return validatationDataError()      
    }
    var total = getNumberAt(coord_total)
    switch(getStringAt(coord_payment_validation_form)) {
      case 'Acquitté':
        if (owed != 0) {
          displayErrorPanel('Payment marqué acquitté avec ' + owed + '€ restant à payer')
          return validatationDataError()
        }
        break
      case 'Non acquitté':
        if (owed == 0) {
          displayErrorPanel('Total dû de 0€, ce paiement devrait être marqué acquitté.')
          return validatationDataError()          
        }
        break
      case 'Accompte versé':
        if (owed == 0) {
          displayErrorPanel('Total dû de 0€, ce paiement devrait être marqué acquitté.')
          return validatationDataError()             
        }
        var down_payment = total - owed
        if (down_payment <= 0) {
          displayErrorPanel('Accompte de ' + down_payment + '€ versé pour un paiement dû de ' + owed + '€')
          return validatationDataError()          
        }
        break
      case 'Autre':
        break
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
                                   legal_disclaimer_validation, ffs_medical_form_validation)    
}

///////////////////////////////////////////////////////////////////////////////
// Invoice emailing
///////////////////////////////////////////////////////////////////////////////

// List possible actions: creating the invoice, emailing it, email it along
// content or just trigger a license creation request email.
var invoiceActions = {
  JUST_GENERATE_INVOICE: 1 << 0,
  LICENSE_REQUEST: 1 << 1,
  EMAIL_FOLDER: 1 << 2,
};

// Verify that a license can be asked - this just requires some payment
// verification
function askingLicenseOK() {
    switch(getStringAt(coord_payment_validation_form)) {
      case 'Acquitté':
        return true
        break
      case 'Non acquitté':
        return false
        break
      case 'Accompte versé':
        return false
        break
      case 'Autre':
        return true
        break      
    }
}

// When just_test is true, just see if a license request should be sent, but
// don't send any email.
// FIXME: Split in two really.
function maybeEmailLicenseSCA(invoice, just_test, ignore_payment) {
  if (! ignore_payment && just_test == false && ! askingLicenseOK()) {
    updateStatusBar("⚠️ PAS de demande de license (voir paiement)", "orange", add=true)      
    return false
  }
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
    if (license_number != '' && !license_number.match("^[0-9]+$")) {
      license_number = '<font color="red"><b>' + license_number + '</b></font>'
    } else if (license_number == '') {
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
  // If we haven't gathered any family members, there's no license request to
  // send and we return
  if (string_family_members == "") {
    return false
  }
  if (just_test) {
    return true
  }
  string_family_members = (
    "<p> " + license_count + Plural(license_count, " licence nécessaire") +
    " pour:</p><blockquote>\n" + string_family_members + "</blockquote>\n");
  
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
  return true
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

function generatePDFAndMaybeSendEmail(config) {
  updateStatusBar("⏳ Validation de la facture...", "orange")      
  var validation = validateInvoice();
  if (validation.error) {
    updateStatusBar("❌ La validation de la facture a échouée", "red")      
    return;
  }
  // Generate and prepare attaching the PDF to the email
  updateStatusBar("⏳ Préparation de la facture...", "orange")
  var pdf_file = generatePDF();
  var pdf = DriveApp.getFileById(pdf_file.getId());
  var attachments = [pdf.getAs(MimeType.PDF)]

  var just_generate_invoice = config & invoiceActions.JUST_GENERATE_INVOICE
  var email_folder = config & invoiceActions.EMAIL_FOLDER
  var license_request = config & invoiceActions.LICENSE_REQUEST
  
  if (email_folder) {
    updateStatusBar("⏳ Génération et envoit du dossier...", "orange")
  }
  else if (license_request) {
    updateStatusBar("⏳ Envoit de la demande de license", "orange")    
  }
  else if (just_generate_invoice) {
    updateStatusBar("⏳ Génération de la facture sans envoi...", "orange")
  }
  else {
    updateStatusBar("❌ Instruction non traitée", "red")      
    return;
  }

  var civility = validation.civility
  var family_name = validation.family_name
  var mail_to = validation.mail_to
  var legal_disclaimer_validation = validation.legal_disclaimer
  var legal_disclaimer_text = ''
  var ffs_medical_form_validation = validation.ffs_medical_form
  var ffs_medical_form_text = ''
  
  function insertNotesAndRules() {
    // Insert the note and rules for the parents anyways
    attachments.push(DriveApp.getFileById(rules_pdf).getAs(MimeType.PDF))
    attachments.push(DriveApp.getFileById(parents_note_pdf).getAs(MimeType.PDF));
    legal_disclaimer_text += (
      "<p>Vous trouverez également en attachement une note adressée aux " +
      "parents, ainsi que le règlement intérieur. Merci de lire ces deux " +
      "documents attentivement.</p>");
  }

  // Determine whether parental consent needs to be generated. If
  // that's the case, we generate additional attachment content.
  if (email_folder) {
    switch (legal_disclaimer_validation) {
      case 'À fournire signée':
        attachments.push(
          DriveApp.getFileById(legal_disclaimer_pdf).getAs(MimeType.PDF))
        legal_disclaimer_text = (
          "<p>Il vous faut compléter, signer et nous retourner la mention " +
          "légale fournie en attachment, couvrant le droit à l'image, le " +
          "règlement intérieur, les interventions médicales et la RGPD.</p>")
        insertNotesAndRules()
        break
      case 'Fournie et signée':
        insertNotesAndRules()
        break
      case 'Non nécessaire':
        break
    }

    switch (ffs_medical_form_validation) {
      case 'Signée, pas de certificat médical à fournir':
        ffs_medical_form_text = ('<p><b><font color="red">' +
                                 'Les réponses négative que vous avez porté au questionaire médical vous ' +
                                 'dispense de fournir un certificat médical.</font></b>' +
                                 '<p>Vous retrouverez aussi en attachement la notice FFS que vous avec déjà ' +
                                 'remplie afin de la conserver.')
        attachments.push(DriveApp.getFileById(ffs_information_leaflet_pdf).getAs(MimeType.PDF))
        break

      case 'Signée, certificat médical à fournir':
        ffs_medical_form_text = ('<p><b><font color="red">' +
                                 'La ou les réponses positives que vous avez porté au questionaire médical vous ' +
                                 'obligent à transmettre au SCA (inscriptions.sca@gmail.com) dans les plus ' +
                                 'brefs délais <u>un certificat médical en cours de validité</u></font></b>.' +
                                 '<p>Vous retrouverez aussi en attachement la notice FFS que vous avec déjà ' +
                                 'remplie afin de la conserver.')                                 
        attachments.push(DriveApp.getFileById(ffs_information_leaflet_pdf).getAs(MimeType.PDF))
        break

      case 'À fournir signée, questionaire médical à évaluer':
        ffs_medical_form_text = ('<p><b><font color="red">' +
                                 'Vous devez évaluer le <b>Questionnaire Santé Sportif MINEUR - ' + season + '</b> ou <b>' +
                                 'le Questionnaire Santé Sportif MAJEUR - ' + season + '</b> fournis en attachement et ' +
                                 'si une des réponses aux questions est OUI, vous devez transmettre au SCA ' +
                                 '(inscriptions.sca@gmail.com) dans les plus brefs délais <u>un certificat médical en cours ' +
                                 'de validité</u>. Il faut également <u>signer ' + ffs_information_leaflet_pages_to_sign +
                                 ' de la notice d\'informations FFS ' + season + '</u> fournie en attachement.' +   
                                 '</font></b>')
        attachments.push(DriveApp.getFileById(ffs_information_leaflet_pdf).getAs(MimeType.PDF))
        attachments.push(DriveApp.getFileById(autocertification_non_adult).getAs(MimeType.PDF))
        attachments.push(DriveApp.getFileById(autocertification_adult).getAs(MimeType.PDF))
        break

      case 'À fournir signée':
        ffs_medical_form_text = ('<p><b><font color="red">' +
                                 'Il faut <u>signer ' + ffs_information_leaflet_pages_to_sign +
                                 ' de la notice d\'informations FFS ' + season + '</u> fournie en attachement.' +   
                                 '</font></b>')
        attachments.push(DriveApp.getFileById(ffs_information_leaflet_pdf).getAs(MimeType.PDF))
      
      case 'Non nécessaire':
        break
    }
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
      ffs_medical_form_text +

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

  var email_quota_threshold = 1 + (cc_to == "" || cc_to == undefined ? 0 : 1)
  if (maybeEmailLicenseSCA([attachments[0]], just_test=true)) {
    email_quota_threshold += 1
  }
  // The final status to display is captured in this variable and
  // the status bar is updated after the aggregation trix has been
  // updated.
  var final_status = ['', 'green']
  if (email_folder) {
    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    if (emailQuotaRemaining < email_quota_threshold) {
      displayWarningPanel("Quota email insuffisant. Envoi " +
                          (just_generate_invoice ? "de la facture" : "du dossier") +
                          " retardé. Quota restant: " + emailQuotaRemaining +
                          ". Nécessaire: " + email_quota_threshold);
      final_status[0] = "⚠️ Quota email insuffisant. Envoi retardé"
      final_status[1] = "orange";
      // Insert a reference to the file in a trix
      var link = '=HYPERLINK("' + SpreadsheetApp.getActive().getUrl() +
                             '"; "' + SpreadsheetApp.getActive().getName() + '")'
      var context = (just_generate_invoice ? 'Facture seule à renvoyer' : 'Dossier complet à renvoyer')
      updateProblematicRegistration(link, context)
    } else {
      // Send the email  
      MailApp.sendEmail(email_options)
      final_status[0] = "✅ Dossier envoyé"
      // When the license wasn't sent, warn and add to the status bar so that
      // it is visible
      if (! maybeEmailLicenseSCA([attachments[0]], just_test=false)) {
        final_status[0] += ".\n⚠️ PAS de demande de license (voir paiement)"
        final_status[1] = "orange";
      } else {
        final_status[0] += ", demande de license faite"
      }
    }
  } else if (license_request) {
    maybeEmailLicenseSCA([attachments[0]], just_test=false, ignore_payment=true);
    final_status[0] = "✅ Demande de license envoyée"
  } else if (just_generate_invoice) {
      final_status[0] = "✅ Facture générée"
  } else {
    updateStatusBar("❌ Traitement impossible", "red")      
    return
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

// This is what the [generate invoice] button runs.
function JustGeneratePDFButton() {
  generatePDFAndMaybeSendEmail(invoiceActions.JUST_GENERATE_INVOICE)
}

// This is what the [generate license creation request] button runs.
function JustGenerateLicenseRequestButton() {
  generatePDFAndMaybeSendEmail(invoiceActions.LICENSE_REQUEST)
}

// This is what the [generate and send folder] button runs.
function GenerateFolderAndSendEmailButton() {
  generatePDFAndMaybeSendEmail(invoiceActions.EMAIL_FOLDER)
}

// This is what the [signal problem] button runs
function SignalProblem() {
  var link = '=HYPERLINK("' + SpreadsheetApp.getActive().getUrl() +
                         '"; "' + SpreadsheetApp.getActive().getName() + '")'
  updateProblematicRegistration(link, getStringAt(coord_status))
  updateStatusBar('⚠️ Facture signalée commme problématique', 'red')
}

// This is what the [magic wand] button runs
function magicWand() {
  updateStatusBar("")
  if (!displayYesNoPanel("Le remplissage automatique va replacer certains choix que vous " +
                         "avez déjà fait (attribution automatique des licenses, achats des " +
                         "licenses, adhésion loisir, etc...) Vous pourrez toujours modifier " +
                         "ces choix et une validation de la facture sera toujours effectuée \n\n" +
                         "Cliquez OK pour continuer, cliquez Annuler pour ne pas utiliser la " +
                         "fonctionalité remplissage automatique")) {
    return
  }
  if (autoComputeLicenses()) {
    if (autoFillLicensePurchases()) {
      if (autoFillNonCompSubscriptions()) {
        if (autoFillCompSubscriptions()) {
          if (autoFillSkiPassPurchases()) {
            var status_bar_text = "✅ Remplissage automatique terminé..."
            if (getTotalSkiPasses() >= 4) {
              status_bar_text += "\n⚠️ Réduction famille? Validez la facture"
            }
            updateStatusBar(status_bar_text, "green")
            return
          }
        }
      }
    }
  }
  updateStatusBar("❌ Le remplissage automatique a échoué", "red") 
}
