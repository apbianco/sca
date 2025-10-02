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
dev_or_prod = "dev"

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
// FIXME: Move this somewhere else
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
  return license != getNoLicenseString() && license in getMemoizedLicensesMap()
}

// If license is not a valid license that's not getNoLicenseString(),
// this will return false and that's fine.
function isLicenseNoLicense(license) {
  return license == getNoLicenseString()
}

// If license is something that's not a known license, this will return
// true and that's fine.
function isLicenseNotDefined(license) {
  return !isLicenseDefined(license)
}

function isLicenseNonCompJunior(license) {
  return license == getNonCompJuniorLicenseString()
}

function isLicenseNonCompAdult(license) {
  return license == getNonCompAdultLicenseString()
}

function isLicenseNonCompFamily(license) {
  return license == getNonCompFamilyLicenseString()
}

// This excludes Executive license.
function isLicenseNonComp(license) {
  return (isLicenseNonCompJunior(license) ||
          isLicenseNonCompFamily(license) ||
          isLicenseNonCompAdult(license))
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
  IsA(key) { return key == this.name }

  // When this method run, we capture the amount licenses of that given type
  // the operator entered.
  UpdatePurchasedLicenseAmountFromTrix() {
    if (this.purchase_range != null) {
      this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                           this.purchase_range.getColumn()])
    }
  }
  LicenseAmount() {
    if (this.purchase_range == null) {
      return 0
    }
    return getNumberAt([this.purchase_range.getRow(),
                        this.purchase_range.getColumn()-1])
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
    if (count > 0 && isLicenseNonCompFamily(this.Name())) {
      count = 1
    }
    if (count > 0) {
      this.purchased_amount = count
      setStringAt(coord, count)
    } else {
      setStringAt(coord, "")
    }
  }
  PurchasedLicenseAmount() {
    return this.purchase_range == null ? 0 : this.purchased_amount
  }
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
      (dob) => {return true},
      "N/A"),
    'CN Jeune (Loisir)': new License(
      getNonCompJuniorLicenseString(),
      getRange(getNonCompJuniorLicenseString()),
      // Remember: no local variable capture possible in functor, use function calls only
      (dob) => {return ageVerificationBornAfterDateIncluded(dob, createDate(getNonCompJuniorLicenseString(), "January 1"))},
      "requiert d'être né en " + getYear(getNonCompJuniorLicenseString()) + " et après"),
    'CN Adulte (Loisir)': new License(
      getNonCompAdultLicenseString(),
      getRange(getNonCompAdultLicenseString()),
      (dob) => {return ageVerificationBornBeforeDateIncluded(dob, createDate(getNonCompAdultLicenseString(), "December 31"))},
      "requiert d'être né en " + getYear(getNonCompAdultLicenseString()) + " et avant",
      (key) => {return isLicenseNonCompAdult(key)}),
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

// Only for const use of the license map.
const getMemoizedLicensesMap = (function() {
  let cached_license_map = null
  return function() {
    if (cached_license_map === null) {
      cached_license_map = createLicensesMap(SpreadsheetApp.getActiveSheet())
    }
    return cached_license_map
  }
})();

///////////////////////////////////////////////////////////////////////////////
// Competitor categories management
///////////////////////////////////////////////////////////////////////////////

// Competitors: categories definition (used to establish the subscription
// range). Also define an array with these values as it's needed everywhere.
// Also define a categories sorting function, used to sort arrays of categories,
// verifying that comp_subscription_categories.sort() == comp_subscription_categories
function getU6String() { return 'U6' }
function getU8String() { return 'U8' }
function getU10String() { return 'U10'}
function getU12PlusString() { return 'U12+' }
var comp_subscription_categories = [
  getU6String(), getU8String(), getU10String(), getU12PlusString()
]

function categoriesAscendingOrder(cat1, cat2) {
  var v1 = cat1.match(/\d+/)
  var v2 = cat2.match(/\d+/)
  if (! v1 || ! v2) {
    displayErrorPanel("categoriesAscendingOrder(" + cat1, ", " + cat2 + ")")
  }
  v1 = parseInt(v1[0])
  v2 = parseInt(v2[0])
  if (v1 == v2) {
    return 0
  }
  if (v1 > v2) {
    return 1
  }
  return -1
}

///////////////////////////////////////////////////////////////////////////////
// Non competitor levels and subscription values management
///////////////////////////////////////////////////////////////////////////////

// Note:
//  - Absence of level is '' - no selection was made when one was expected.
//  - No LevelString hints at the existence of a level that hasn't yet been defined.  
//  - Only License is a level (so it defines one) that indicates that the skier
//    doesn't seek receiving ski instructions.
function getNoLevelString() { return 'Non déterminé' } // DOES define a level
function getOnlyLicense() { return 'Licence seule'}    // DOES define a level
function getLevelCompString() {return "Compétiteur" }

// Subscription categories
function getAdultString() { return 'Adulte' }
function getRiderLevelString() { return 'Rider' }
function getFirstKidString() { return '1er enfant' }
function getSecondKidString() { return '2ème enfant' }
function getThirdKidString() { return '3ème enfant' }
function getFourthKidString() { return '4ème enfant' }

var noncomp_subscription_categories = [
  getAdultString(),
  getRiderLevelString(),
  getFirstKidString(),
  getSecondKidString(),
  getThirdKidString(),
  getFourthKidString()
]

// A level is not adjusted when it starts with "⚠️ "
function isLevelNotAdjusted(level) {
  return level.substring(0, 3) == "⚠️ ";
}

// NOTE: A level is not defined when it has not been entered.
// NOTABLY:
//  - A level of getNoLevelString() value *DEFINES* a level.
//  - A level not yet adjusted *DEFINES* a level.
function isLevelNotDefined(level) {
  return level == ''
}

function isLevelLicenseOnly(level) {
  return level == getOnlyLicense()
}

function isLevelDefined(level) {
  return level != ''
}

function isLevelComp(level) {
  return level == getLevelCompString(level)
}

function isLevelNotComp(level) {
  return (isLevelLicenseOnly(level) ?
          false : (! isLevelDefined(level) ?
                   false : ! isLevelComp(level)))
}

function isLevelRider(level) {
  return level == getRiderLevelString()
}

function isLevelRecreationalNonRider(level) {
  return (isLevelLicenseOnly(level) ?
          false : (! isLevelDefined(level) ?
                   false : (isLevelRider(level) ?
                            false : isLevelNotComp(level))))
}

function isSubscriptionAdult(subscription) {
  return isLevelDefined(subscription) && subscription == getAdultString()
}

///////////////////////////////////////////////////////////////////////////////
// Subscription management code
///////////////////////////////////////////////////////////////////////////////

// A Subscription class to create an object that has a name, a range in the trix at
// which it can be marked as purchased, a validation method that takes a DoB
// as input and can keep track of occurences and purchases.
class Subscription {
  constructor(name, hr_name, purchase_range, dob_validation_method) {
    // The name of the subscription - it's also a key so not always
    // human readable
    this.name = name
    // Human readable name
    this.hr_name = hr_name
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
  HumanReadableName() { return this.hr_name }
  IsA(key) { return this.name == key}

  // When this method run, we capture the amount of subscription of that nature
  // the operator entered.
  UpdatePurchasedSubscriptionAmountFromTrix() {
    this.purchased_amount = getNumberAt([this.purchase_range.getRow(),
                                         this.purchase_range.getColumn()])
  }
  PurchasedSubscriptionAmount() { return this.purchased_amount }
  SubscriptionAmount() {
    return getNumberAt([this.purchase_range.getRow(),
                        this.purchase_range.getColumn()-1])
  }  
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

// Definition of all possible subscription values. This map can not be memoized
// as not const instances of it are created (map elements are always modified)
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
  function getRange(label, adjusted_rank) {
    // Use adjusted_rank to produce [54, 5] and then [58, 5], etc...
    return sheet.getRange(4*adjusted_rank + comp_subscription_map[label][2],
                          comp_subscription_map[label][3])
  }
  var to_return = {}
  for (var rank = 1; rank <= comp_kids_per_family; rank +=1) {
    var label = getU6String()
    var ranked_label = rank + label
    to_return[ranked_label] = new Subscription(
      ranked_label,
      nthString(rank) + " " + label,
      getRange(label, rank-1),
      // Remember: no local variable capture possible in functor, use function calls only
      (dob) => {return ageVerificationBornBetweenYearsIncluded(dob, getFirstYear(getU6String()), getLastYear(getU6String()))})
    label = getU8String()
    ranked_label = rank + label
    to_return[ranked_label] = new Subscription(
      ranked_label,
      nthString(rank) + " " + label,
      getRange(label, rank-1),
      (dob) => {return ageVerificationBornBetweenYearsIncluded(dob, getFirstYear(getU8String()), getLastYear(getU8String()))})
    label = getU10String()
    ranked_label = rank + label
    to_return[ranked_label] = new Subscription(
      ranked_label,
      nthString(rank) + " " + label,
      getRange(label, rank-1),
      (dob) => {return ageVerificationBornBetweenYearsIncluded(dob, getFirstYear(getU10String()), getLastYear(getU10String()))})
    label = getU12PlusString()
    ranked_label = rank + label
    to_return[ranked_label] = new Subscription(
      ranked_label,
      nthString(rank) + " " + label,
      getRange(label, rank-1),
      (dob) => {return ageVerificationBornBeforeYearIncluded(dob, getFirstYear(getU12PlusString()))})
  }
  validateClassInstancesMap(to_return, 'createCompSubscriptionMap')
  return to_return
}

function createNonCompSubscriptionMap(sheet) {
  var to_return = {}
  var row = coord_noncomp_start_row
  for (index in noncomp_subscription_categories) {
    var label = noncomp_subscription_categories[index]
    to_return[label] = new Subscription(
      label,
      label,
      sheet.getRange(row, coord_noncomp_column),
      (dob) => {return true}),
    row += 1
  }
  validateClassInstancesMap(to_return, 'createNonCompSubscriptionMap')
  return to_return
}

function UpdateBasicSubscriptionNumber(basic_subscriptions_number) {
  setStringAt(basic_subscription_coord, basic_subscriptions_number <= 0 ? '' : basic_subscriptions_number)
}

function GetBasicSubscriptionNumber() {
  return getNumberAt(basic_subscription_coord)
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

  SkiPassAmount() {
    if (this.purchase_range == null) {
      return 0
    }
    return getNumberAt([this.purchase_range.getRow(),
                        this.purchase_range.getColumn()-1])
  } 

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
      // Note: the use of getEarlyDate (January 1st, <YEAR>) is sub-obtimal here. We should be using
      // December 31st, <YEAR>. But getEarlyDate is a clean way to fetch the content. Oh well.
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

function arrayToRawString(v) {
  var to_return = ""
  for (const entry of v) {
    to_return += " " + entry
  }
  return to_return
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

function rangeToString(range) {
  return 'row=' + range.getRow() + ", col=" + range.getColumn()
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
  updateStatusBar("Attribution automatique des licences...", "grey")
  // FIXME: Eligible to using getMemoizedLicensesMap.
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
    // For an exec license, we adjust the level accordingly.
    if (isLicenseExec(selected_license)) {
      setStringAt([row, coord_level_column], getOnlyLicense())
      continue
    }
    // For a competitor license, we adjust the level accordingly
    if (isLicenseCompAdult(selected_license) || isLicenseCompJunior(selected_license)) {
      setStringAt([row, coord_level_column], getLevelCompString())
      continue      
    }
    if (isLicenseNoLicense(selected_license)) {
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
  updateStatusBar("Achat automatique des licences...", "grey", add=true)
  // Not eligible to use getMemoizedLicensesMap.
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
  //       current_non_rider_slot
  //           rider_index       \
  //                       \      \
  //                        V      V
  //                        Rider, 1st Kid, 2nd Kid, 3rd Kid, 4th Kid
  var subscription_slots = [0,     0,       0,       0,       0]
  var rider_index = 0
  var current_non_rider_slot = 1
  var number_of_adults = 0
  var basic_subscriptions_number = 0
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
    // Level not defined or level is competitor or license is competition or license is exec: we skip
    if (!isLevelDefined(level) || isLevelComp(level) ||
         isLicenseComp(selected_license) || isLicenseExec(selected_license)) {
      continue
    }
    // Handle non competitors license with level set to license only, which indicates
    // there's no interest in being under the supervision of an instructure. This means
    // that a basic subscription has to be registred. Executive licenses are exclused.
    if (isLevelLicenseOnly(level)) {
      basic_subscriptions_number += 1
      continue
    }
    // Handle non competitor license with a level defined, indicating interest
    // in being under the supervision of an instructor, which includes adults.
    // Riders are accumulated
    if (isLevelRider(level)) {
      subscription_slots[rider_index] += 1
      continue
    } else {
      // If we have an adult by DOB, we fill in the adult section. Again, an
      // adult getting a executive license is excluded.
      if (isAdult(getDoB([row, coord_dob_column]))) {
        number_of_adults += 1
      } else {
        // Non riders are dispatched. We stop filling things past 4
        // FIXME: Issue a warning?
        if (current_non_rider_slot < 5) {
          subscription_slots[current_non_rider_slot] = 1
          current_non_rider_slot += 1
        }
      }
      continue
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
  // Insert the number of adults 
  subscription_slots.splice(0, 0, number_of_adults)

  for (var index in noncomp_subscription_categories) {
    var subscription = noncomp_subscription_categories[index]
    subscription_map[subscription].SetPurchasedSubscriptionAmount(subscription_slots[index])
  }
  UpdateBasicSubscriptionNumber(basic_subscriptions_number)
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
  constructor(first_name, last_name, dob, sex, city,
              license_type, license_number, level,
              parent1_email, parent1_phone, parent2_email, parent2_phone, parent_city) {
    this.first_name = first_name
    this.last_name = last_name
    this.dob = dob
    this.sex = sex
    this.city = city
    this.license_type = license_type
    this.license_number = license_number
    this.level = level
    this.parent1_email = parent1_email
    this.parent1_phone = parent1_phone
    this.parent2_email = parent2_email
    this.parent2_phone = parent2_phone
    if (city == "\\") {
      this.parent_city = parent_city
    } else {
      this.parent_city = city
    }
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

    var parent1_email = getStringAt(coord_family_email)
    var parent2_email = getStringAt(coord_cc)
    var parent1_phone = getStringAt(coord_family_phone1)
    var parent2_phone = getStringAt(coord_family_phone2)
    var parent_city = getStringAt(coord_family_city)

    family.push(new FamilyMember(first_name, last_name, birth,
                                 sex, city, license, license_number, level,
                                 parent1_email, parent1_phone,
                                 parent2_email, parent2_phone,
                                 parent_city))
  }
  return family
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

function TESTValidateInvoice() {
  function test(f) {
    var result = f()
    if (result == {}) {
      displayErrorPanel("Error during test")
    }
  }
  test(validateInvoice)
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
  
  updateStatusBar("Validation des coordonnées...", "grey")
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
  var city = getStringAt(coord_family_city).toUpperCase()
  if (street_address == "" || zip_number == "" || city == "") {
    displayErrorPanel(
      "Vous n'avez pas correctement renseigné une adresse de facturation:\n" +
      "numéro de rue, code postale ou commune - ou " +
      "vous avez oublié \n" +
      "de valider une valeur entrée par [return] ou [enter]...")
    return validatationDataError()
  }
  // Write city back using all caps
  setStringAt(coord_family_city, city)

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

    // Perform the validation of the level that says license only
    validation_error = validateOnlyLicensesLevel()
    if (validation_error) {
        displayErrorPanel(validation_error);
        return validatationDataError()      
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
      displayErrorPanel('Le montant dû (' + owed + '€) ne peux pas être négatif')
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
    updateStatusBar("⚠️ PAS de demande de licence (voir paiement)", "orange", add=true)      
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
      "Numéro Licence: " + license_number + "<br>" +
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
    updateStatusBar("⏳ Envoit de la demande de licence", "orange")    
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
      "marlene.czajka@gmail.com (06-60-69-75-39) / Anne-So Marchand: " +
      "annesophie.marchand6857@gmail.com (06-26-26-27-97) pour le ski loisir ou " +
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
        final_status[0] += ".\n⚠️ PAS de demande de licence (voir paiement)"
        final_status[1] = "orange";
      } else {
        final_status[0] += ", demande de licence faite"
      }
    }
  } else if (license_request) {
    maybeEmailLicenseSCA([attachments[0]], just_test=false, ignore_payment=true);
    final_status[0] = "✅ Demande de licence envoyée"
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
  // Now we can update the accounting trix as well
  updateStatusBar("⏳ Enregistrement des données comptables...", "orange")
  updateAccountingTrix()
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
                         "avez déjà fait (attribution automatique des licences, achats des " +
                         "licences, adhésion loisir, etc...) Vous pourrez toujours modifier " +
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
