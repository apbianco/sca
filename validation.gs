///////////////////////////////////////////////////////////////////////////////
//
// Version: 2025-10-09T10:48 - Competitors registration.
//
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
// validateOnlyLicensesLevel     | Validate the level that says that only a    | NO
//                               | license is required                         | string if error
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
// - validateOnlyLicensesLevel(). | if error, bail
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
  test(validateOnlyLicensesLevel)
  test(validateCompSubscriptions)
  test(validateNonCompSubscriptions)
  test(validateSkiPasses)
}

// Verify that family members are properly defined, especially with regard to
// what they are claiming to be associated to (level, type of license, etc...)
// When this succeeds, the family members should be defined in such a way that
// other validations are easier to implement.
function validateFamilyMembers() {
  updateStatusBar("Validation des données de la famille...", "grey", add=true)
  // Get a licene map so that we can immediately verify the age picked
  // for a license is correct.
  // FIXME: eligible to use getMemoizedLicensesMap.
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
    // 1- When a level is defined, it must have been adjusted
    if (isLevelDefined(level) && isLevelNotAdjusted(level)) {
      return ("Ajuster le niveau de " + first_name + " " + last_name + " en précisant " +
              "le niveau de pratique pour la saison " + season)
    }

    // 2- When a level is defined a license must be defined.
    if (isLevelDefined(level) && isLicenseNotDefined(license)) {
      return ("Le niveau de pratique '" + level + "' est défini pour " +
               first_name + " " + last_name +
              " ce qui traduit l'intention de prendre une adhésion au club. Choisissez " +
              "une licence appropriée pour cette personne")
    }
    // 3- A level must be defined for non competitor
    if (isLevelNotDefined(level) && isLicenseNonComp(license)) {
      return (
        "Vous devez fournir un niveau de pratique ou choisir le niveau " +
        "'Licence seule' pour le  non compétiteur " + 
        first_name + " " + last_name + " qui prend une licence " + license + ""
      )
    }
    // 4- A competitor can not declare a level, it will confuse the rest of the validation
    if (isLevelNotComp(level) && isLicenseComp(license)) {
      return (
        "Vous devez utiliser le niveau 'Compétiteur' pour " + 
        first_name + " " + last_name + " qui prend une licence " + license +
        " ou choisir une autre licence ou un autre niveau de pratique"         
      )      
    }

    // License validation:
    //
    // 0- A license must be in the list of possible licenses
    // 1- A license must match the age of the person it's attributed to
    // 2- An exec license requires a city of birth an a License only level
    // 3- An existing non exec license doesn't require a city of birth
    if (!license_map.hasOwnProperty(license)) {
      return (first_name + " " + last_name + "La licence attribuée '" + license +
              "' n'est pas une licence existante !")
    }   
    // When a license for a family member, it must match the dob requirement
    if (isLicenseDefined(license) && ! license_map[license].ValidateDoB(dob)) {
      return (first_name + " " + last_name +
              ": l'année de naissance " + getDoBYear(dob) +
              " ne correspond critère de validité de la " +
              "licence choisie. Une " + license + ' ' +
              license_map[license].ValidDoBRangeMessage() + '.')
    }
    if (isLicenseExec(license)) {
      if (city == '') {
        return (first_name + " " + last_name + ": la licence attribuée [" +
                license + "] requiert de renseigner une ville et un pays de naissance");
      }
      if (!isLevelLicenseOnly(level)) {
        return (first_name + " " + last_name + ": la licence attribuée [" +
                license + "] requiert de choisir le niveau '" + getOnlyLicense() + "'.");        
      }
    }
    if (isLicenseDefined(license) && !isLicenseExec(license) && city != '') {
      return (first_name + " " + last_name + ": la licence attribuée [" +
              license + "] ne requiert pas de renseigner une ville et un pays de naissance. " +
              "Supprimez cette donnée")
    }
  }
  return ''
}

// Cross check the attributed licenses with the ones selected for payment
function validateLicenses() {
  updateStatusBar("Validation du choix des licences...", "grey", add=true)
  // Not eligible to use getMemoizedLicensesMap.
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
    if (isLicenseNotDefined(index) || isLicenseNonCompFamily(index)) {
      continue
    }
    // What you attributed must match what you're purchasing...
    if (license_map[index].PurchasedLicenseAmount() != license_map[index].AttributedLicenseCount()) {
        var alc = license_map[index].AttributedLicenseCount()
        var pla = license_map[index].PurchasedLicenseAmount()
      return (
        "Le nombre de " + Plural(alc, "licence") + " '" + index + "' "+
        Plural(alc, "attribuée") + " [" + alc + "]\n" +
        "ne correspond pas au nombre de " +
        Plural(pla, "licence achetée") + " [" + pla  + "]")
    }
  }
  return ''
}

// Validate those who picked up just a license but aren't going ski under
// supervision
function validateOnlyLicensesLevel() {
  updateStatusBar("Validation licences seules...", "grey", add=true)
  var basic_subscription_number = 0
  coords_identity_rows.forEach(function(row) {
    var level = getStringAt([row, coord_level_column])
    var license = getStringAt([row, coord_license_column])
    if (isLevelLicenseOnly(level) && !isLicenseExec(license)) {
      basic_subscription_number += 1
    }
  })
  // The number of basic subscription collected must match the number of charges for that item
  var subscribed_basic_subscription_number = GetBasicSubscriptionNumber()
  if (basic_subscription_number != subscribed_basic_subscription_number) {
    return ("Le nombre de " + Plural(basic_subscription_number, "licence") + " sans enseignement de ski " +
            Plural(basic_subscription_number, "souscrite") + " [" + basic_subscription_number + "] " +
            "ne correspond pas au nombre de " + Plural(subscribed_basic_subscription_number, "licence") +
            " sans enseignement de ski " + Plural(subscribed_basic_subscription_number, "renseignée") +
	    " [" + subscribed_basic_subscription_number + "]")        
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
        return ("Le nombre d'" + Plural(current_purchased, "adhésion") + "'" + category + "' " +
	        Plural(current_purchased, "achetée") + " ['" + current_purchased + "'] n'est pas valide.")
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
      return ("[" + total_purchased + "] " + Plural(total_purchased, "adhésion") + " compétition '" +
      	      category + "'" + Plural(total_purchased, " achetée") + " pour [" + total_existing + "] " +
	      Plural(total_existing, "license") + " compétiteur dans cette tranche d\'âge")
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
  var non_rider_kid_number = 0
  var non_rider_adult_number = 0
  coords_identity_rows.forEach(function(row) {
    var level = getStringAt([row, coord_level_column])
    var license = getStringAt([row, coord_license_column])
    if (isLevelRider(level)) {
      rider_number += 1
    }
    if (isLevelRecreationalNonRider(level)) {
      // Here we need to make a difference between a kid and an adult:
      //   - License is NonCompAdult or Familly, person is an adult.
      var dob = getDoB([row, coord_dob_column])
      if (isLicenseNonCompAdult(license) || (isLicenseNonCompFamily(license) && isAdult(dob))) {
        non_rider_adult_number += 1
      } else {
        non_rider_kid_number += 1
      }
    }
  })

  // 1- The number of riders must be equal to the number of riders we found. First count them all
  //    and the perform the verification
  var subscribed_rider_number = subscription_map[getRiderLevelString()].PurchasedSubscriptionAmount()
  if (rider_number != subscribed_rider_number) {
    return ("Le nombre d'" + Plural(subscribed_rider_number, "adhésion") + " rider " +
            Plural(subscribed_rider_number, "souscrite") + " [" + subscribed_rider_number +
	    "] ne correspond pas au nombre de " +
	    Plural(rider_number, "rider renseigné") + " [" + rider_number + "]")
  }

  // 2- If we have N riders, the N first non Rider subscriptions can not be purchased,
  //    we jump directly to N+1
  var kids = [getFirstKidString(), getSecondKidString(), getThirdKidString(), getFourthKidString()]
  for (var index = 0; index < rider_number; index += 1) {
    var kid = subscription_map[kids[index]].PurchasedSubscriptionAmount()
    if (kid != 0) {
    return ("Une ou plusieurs adhésions Rider comptent comme des Adhésions / Stage / Transport - " +
            "Veuillez saisir les adhésions non Rider à partir du " + kids[rider_number])
    }
  }

  // 3- Start with the first non rider subscription that is allowed to exist and verify that
  //    if it exists, it is followed by an other subscription or no subscription. As soon as
  //    no subscription is found, no other subscrition can exist. This is a state machine
  //    with the following allowed transitions: ? -> {1, 0}, 1 -> {1, 0}, 0 -> {0}
  var state = -1
  var adjusted_noncomp_subscription_categories = noncomp_subscription_categories.slice(2+rider_number)
  for (var index in adjusted_noncomp_subscription_categories) {
    var subscription = adjusted_noncomp_subscription_categories[index]
    var current_purchased = subscription_map[subscription].PurchasedSubscriptionAmount()
    if (current_purchased < 0 || current_purchased > 1) {
      return ("La valeur du champ Adhésion / Stage / Transport pour le " +
              subscription + " ne peut prendre que la valeur 0 ou 1 et non " + current_purchased)      
    }
    if (state == 0) {
      if (current_purchased == 0) {
        continue
      }
      return ('Une adhésion existe pour un ' + subscription +
              ' sans adhésion déclarée pour un ' + adjusted_noncomp_subscription_categories[index-1])
    }
    if (state == -1 || state == 1) {
      state = current_purchased
      continue
    }
    return ("Error de vérification d'adhésion: state=" + state + 
            ", current_purchased" + current_purchased)
  }

  // 4- Go over the declared subscriptions and accumulate the number of subscriptions entered.
  //    that number must match the number of non rider member entered.
  var subscribed_non_rider_kid_number = 0
  for (var index in adjusted_noncomp_subscription_categories) {
    var subscription = adjusted_noncomp_subscription_categories[index]
    var found_non_rider_kid_number = subscription_map[subscription].PurchasedSubscriptionAmount()
    subscribed_non_rider_kid_number += found_non_rider_kid_number
  }
  if (subscribed_non_rider_kid_number != non_rider_kid_number) {
    return ("Le nombre d'" + Plural(subscribed_non_rider_kid_number, "adhésion") +
            " NON rider " + Plural(subscribed_non_rider_kid_number, "souscrite") +
	    " [" + subscribed_non_rider_kid_number + "] ne correspond pas au nombre " +
            "de NON " + Plural(non_rider_kid_number, "rider renseigné") +
	    " [" + non_rider_kid_number + "]")
  }

  // 5- The number of non rider adults found and the number of matching subscriptions must agree.
  var subscribed_non_rider_adult_number = subscription_map[getAdultString()].PurchasedSubscriptionAmount()
  if (subscribed_non_rider_adult_number != non_rider_adult_number) {
    return ("Le nombre d'" + Plural(subscribed_non_rider_adult_number, "adhésion") + " loisir adulte " +
	    Plural(subscribed_non_rider_adult_number, "souscrite") + " [" + subscribed_non_rider_adult_number + "] " +
            "ne correspond pas au nombre " +
            "de loisir adulte " + Plural(non_rider_adult_number, "renseigné") + " [" + non_rider_adult_number + "]")
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

function getTotalSkiPassesAmountBeforeRebate() {
  return getNumberAt(coord_total_ski_pass_amount)
}

// Return the rebate amount that may apply to the total
// paid for all ski passes
function getSkiPassesRebateAmount() {
  var rebate = getNumberAt(coord_rebate_family_of_4_amount)
  if (rebate) {
    return rebate
  }
  return getNumberAt(coord_rebate_family_of_5_amount)
}

function validateSkiPasses() {
  updateStatusBar("Validation des forfaits...", "grey", add=true)
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
  var count_adults = 0
  var family_size = 0
  var number_of_non_student_adults = 0
  var total_paid_skipass = 0
  var total_adult_skipasses = 0
  var total_student_skipasses = 0
  for (var skipass_name in ski_passes_map) {
    skipass = ski_passes_map[skipass_name]
    skipass.UpdatePurchasedSkiPassAmountFromTrix()
    var purchased_amount = skipass.PurchasedSkiPassAmount()
    if (purchased_amount < 0 || isNaN(purchased_amount)) {
      return (purchased_amount + ' forfait ' + skipass_name + 
              ' acheté n\'est pas un nombre valide')
    }
    // This count_adults update hinges on the fact that a student is necessarily an adult.
    if (purchased_amount > 0 && (skipass.IsStudent() || skipass.IsAdult())) {
      count_adults += purchased_amount
    }
    if (skipass.IsStudent()) {
      total_student_skipasses += purchased_amount
    }
    if (skipass.IsAdult()) {
      total_adult_skipasses += purchased_amount
    }
    // Family size MUST exclude students
    if (! skipass.IsStudent() && purchased_amount > 0) {
      if (skipass.IsAdult()) {
        number_of_non_student_adults += purchased_amount
      }
      // Toddler ski pass doesn't count toward a familly rebate...
      if (!skipass.IsToddler()) {
        family_size += purchased_amount
      }
      total_paid_skipass += skipass.GetTotalPrice() 
    }
  }

  if (total_adult_skipasses + total_student_skipasses != count_adults) {
    return ("[" + count_adults + "] "+ Plural(count_adults, "forfait adulte renseigné") +
            " pour [" + total_adult_skipasses + "] " + Plural(total_adult_skipasses, "forfait adulte attribué") +
            " et [" + total_student_skipasses + "] " + Plural(total_student_skipasses, "forfait étudiant attribué"))
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
      var message = ("["+ total_purchased + "]" + Plural(total_purchased, ' forfait') + ' ' +
                     zone1 + '/' + zone2 + Plural(total_purchased, ' acheté') + ' pour [' +
                     zone1_count + "] " + Plural(zone1_count, ' personne') +
                     ' dans cette tranche d\'âge (' + ski_passes_map[zone1].ValidDoBRangeMessage() + ")")
      message += ("\n\nIl ce peut que ce choix de forfait soit valide mais pas " +
                  "automatiquement vérifiable. C'est le cas lorsque plusieurs options  " +
                  "correctes existent comment par exemple un membre étudiant qui est aussi " +
                  "d'âge adulte ou un adulte qui prend un licence mais pas de forfait.")                     
      if (! displayYesNoPanel(augmentEscapeHatch(message))) {
        return 'BUTTON_NO'
      }
    }
  }

  // FIXME: Validation of the students. Also: this should be in the magic wand
  if (number_of_non_student_adults == 2) {
    var rebate_applied = false
    if (family_size == 4) {
      var rebate = -(total_paid_skipass * 0.1)
      setStringAt(coord_rebate_family_of_4_count , "1")
      setStringAt(coord_rebate_family_of_4_amount, rebate)
      rebate_applied = true
    }
    if (family_size >= 5) {
      var rebate = -(total_paid_skipass * 0.15)
      setStringAt(coord_rebate_family_of_5_count , "1")
      setStringAt(coord_rebate_family_of_5_amount, rebate)    
      rebate_applied = true
    }
    if (rebate_applied) {
      updateStatusBar("💰 Réduction appliquée", "grey", add=true)
    }
    SpreadsheetApp.flush();
  }
  return ''
}
