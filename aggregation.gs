///////////////////////////////////////////////////////////////////////////////
// Update various aggregation trix:
//  - There is one that is used to project future level groups
//  - There is one used by accounting
///////////////////////////////////////////////////////////////////////////////

// Search for data.{first_name, last_name} starting at data.row_start in sheet.
// Return the index at which the element was found or the index of the first
// empty row. Return -1 if more than one element matched the search criteria.
function SearchEntry(sheet, data, positions) {
  var global_index = positions.row_start
  var index = global_index
  var stop_incrementing_index = false
  var first_empty_row = global_index
  var stop_updating_first_empty_row = false
  // slice() skips header data... Note: we have populated a rank/index column at position
  // 1 for a sufficient number of rows so that the sheet has data past the last element,
  // which allows us to determine there is an empty element past that.
  const allData = sheet.getDataRange().getValues().slice(index-1);
  const matchingRows = allData.filter(
      row => {
        if (row[positions.last_name] == '' && !stop_updating_first_empty_row) {
          first_empty_row = global_index
          stop_updating_first_empty_row = true
        }
        var found = row[positions.last_name] === data.last_name &&
                    row[positions.first_name] === data.first_name
        if (found) {
          stop_incrementing_index = true
        }
        if (! stop_incrementing_index) {
          index += 1
        }
        global_index += 1
        return found
      }
  )
  if (matchingRows.length > 1) {
    return -1
  }
  if (matchingRows.length == 0) {
    return first_empty_row
  }
  return index
}

function doUpdateAggregationTrix(data) {
  // Update the row at range in sheet with data
  function UpdateRow(sheet, row, data) {
    var column = 2
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
  var entire_range = sheet.getRange(license_trix_all_range)

  for (var index in data) {
    var row = SearchEntry(sheet, data[index],
                          {row_start: license_trix_row_start,
                           last_name:1,
                           first_name:2})
    if (row == -1) {
      // FIXME: What to signal to the user?
      continue
    }
    UpdateRow(sheet, row, data[index])
  }
  // Sort the spread sheet by last name and then first name and sync the spreadsheet.
  entire_range.sort([{column: entire_range.getColumn()}, {column: entire_range.getColumn()+5}])
  SpreadsheetApp.flush()
}

function updateAggregationTrix() {
  var family_dict = getListOfFamilyPurchasingALicense()
  var family = []
  for (var index in family_dict) {
    var family_member = family_dict[index]
    // Skip over an empty name
    if (family_dict[index].last_name == "") {
      continue
    }
    // Retain someone with a non comp license that has a ski level. Setting a ski level
    // signal the intention to be placed under the supervision of an instructor.
    if (isLevelRecreationalNonRider(family_member.level) || isLevelRider(family_member.level)) {
      // If the member is an adult, its category is changed to "Adulte"
      if (isAdult(family_member.dob)) {
        family_member.level = 'Adulte'
      }
      family.push(family_member)
    }
  }
  doUpdateAggregationTrix(family)
}

function doUpdateAccountingTrix(data) {
  // Update the row at range in sheet with data
  function UpdateRow(sheet, row, data) {
    var column = 2
    sheet.getRange(row, column).setValue(data.last_name)
    sheet.getRange(row, column+1).setValue(data.first_name)
    sheet.getRange(row, column+2).setValue(data.sex)
    sheet.getRange(row, column+3).setValue(data.dob)
    sheet.getRange(row, column+4).setValue(data.parent_city)
    sheet.getRange(row, column+5).setValue(data.level)
    sheet.getRange(row, column+6).setValue(data.license_type)
    sheet.getRange(row, column+7).setValue(data.license_number)
    sheet.getRange(row, column+8).setValue(data.license_fee)
    sheet.getRange(row, column+9).setValue(data.subscription_type)
    sheet.getRange(row, column+10).setValue(data.subscription_fee)
    sheet.getRange(row, column+11).setValue(data.skipass_type)
    sheet.getRange(row, column+12).setValue(data.skipass_fee)
    sheet.getRange(row, column+13).setValue(data.parent1_phone)
    sheet.getRange(row, column+14).setValue(data.parent1_email)
    sheet.getRange(row, column+15).setValue(data.parent2_phone)
    sheet.getRange(row, column+16).setValue(data.parent2_email)
  }

  function dispatchNonCompSubscriptions() {
    // Go through sections where a charge applies and match the charges with
    // entries in data
    var non_comp_subscriptions = createNonCompSubscriptionMap(sheet)
    for (const key in non_comp_subscriptions) {
      function levelRider(entry) { return isLevelRider(entry.level) }
      function levelRecreationalNonRider(entry) { return isLevelRecreationalNonRider(entry.level)}
      function levelAdult(entry) { return isAdult(new Date(entry.dob)) }
      var subscription = non_comp_subscriptions[key]
      subscription.UpdatePurchasedSubscriptionAmountFromTrix()
      var number_charges = subscription.PurchasedSubscriptionAmount()
      // No charge, no need to process that subscription type.
      if (number_charges == 0) {
        continue
      }
      var fee = subscription.SubscriptionAmount()
      // How to determine what applies - as we're going to all the entries in
      // data, this is the last filter to select a particular entry in data.
      var determination = (isSubscriptionAdult(key) ?
                           levelAdult : isLevelRider(key) ?
                                        levelRider : levelRecreationalNonRider)
      // Go over all entries and dispatch charges as possible
      for (var entry of data) {
        // Stop when we have dispatched all existing charges
        if (number_charges == 0) {
          break
        }
        // Do not change an entry that has already been set.        
        if ('subscription_type' in entry) {
          continue
        }
        // Determination has been previously picked to be the right
        // function.
        if (determination(entry)) {
          entry.subscription_type = subscription.HumanReadableName() + " loisir"
          entry.subscription_fee = fee
          number_charges -= 1
        }
      }
    }    
  }

  function dispatchCompSubscriptions() {
    var comp_subscriptions = createCompSubscriptionMap(sheet)
    for (const key in comp_subscriptions) {
      var subscription = comp_subscriptions[key]
      subscription.UpdatePurchasedSubscriptionAmountFromTrix()
      var number_charges = subscription.PurchasedSubscriptionAmount()
      // No charge, no need to process that subscription type.
      if (number_charges == 0) {
        continue
      }
      var fee = subscription.SubscriptionAmount()
      for (var entry of data) {
        // Stop when we have dispatched all existing charges
        if (number_charges == 0) {
          break
        }
        // Do not change an entry that has already been set.        
        if ('subscription_type' in entry) {
          continue
        }
        // Determination is done on the collected age...
        if (subscription.ValidateDoB(entry.dob)) {
          entry.subscription_type = subscription.HumanReadableName()
          entry.subscription_fee = fee
          number_charges -= 1
        }
      }
    }
  }

  function dispatchLicenses() {
    var licenses = createLicensesMap(sheet)
    var family = licenses[getNonCompFamilyLicenseString()]
    family.UpdatePurchasedLicenseAmountFromTrix()
    var number_charges = family.PurchasedLicenseAmount()
    if (number_charges > 0) {
      for (var entry of data) {
        // We pick the first one match as the recipient
        if (family.IsA(entry.license_type)) {
          if (number_charges > 0) {
            entry.license_fee = familly.LicenseAmount()
            number_charges -= 1
          } else {
            entry.license_fee = 'Famille'
          }
        }
      }
    }
    // Delete the one we just processed and go over the other type
    // of licenses.
    delete licenses[getNonCompFamilyLicenseString()]
    for (const key in licenses) {
      var license = licenses[key]
      license.UpdatePurchasedLicenseAmountFromTrix()
      var number_charges = license.PurchasedLicenseAmount()      
      // No charge, no need to process that subscription type.
      if (number_charges == 0) {
        continue
      }
      var fee = license.LicenseAmount()
      // Go over all entries and dispatch charges as possible
      for (var entry of data) {
        // Stop when we have dispatched all existing charges
        if (number_charges == 0) {
          break
        }
        // Do not change an entry that has already been set.        
        if ('license_fee' in entry) {
          continue
        }
        if (license.IsA(entry.license_type)) {
          entry.license_fee = fee
          number_charges -= 1
        }
      }
    }
  }

  function dispatchSkiPasses() {
    // If there's a rebate, we have a familly ski pass. Compute and assign the total amoun
    // amount paid in ski passes (rebate included) to the first familly member, mark the
    // others as 'Family' and we're done.
    var ski_pass_rebate_amount = getSkiPassesRebateAmount()
    if (ski_pass_rebate_amount) {
      var final_amount = getTotakSkiPassesAmountBeforeRebate() - ski_pass_rebate_amount;
      for (var entry of data) {
        if (ski_pass_rebate_amount) {
          entry.skipass_fee = final_amount
          entry.skipass_type = 'Famille'
          ski_pass_rebate_amount = 0
        } else {
          entry.skipass_type = 'Famille'
        }
      }
      return
    }
    var skipasses = createSkipassMap(sheet)
    for (const key in skipasses) {
      var skipass = skipasses[key]
      skipass.UpdatePurchasedSkiPassAmountFromTrix()
      var number_charges = skipass.PurchasedSkiPassAmount()
      // No charge, no need to process that subscription type.
      if (number_charges == 0) {
        continue
      }
      var fees = skipass.SkiPassAmount()
      for (var entry of data) {
        // Stop when we have dispatched all existing charges
        if (number_charges == 0) {
          break
        }
        // Do not change an entry that has already been set.        
        if ('skipass_fee' in entry) {
          continue
        }
        if (skipass.ValidateDoB(entry.dob)) {
          entry.skipass_fee = fees
          entry.skipass_type = key
          number_charges -= 1
        }
      }
    }
  }

  var sheet = SpreadsheetApp.openById(accounting_trix).getSheets()[0]
  // Dispatch all the data
  dispatchNonCompSubscriptions()
  dispatchCompSubscriptions()
  dispatchLicenses()
  dispatchSkiPasses()

  // Update all entries, sort the spreadsheet and sync it.
  var entire_range = sheet.getRange(accounting_trix_all_range)
  var max_row = 0
  for (var index in data) {
    var row = SearchEntry(sheet, data[index],
                          {row_start: accounting_trix_row_start,
                           last_name:1,
                           first_name:2})
    if (row == -1) {
      // FIXME: What to signal to the user?
      continue
    }
    max_row = Math.max(max_row, row)
    UpdateRow(sheet, row, data[index])
  }
  // Sort the spread sheet by last name and then first name and sync the spreadsheet.
  entire_range.sort([{column: entire_range.getColumn()}, {column: entire_range.getColumn()+5}])
  // ...
  var entireRowRange = sheet.getRange(max_row, 1, 1, sheet.getMaxColumns());
  entireRowRange.setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
  SpreadsheetApp.flush()
}

function updateAccountingTrix() {
  var family_dict = getListOfFamilyPurchasingALicense()
  var family = []
  for (var index in family_dict) {
    var family_member = family_dict[index]
    // Skip over an empty name
    if (family_dict[index].last_name == "") {
      continue
    }
    // Retain the right entries: only folks placed under the supervision of an
    // instructor
    if (isLevelComp(family_member.level) ||
        isLevelRecreationalNonRider(family_member.level) ||
        isLevelRider(family_member.level)) {
      family.push(family_member)
    }
  }
  doUpdateAccountingTrix(family)
}

///////////////////////////////////////////////////////////////////////////////
// Updating the tab tracking registration with issues in the aggregation trix
// with a link and some context.
///////////////////////////////////////////////////////////////////////////////

function updateProblematicRegistration(link, context) {
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
