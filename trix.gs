// The trix to update this season and several ranges that exist in that trix.
var license_trix = '13akc77rrPaI6g6mNDr13FrsXjkStShviTnBst78xSVY'

// For each level, the column offset relative to the last_name_range for
// a given level
var levels_to_columns_map = {
  'FIRST': 7,
  'Débutant/Ourson': 7,
  'Flocon': 8,
  'Étoile 1': 9,
  'Étoile 2': 10,
  'Étoile 3': 11,
  'Bronze': 12,
  'Argent': 13,
  'Or': 14,
  'Ski/Fun': 15,
  'Rider': 16,
  'Snow Découverte': 17,
  'Snow 1': 18,
  'Snow 2': 19,
  'Snow 3': 20,
  'Snow Expert': 21,
  'LAST': 21
}

function UpdateTrix(data) {

  function FirstEmptySlotRange(sheet, range) {
    // Range must be a column
    var values = range.getValues();
    var ct = 0;
    while ( values[ct] && values[ct][0] != "" ) {
      ct++;
    }
    return sheet.getRange(range.getRow() + ct,range.getColumn())
  }
  
  // Search for elements in data in sheet over range. Returns null if nothing can be found
  function SearchEntry(sheet, range, data) {
    var finder = range.createTextFinder(data.last_name)
    while (true) {
      var current_range = finder.findNext()
      if (current_range == null) {
        return FirstEmptySlotRange(sheet, range)
      }
      var row = current_range.getRow()
      var col = current_range.getColumn()
      // This assumes that first_name will be found at col+1 relative to last_name.
      if (sheet.getRange(row, col+1).getValue().toString() == data.first_name) {
        return sheet.getRange(row, col)
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
    var dob_year = new RegExp("[0-9]+/[0-9]+/([0-9]+)", "gi").exec(data.dob)[1];
    sheet.getRange(row,column+5).setValue(dob_year)
    // Insert the level after having determined which column it should go to.
    if (data.level in levels_to_columns_map) {
      // First clear the entire level row before inserting the marker.
      sheet.getRange(row, levels_to_columns_map['FIRST'],
                     row, levels_to_columns_map['LAST']).clearContent()
      var offset_level = levels_to_columns_map[data.level]
      sheet.getRange(row,column+offset_level).setValue(1)
    }
  }

  // Open the spread sheet, insert the name if the operation is possible. Sync
  // the spreadsheet.
  var sheet = SpreadsheetApp.openById(license_trix).getSheetByName('FFS');
  var last_name_range = sheet.getRange('B5:B')
  var entire_range = sheet.getRange('B5:Y')
  var res = SearchEntry(sheet, last_name_range, data)
  if (res != null) {
    UpdateRow(sheet, res, data)
    SpreadsheetApp.flush()
  }

  // Sort the spread sheet and sync the spreadsheet again
  // Totally counter intuitive:
  var x = entire_range.getColumn()
  entire_range.sort([{column: entire_range.getColumn()}, {column: entire_range.getColumn()+1}])
  SpreadsheetApp.flush()
}

class TrixUpdate {
  constructor(first_name, last_name, sex, dob, license_number, level) {
    this.first_name = first_name
    this.last_name = last_name
    this.sex = sex
    this.dob = dob
    this.license_number = license_number
    this.level = level
  }
}

function Run() {
  console.log("X part of code is running here") 
  data = new TrixUpdate('FN', 'BBB_LN', 'M', '12/12/2022', '12345AABB', 'Étoile 3')
  UpdateTrix(data)
  data = new TrixUpdate('FN', 'BBB_LN', 'M', '12/12/2022', '12345AABB', 'Snow 1')
  UpdateTrix(data)
  data = new TrixUpdate('Zorra', 'La rousse', 'F', '12/12/2005', 'F12414412', 'Débutant/Ourson')
  UpdateTrix(data)
  console.log('Done')
}
