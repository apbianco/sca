// The trix to update this season and several ranges that exist in that trix.
var license_trix = '13akc77rrPaI6g6mNDr13FrsXjkStShviTnBst78xSVY'

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
    sheet.getRange(row,column+6).setValue(data.level)
  }

  // Open the spread sheet, insert the name if the operation is possible. Sync
  // the spreadsheet.
  var sheet = SpreadsheetApp.openById(license_trix).getSheetByName('FFS2');
  var last_name_range = sheet.getRange('B7:B')
  var entire_range = sheet.getRange('B7:N')
  var res = SearchEntry(sheet, last_name_range, data)
  if (res != null) {
    UpdateRow(sheet, res, data)
  }

  // Sort the spread sheet by last name and then first name and sync the spreadsheet.
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
  data = new TrixUpdate('Regis', 'Skimal', 'M', '12/03/2015', 'F124199990412', 'Étoile 3')
  UpdateTrix(data)
  console.log('Done')
}

