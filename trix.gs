// The DB folder for the CURRENT season
var license_trix = '13akc77rrPaI6g6mNDr13FrsXjkStShviTnBst78xSVY'

function Log(s) {
  console.log(s)
}

function UpdateTrix(data) {
  function SearchEntry(sheet, range, data) {
    var finder = range.createTextFinder(data.first_name)
    while (true) {
      var range = finder.findNext()
      if (range == null) {
        return null
      }
      var row = range.getRow()
      var col = range.getColumn()
      if (sheet.getRange(row, col+1).getValue().toString() == data.last_name) {
        return sheet.getRange(row, col)
      }
    }
  }

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
    SpreadsheetApp.flush()
  }

  var sheet = SpreadsheetApp.openById(license_trix).getSheetByName('FFS');
  var last_name_range = sheet.getRange('B5:B')
  var res = SearchEntry(sheet, last_name_range, data)
  if (res != null) {
    UpdateRow(sheet, res, data)
  }
  console.log('Done')
}

class TrixUpdate {
  constructor(first_name, last_name, sex, dob, license_number, level) {
    this.first_name = first_name
    this.last_name = last_name
    this.sex = sex
    this.dob = dob
    this.license_number = license_number
    this.level
  }
}

function Run() {
  data = new TrixUpdate('Xulu', 'Albert', 'M', '12/12/2022', '12345AABB', 'Novice')
  UpdateTrix(data)
}
