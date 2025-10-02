///////////////////////////////////////////////////////////////////////////////
// Seasonal parameters - change for each season or change when the trix
// is changing (adding new rows/columns, etc...)
///////////////////////////////////////////////////////////////////////////////
//
// - Name of the season
var season = "2025/2026"
//
// - A map of available licenses and validation dates. This map is used
//   to create a map of properly configured license objects. Edit this
//   map when the year of validity is changing.
//
var licenses_configuration_map = {
  'CN Jeune (Loisir)':        [2011,  42,   5],
  'CN Adulte (Loisir)':       [2010,  43,   5],
  'CN Famille (Loisir)':      [-1,    44,   5],
  'CN Dirigeant':             [2010,  45,   5],
  'CN Jeune (Compétition)':   [2011,  53,   5],
  'CN Adulte (Compétition)':  [2010,  54,   5],
}
var basic_subscription_coord = [41, 5]
//
// - A map of available subscription for competitors and their validations
//   dates. This map is used to create a map of properly configured non competitor
//   subscription objects. Edit this map when the year of validity is changing.
//
var comp_subscription_map = {
// What | Dob/Y | Dob/Y | Row | Col
  'U6':   [2020,  2021,   55,   5],
  'U8':   [2018,  2019,   56,   5],
  'U10':  [2016,  2017,   57,   5],
  // U12+ The first listed year is the LAST year of the category
  'U12+': [2015, -1,      58,   5],
}
// Number of kids per competitor families
var comp_kids_per_family = 4
//
// - A map of available ski passes and validation dates. This map is used
//   to create a map of properly configured skip pass objects. Edit this
//   map when the year or age of validity is changing.
//
var skipass_configuration_map = {
// What     | Dob/Y1n| Dob/Y2 | Row | Col
  // ageVerificationRangeIncluded
  'Senior':   [70,     75,      25,   5],
  // ageVerificationStrictlyOldOrOlder
  'Vermeil':  [76,    -1,       26,   5],
  // ageVerificationBornBeforeDateIncluded (Jan 1st Y1) + 
  // ageVerificationStrictlyYounger
  'Adulte':   [2007,   70,      27,   5],
  // ageVerificationBornBetweenDatesIncluded (Jan 1st Y1, Dec 31 Y2)
  'Étudiant': [1995,   2006,    28,   5],
  'Junior':   [2007,   2014,    29,   5],
  'Enfant':   [2015,   2019,    30,   5],
  // ageVerificationBornAfterDateIncluded (Jan 1st Y1)
  'Bambin':   [2020,   -1,      31,   5],
}
// The row offset to add to an entry in skipass_configuration_map to
// obtain the corresponding row in the 3 domains offer. (25+8=33)
var skipass_configuration_map_3d_row_offset = 8
//
// - Storage for the current season's database.
//
var db_folder =                            '1dLImiBFObJDxSS1XJOmSSq2aiIVLKqai'
//
//
// Level aggregation trix to update when a new entry is added and the row at which
// data starts
//
var license_trix =                         '1TJrV0x_y387WZ4wKYqVp1wnWd-lph4CmUGxe0MtfhYk'
var license_trix_pb_sheet =                'Dossiers problématiques'
var license_trix_pb_sheet_search_range =   'A2:A'
var license_trix_pb_sheet_whole_range =    'A2:D'
var license_trix_ffs_sheet =               'Groupes de niveaux'
var license_trix_ffs_row_start =           10
var license_trix_ffs_all_range =           'B10:H'
//
// Accounting aggregation trix to update when a new entry is added
//
var accounting_trix =                      '1_X6bL8HiDabmbyZdC0N3oQ11rlJ_IxWAo0i9uJiWv34'
var accounting_trix_sheet =                'Pointage tréso'
var accounting_trix_row_start =            3
var accounting_trix_all_range =            'A3:Q'
//
// - ID of attachements to be sent with the invoice - some may change
//   from one season to an other when they are refreshed. These documents
//   are also part of the registration bundle sent to parents.
//
var legal_disclaimer_pdf =                  '1BcbIL1LUaL49OI8BGbLQQA-lxlijt_0m'
var rules_pdf =                             '1aTtQ4Hx3JJPgek_d1_NV4WRSBOfVT_yr'
var parents_note_pdf =                      '1IYzKukze9O7-ZFUiRI5x12iZAQmr0fIF'
var ffs_information_leaflet_pdf =           '12DH7chhB-Ye8S29VuGu0lkuCHl5SRJVO'
var ffs_information_leaflet_pages_to_sign = 'les pages 16 et 17'
var autocertification_non_adult =           '1rFAJghPwE-6xHzSWdRrMjP8gQKUldX6W'
var autocertification_adult =               '1UrAbl7mKxJG7jdya8P32kvSJ5Mn1BQqZ'
//
// - Spreadsheet parameters (row, columns, etc...). Adjust as necessary
//   when the master invoice is modified.
// 
// - Locations of family details:
//
var coord_family_civility =           [6,  3]
var coord_family_name =               [6,  4]
var coord_family_street =             [8,  3]
var coord_family_zip =                [8,  4]
var coord_family_city =               [8,  5]
var coord_family_email =              [9,  3]
var coord_cc =                        [9,  5]
var coord_family_phone1 =             [10, 3]
var coord_family_phone2 =             [10, 5]
//
// - Locations of various status line and collected input, located
//   a the bottom of the invoice.
// 
var coord_total =                     [71, 7]
var coord_rebate =                    [82, 4]
var coord_charge =                    [83, 4]
var coord_owed =                      [84, 7]
var coord_payment_validation_form =   [85, 7]
var coord_personal_message =          [91, 3]
var coord_timestamp =                 [92, 2]
var coord_version =                   [92, 3]
var coord_legal_disclaimer =          [92, 5]
var coord_ffs_medical_form =          [92, 7]
var coord_callme_phone =              [92, 9]
var coord_yolo =                      [93, 3]
var coord_status =                    [94, 4]
var coord_generated_pdf =             [94, 6]
//
// - Rows where the family names are entered
// 
var coords_identity_rows = [14, 15, 16, 17, 18, 19];
//
// - Row where the non competitor subscriptions start and
//   the column at which data exists
var coord_noncomp_start_row =         46
var coord_noncomp_column =            5
//
// - Columns where information about family members can be found
//
var coord_first_name_column =         2
var coord_last_name_column =          3
var coord_dob_column =                4
var coord_cob_column =                5
var coord_sex_column =                6
var coord_level_column =              7
var coord_license_column =            8
var coord_license_number_column =     9
//
// - Location where the cells computing the family pass
//   rebates are defined.
var coord_rebate_family_of_4_amount = [23, 4]
var coord_rebate_family_of_4_count  = [23, 5]
var coord_rebate_family_of_5_amount = [24, 4]
var coord_rebate_family_of_5_count  = [24, 5]
// - Location where total number of order ski passes is stored and where
//   the total amount paid for ski passes is due (before rebate)
var coord_total_ski_pass            = [40, 6]
var coord_total_ski_pass_amount    =  [40, 9]
//
// - Parameters defining the valid ranges to be retained during the
//   generation of the invoice's PDF
//
var coords_pdf_row_column_ranges = {
    'start':                          [1, 0],
    'end':                            [92, 9]
}
