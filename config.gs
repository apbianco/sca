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
  'CN Jeune (Compétition)':   [2011,  52,   5],
  'CN Adulte (Compétition)':  [2010,  53,   5],
}
//
// - A map of available subscription for competitors and their validations
//   dates. This map is used to create a map of properly configured non competitor
//   subscription objects. Edit this map when the year of validity is changing.
//
var comp_subscription_map = {
// What | Dob/Y | Dob/Y | Row | Col
  'U6':   [2020,  2021,   54,   5],
  'U8':   [2018,  2019,   55,   5],
  'U10':  [2016,  2017,   56,   5],
  // U12+ The first listed year is the LAST year of the category
  'U12+': [2015, -1,      57,   5],
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
var license_trix_row_start =               10
var license_trix_all_range =               'B10:H'
//
// Accounting aggregation trix to update when a new entry is added
//
var accounting_trix =                      '1_X6bL8HiDabmbyZdC0N3oQ11rlJ_IxWAo0i9uJiWv34'
//
// - ID of attachements to be sent with the invoice - some may change
//   from one season to an other when they are refreshed. These documents
//   are also part of the registration bundle sent to parents.
//
var legal_disclaimer_pdf =                  '1BcbIL1LUaL49OI8BGbLQQA-lxlijt_0m'
var rules_pdf =                             '1aTtQ4Hx3JJPgek_d1_NV4WRSBOfVT_yr'
var parents_note_pdf =                      '1fVb5J3o8YikPcn4DDAplt9X-XtP9QdYS' // TODO
var ffs_information_leaflet_pdf =           '1zxP502NjvVo0AKFt_6FCxs1lQeJnNxmV' // TODO
var ffs_information_leaflet_pages_to_sign = 'les pages 16 et 17'                // TODO
var autocertification_non_adult =           '1Ir-1TlAB7SoZcamRoX2_i1uk9_l19YUJ' // Note: reusing 2024/2025
var autocertification_adult =               '1rFAJghPwE-6xHzSWdRrMjP8gQKUldX6W' // Note: reusing 2024/2025
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
var coord_total =                     [70, 7]
var coord_rebate =                    [81, 4]
var coord_charge =                    [82, 4]
var coord_owed =                      [83, 7]
var coord_payment_validation_form =   [84, 7]
var coord_personal_message =          [90, 3]
var coord_timestamp =                 [91, 2]
var coord_version =                   [91, 3]
var coord_legal_disclaimer =          [91, 5]
var coord_ffs_medical_form =          [91, 7]
var coord_callme_phone =              [91, 9]
var coord_yolo =                      [92, 3]
var coord_status =                    [93, 4]
var coord_generated_pdf =             [93, 6]
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
// - Location where the total amount of ski passes is
//   stored
var coord_total_ski_pass            = [40, 6]
//
// - Parameters defining the valid ranges to be retained during the
//   generation of the invoice's PDF
//
var coords_pdf_row_column_ranges = {
    'start':                          [1, 0],
    'end':                            [91, 9]
}
