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
  'CN Jeune (Loisir)':       2010,
  'CN Jeune (Compétition)':  2010,
  'CN Adulte (Loisir)':      2009,
  'CN Adulte (Compétition)': 2009,
  'CN Dirigeant':            2009,
}
//
// - A map of available subscription for competitors and their validations
//   dates. This map is used to create a map of properly configured non competitor
//   subscription objects. Edit this map when the year of validity is changing.
//
var comp_subscription_map = {
// What | Dob/Y | Dob/Y |  Row  | Col
  'U8':   [2018,  2019],
  'U10':  [2016,  2017],
  'U12+': [2015]
}
// Number of kids per competitor families
var comp_kids_per_family = 4
//
// - A map of available ski passes and validation dates. This map is used
//   to create a map of properly configured skip pass objects. Edit this
//   map when the year or age of validity is changing.
//
var skipass_configuration_map = {
// What     | Dob/Y | Dob/Y | Row | Col
  'Senior':   [70,    74,     25,   5],
  'Vermeil':  [75,   -1,      26,   5],
  'Adulte':   [2006,  70,     27,   5],
  'Étudiant': [1995,  2006,   28,   5],
  'Junior':   [2007,  2014,   29,   5],
  'Enfant':   [2015,  2019,   30,   5],
  'Bambin':   [2020, -1,      31,   5],
}
var skipass_configuration_map_3d_row_offset = 2
//
// - Storage for the current season's database.
//
var db_folder =                            '1GmOdaWlEwH1V9xGx3pTp1L3Z4zXQdjjn'
//
//
// Level aggregation trix to update when a new entry is added
//
var license_trix =                         '1tR3HvdpXWwjziRziBkeVxr4CIp10rHWyfTosv20dG_I'
//
// - ID of attachements to be sent with the invoice - some may change
//   from one season to an other when they are refreshed. These documents
//   are also part of the registration bundle sent to parents.
//
var legal_disclaimer_pdf =                  '18jFQWTmLnmBa9HGmPkFS58xr0GjNqERu'
var rules_pdf =                             '1U-eeiEFelWN4aHMwjHJ9IQRH3h2mZJoW'
var parents_note_pdf =                      '1fVb5J3o8YikPcn4DDAplt9X-XtP9QdYS'
var ffs_information_leaflet_pdf =           '1zxP502NjvVo0AKFt_6FCxs1lQeJnNxmV'
var ffs_information_leaflet_pages_to_sign = 'les pages 16 et 17'
var autocertification_non_adult =           '1Ir-1TlAB7SoZcamRoX2_i1uk9_l19YUJ'
var autocertification_adult =               '1nDNByrpln58YET8Gqy6vOXkSQb_aBTrf'
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
var coord_total =                     [65, 7]
var coord_rebate =                    [76, 4]
var coord_charge =                    [77, 4]
var coord_owed =                      [78, 7]
var coord_payment_validation_form =   [79, 7]
var coord_personal_message =          [85, 3]
var coord_timestamp =                 [86, 2]
var coord_version =                   [86, 3]
var coord_legal_disclaimer =          [86, 5]
var coord_ffs_medical_form =          [86, 7]
var coord_callme_phone =              [86, 9]
var coord_yolo =                      [87, 3]
var coord_status =                    [88, 4]
var coord_generated_pdf =             [88, 6]
//
// - Rows where the family names are entered
// 
var coords_identity_rows = [14, 15, 16, 17, 18, 19];
//
// - Row where competitor subscription starts
var coord_comp_start_row =      53
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
    'end':                            [86, 9]
}
