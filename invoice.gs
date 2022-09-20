// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.
// This script runs with duplicates of the following shared doc: 
// shorturl.at/EJM58

// TODO: Authorisation: à fournier, etc...

// Dev or prod? "dev" sends all email to email_dev. Prod is the
// real thing: family will receive invoices, and so will email_license.
var dev_or_prod = "dev"

// Seasonal parameters - change for each season
// 
// - Name of the season
var season = "2022/2023"
//
// - Storage for the current season's database.
//
var db_folder = '1apITLkzOIkqCI7oIxpiKA5W_QR0EM3ey'
//
// - ID of attachements to be sent with the invoice - some may change
//   from one season to an other when they are refreshed.
//
// TODO: Change for 2022/2034
var parental_consent_pdf = '1LaWS0mmjE8GPeendM1c1mCGUcrsBIUFc'
var rules_pdf = '1tNrkUkI2w_DYgXg9dRXPda_25gBh1TAG'
var parents_note_pdf = '1zkI5NapvYyLn_vEIxyejKJ4WzAvEco6z'
var pass_pdf = '1fsjge7JAuV3PTPXBnLbX9PGkSGW3nVHL'

// Spreadsheet parameters (row, columns, etc...). Adjust as necessary
// when the master invoice is modified.
// 
// - Locations of family details:
//
var coord_family_civility = [6, 3]
var coord_family_name = [6, 4]
var coord_family_email = [9, 3]
var coord_cc = [9, 5]
//
// - Locations of various status line and collected input, located
//   a the bottom of the invoice.
// 
var coord_personal_message = [78, 3]
var coord_callme_phone = [79, 7]
var coord_timestamp = [79, 2]
var coord_parental_consent = [79, 5]
var coord_status = [81, 4]
var coord_generated_pdf = [81, 6]
//
// - Parameters for collecting familly members
// 
var coords_identity_lines = [14, 15, 16, 17, 18, 19];

//
// - Parameters defining the valid ranges to be retained during the
//   generation of the invoice's PDF
//
var coords_pdf_row_column_ranges = {'start': [1, 0], 'end': [80, 7]}

// - Range for the attributed license validation. Please change
//   to match both the coordinate and the cell values
//   FIXME: don't use n_rows?
var coords_attributed_licenses_start = [14, 7];
var coords_attributed_licenses_n_rows = 5;
var attributed_licenses_values = [
  // This entry must always be the first one...
  'Aucune',
  'CN Jeune (Loisir)',
  'CN Adulte (Loisir)',
  'CN Famille (Loisir)',
  'CN Dirigeant',
  'CN Jeune (Compétition)',
  'CN Adulte (Compétition)'];

// - DoB validation for a given type of license: change the
//   start of end of ranges not featuring a negative number
var attributed_licenses_dob_validation = {
  'Aucune':                  [-1,   2051],
  'CN Jeune (Loisir)':       [2007, 2050],
  'CN Adulte (Loisir)':      [1900, 2006],
  'CN Famille (Loisir)':     [-1,   2051],
  'CN Dirigeant':            [1900, 2004],
  'CN Jeune (Compétition)':  [2007, 2050],
  'CN Adulte (Compétition)': [1900, 2006]};

// Coordinates of where the various license purchases are indicated.
var coord_purchased_licenses = {
  // 'Aucune' doesn't exist.
  'CN Jeune (Loisir)':       [34, 5],
  'CN Adulte (Loisir)':      [35, 5],
  'CN Famille (Loisir)':     [36, 5],
  'CN Dirigeant':            [37, 5],
  'CN Jeune (Compétition)':  [44, 5],
  'CN Adulte (Compétition)': [45, 5]};

// Email configuration - these shouldn't change very often
var allowed_user = 'inscriptions.sca@gmail.com'
var email_loisir = 'sca.loisir@gmail.com'
var email_comp = 'skicluballevardin@gmail.com'
var email_dev = 'apbianco@gmail.com'
var email_license_ = 'licence.sca@gmail.com'
var email_license = (isProd() ? email_license_: email_dev)

function isProd() {
  return dev_or_prod == 'prod'
}

function Debug(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

function createHyperLinkFromURL(url, link_text) {
  return '=HYPERLINK("' + url + '"; "' + link_text + '")';
}

function clearRange(sheet, coord) {
  sheet.getRange(coord[0],coord[1]).clear();
}

function setRangeTextColor(sheet, coord, text, color) {
  var x = coord[0];
  var y = coord[1];
  sheet.getRange(x,y).setValue(text);
  sheet.getRange(x,y).setFontColor(color);
}

function getStringAt(coord) {
  var x = coord[0];
  var y = coord[1];
  return SpreadsheetApp.getActiveSheet().getRange(x, y).getValue().toString()
}

function getFamilyDictionary() {
  var family = []
  var no_license = attributed_licenses_values[0];
  for (var index in coords_identity_lines) {
    
    var first_name = getStringAt([coords_identity_lines[index], 2]);
    var last_name = getStringAt([coords_identity_lines[index], 3]);    
    if (first_name == "" || last_name == "") {
      continue;
    }

    if (getStringAt([coords_identity_lines[index], 4]) != '') {
      var birth = new Date(getStringAt([coords_identity_lines[index], 4]))
      birth = Utilities.formatDate(birth, "GMT", "dd/MM/yyyy")
    } else {
      birth = "??/??/????"
    }
    var city = getStringAt([coords_identity_lines[index], 5])
    if (city == "") {
      city = "\\";
    }
    var sex = getStringAt([coords_identity_lines[index], 6])
    if (sex == "") {
      sex = "?"
    }
    var license = getStringAt([coords_identity_lines[index], 7]);
    if (license == "" || license == no_license) {
      continue;
    }
    
    family.push({'first': first_name, 'last': last_name,
                 'birth': birth, 'city': city, 'sex': sex,
                 'license': license})
  }
  return family
}

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
      + '&top_margin=0.1' + '&bottom_margin=0.1' + '&left_margin=0.1' + '&right_margin=0.1'           
      + '&sheetnames=false' + '&printtitle=true' + '&pagenum=CENTER' + '&gridlines=false'
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

function displayErrorPanel(message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(message, ui.ButtonSet.OK);
}

// Validate a cell at (x, y) whose value is set via a drop-down
// menu. We rely on the fact that a cell not yet set a proper value
// always has the same value.  When the value is valid, it is
// returned. Otherwise, '' is returned.
function validateAndReturnDropDownValue(coord, message) {
  var value = getStringAt(coord)
  if (value == 'Choix non renseigné' || value == '') {
    displayErrorPanel(message)
    return ''
  }
  return value
}

// Cross check the attributed licenses with the ones selected for payment
function validateLicenseCrossCheck() {
  // First count how many licenses have been attributed to the
  // registered family. We use the values (after validation) directly

  function validationToString(v) {
    var to_return = "";
    for (const [key, value] of Object.entries(v)) {
      to_return += (key + ": " + value + "\n");
    }
    return to_return;
  }
  
  var attributed_licenses = {}
  attributed_licenses_values.forEach(function(key) {
    attributed_licenses[key] = 0;
  });
  
  // Capture the no license marker, we're going to use it a lot.
  // This is the reason why it should be the first element in 
  // the attributed_licenses_values array.
  var no_license = attributed_licenses_values[0];
  var purchased_licenses = {}
  attributed_licenses_values.forEach(function(key) {
    // Entry indicating no license is skipped because it can't
    // be collected.
    if (key != no_license) {
      purchased_licenses[key] = 0;
    }
  });  
  
  // Collect the attributed licenses into a hash
  var attributed_licenses_row = coords_attributed_licenses_start[0];
  var col = coords_attributed_licenses_start[1];
  for (row = attributed_licenses_row;
       row <= attributed_licenses_row + coords_attributed_licenses_n_rows;
       row ++ ) {
    var value = getStringAt([row, col]);
    if (value === '') {
      value = no_license;
    }
    // You can't have no first/last name and an assigned license
    var first_name = getStringAt([row, 2]);
    var last_name = getStringAt([row, 3]);
    // If there's no name on that row, the only possible value is None
    if (first_name === '' && last_name === '') {
      if (value != no_license) {
        return "'" + value + "' attribuée à un membre de famile inexistant!";
      }
      continue;
    }
    var found = false;
    attributed_licenses_values.forEach(function(key) {
      if (value === key) {
        attributed_licenses[key] += 1;
        found = true;
        return;
      }
    });
    if (! found) {
      return "'" + value + "' n'est pas une license attribuée possible!";
    }
    // Validate DoB and the type of license
    var dob = new Date(getStringAt([row, 4]))
    dob = Number(Utilities.formatDate(dob, "GMT", "yyyy"));
    if (dob < 1900 || dob > 2050) {
      return first_name + " " + last_name + " a une année de naissance erronnée: " + dob;
    }
    dob_start = attributed_licenses_dob_validation[value][0];
    dob_end = attributed_licenses_dob_validation[value][1];
    if (dob < dob_start || dob > dob_end) {
      return (first_name + " " + last_name + ": l'année de naissance " + dob +
              " ne correspond pas à la license choisie: '" + value + "' (" +
              dob_start + "/" + dob_end + ")");
    }
  }
  
  // Collect the selected licenses into a hash
  attributed_licenses_values.forEach(function(key) {
    // Entry indicating no license is skipped because it can't
    // be collected.
    if (key != no_license) {
      var row = coord_purchased_licenses[key][0];
      var col = coord_purchased_licenses[key][1];
      purchased_licenses[key] = Number(SpreadsheetApp.getActiveSheet().getRange(row, col).getValue());
    }
  });

  if (! isProd()) {
    Debug(validationToString(attributed_licenses));
    Debug(validationToString(purchased_licenses));
  }
  
  // Perform the verification
  // FIXME: use for (var index in attributed_licenses_values)
  for (var index in attributed_licenses_values) {
    // Entry indicating no license is skipped because it can't
    // be collected.
    var key = attributed_licenses_values[index];
    if (key != no_license) {
      if (purchased_licenses[key] != attributed_licenses[key]) {
        return ("Le nombre de licences '" + key + "' selectionée(s) (au nombre de " +
                attributed_licenses[key] + ")\n" +
                "ne correspond pas au nombre de licences achetée(s) (au nombre de " +
                purchased_licenses[key] + ")");        
      }
    }
  }
  return "";
}

function isEmpty(obj) {
  return Object.keys(obj).length === 0;
}

function getOperator() {
  return SpreadsheetApp.getActiveSpreadsheet().getName().toString().split(':')[0]
}

function getFamilyName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName().toString().split(':')[1]
}

// Runs when the [secure authorization] button is pressed.
function GetAuthorization() {
  ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
}

// Create the invoice as a PDF: first create a blob and then save
// the blob as a PDF and move it to the <db>/<OPERATOR:FAMILY>
// directory. Return the PDF file ID.  
function generatePDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var blob = createPDF(spreadsheet)
  var pdf_filename = spreadsheet.getName() + '.pdf'
  var file = savePDF(blob, pdf_filename)
  
  var spreadsheet_folder_id =
    DriveApp.getFolderById(spreadsheet.getId()).getParents().next().getId()
  DriveApp.getFileById(file.getId()).moveTo(
    DriveApp.getFolderById(spreadsheet_folder_id))
  return file;
}

// Validate the invoice and return a dictionary of values
// to use during invoice generation.
function validateInvoice() {
  if (! isProd()) {
    Debug('Cette facture est en mode developpement. Aucun email ne sera envoyé, ' +
          'ni à la famile ni à ' + email_license_ + '.\n\n' +
          'Vous pouvez néamoins continuer et un dossier sera préparé et ' +
          'les mails serons envoyés à ' + email_dev + '.\n\n' +
          'Contacter ' + email_dev + 'pour obtenir plus d\'aide.')
          
  }
  
  // Make sure only an allowed user runs this.
  if (Session.getEffectiveUser() != allowed_user) {
    displayErrorPanel(
      "Vous n'utilisez pas cette feuille en tant que " +
      allowed_user + ".\n\n" +
      "Veuillez vous connecter d'abord à ce compte avant " +
      "d'utiliser cette feuille.");
    return {};
  }
  
  // Validation: proper civility
  var civility = validateAndReturnDropDownValue(
    coord_family_civility,
    "Vous n'avez pas renseigné de civilité")
  if (civility == '') {
    return {};
  }
  
  // Validation: a family name
  var family_name = getStringAt(coord_family_name)
  if (family_name == '') {
    displayErrorPanel(
      "Vous n'avez pas renseigné de nom de famille ou " +
      "vous avez oublié \n" +
      "de valider le nom de famille par [return] ou [enter]...")
    return {};
  }

  // Validation: proper email adress.
  var mail_to = getStringAt(coord_family_email)
  if (mail_to == '') {
    displayErrorPanel(
      "Vous n'avez pas saisi d'adresse email principale ou " +
      "vous avez oublié \n" +
      "de valider l'adresse email par [return] ou [enter]...")
    return {}
  }
  
  // Validation: parental consent set.
  var consent = validateAndReturnDropDownValue(
    coord_parental_consent,
    "Vous n'avez pas renseigné la nécessitée ou non de devoir " +
    "fournir une autorisation parentale.");
  if (consent == '') {
    return {}
  }
  
  var license_cross_check_error = validateLicenseCrossCheck();
  if (license_cross_check_error) {
    displayErrorPanel(license_cross_check_error);
    return {}
  }

  // Update the timestamp. 
  setRangeTextColor(SpreadsheetApp.getActiveSheet(), coord_timestamp,
                    'Dernière MAJ le ' +
                    Utilities.formatDate(new Date(),
                                         Session.getScriptTimeZone(),
                                         "dd-MM-YY, HH:mm"), 'black')

  return {'civility': civility,
          'family_name': family_name,
          'mail_to': isProd() ? mail_to : email_dev,
          'consent': consent};
}

function displayPDFLink(pdf_file, offset) {
  var link = createHyperLinkFromURL(pdf_file.getUrl(),
                                    "📁 Ouvrir " + pdf_file.getName())
  x = coord_generated_pdf[0]
  y = coord_generated_pdf[1]
  SpreadsheetApp.getActiveSheet().getRange(x, y).setFormula(link); 
}

// This is what the [generate] button runs
function GeneratePDFButton() {
  var validation = validateInvoice();
  if (isEmpty(validation)) {
    return;
  }
  setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                    coord_status, 
                    "⏳ Préparation de la facture...", "orange")  
  SpreadsheetApp.flush()
  displayPDFLink(generatePDF());
}

function maybeEmailLicenseSCA(invoice) {
  var operator = getOperator()
  var family_name = getFamilyName()

  var family_dict = getFamilyDictionary() 
  var string_family_members = "<blockquote>\n"
  for (var index in family_dict) {
    if (family_dict[index]['last'] == "") {
      continue
    }
    string_family_members += ("<tt>" +
                              "Nom: <b>" + family_dict[index]['last'].toUpperCase() + "</b><br>" +
                              "Prénom: " + family_dict[index]['first'] + "<br>" +
                              "Naissance: " + family_dict[index]['birth'] + "<br>" +
                              "Fille/Garçon: " + family_dict[index]['sex'] + "<br>" +
                              "Ville de Naissance: " + family_dict[index]['city'] + "<br>" +
                              "Licence: " + family_dict[index]['license'] + "<br>" +
                              "----------------------------------------------------</tt><br>\n");
  }
  string_family_members += "</blockquote>\n"
  
  var email_options = {
    name: family_name + ": nouvelle inscription",
    to: email_license,
    subject: family_name + ": nouvelle inscription",
    htmlBody:
      "<p>Licence(s) nécessaire(s) pour:</p>" +
      string_family_members +
      "<p>Dossier saisi par: " + operator + "</p>" +
      "<p>Facture en attachement</p>" +
      "<p>Merci!</p>",
    attachments: invoice,
  } 
  MailApp.sendEmail(email_options)
}

// This is what the [generate and send invoice] button runs.
function GeneratePDFAndSendEmailButton() {
  generatePDFAndMaybeSendEmail(true)
}

// This is what the [generate invoice] button runs.
function GeneratePDFButton() {
  generatePDFAndMaybeSendEmail(false)
}

function generatePDFAndMaybeSendEmail(send_email) {
  var validation = validateInvoice();
  if (isEmpty(validation)) {
    return;
  }
  setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                    coord_status, 
                    "⏳ Préparation de la facture...", "orange")
  SpreadsheetApp.flush()
  
  // Generate and prepare attaching the PDF to the email
  var pdf_file = generatePDF();
  var pdf = DriveApp.getFileById(pdf_file.getId());
  var attachments = [pdf.getAs(MimeType.PDF)]

  setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                    coord_status, 
                    "⏳ Génération " +
                    (send_email? "et envoit " : " ") + "du dossier...", "orange")
  SpreadsheetApp.flush()

  // 2021-2022: FFS sanitary pass documentation
  attachments.push(DriveApp.getFileById(pass_pdf).getAs(MimeType.PDF))
  
  var civility = validation['civility'];
  var family_name = validation['family_name'];
  var mail_to = validation['mail_to'];
  var consent = validation['consent'];
  
  // Determine whether parental consent needs to be generated. If
  // that's the case, we generate additional attachment content.
  var pass, rules, parents_note, parental_consent, parental_consent_text = ''
  if (consent == 'Nécessaire') {
    attachments.push(DriveApp.getFileById(parental_consent_pdf).getAs(MimeType.PDF))
    attachments.push(DriveApp.getFileById(rules_pdf).getAs(MimeType.PDF))
    
    parental_consent_text = (
      "<p>Il vous faut également compléter et signer l'autorisation " +
      "parentale fournie en attachment, couvrant le droit à l'image, le " +
      "règlement intérieur et les interventions médicales.</p>" +
      "<p>Le règlement intérieur mentionné dans l'autorisation parentale " +
      "est joint à ce message en attachement.</p>")
  }
  
  // Insert the note for the parents anyways
  attachments.push(DriveApp.getFileById(parents_note_pdf).getAs(MimeType.PDF))
  
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
      "qu\'elle correspond à vos besoins.</p>" +
    
      "<p>Certificats médicaux et photos nécessaires à l\'établissement " +
      "de l\'inscription sont à faire parvenir à " + allowed_user + "</p>" +
    
      parental_consent_text +
    
      "<p>Vous trouverez également en attachement une note adressée aux " +
      "parents, merçi de la lire attentivement.</p>" +

      "<p>Des questions concernant cette facture? Contacter Marie-Pierre: " +
      "marie-pierreberanger@orange.fr (06 21 18 00 89).</p>" +
      "<p>Des questions concernant la saison " + season + " ? " +
      "Envoyez un mail à " + email_loisir + " (ski loisir) " +
      "ou à " + email_comp + " (ski compétition)</p>" +
    
      "<p>Nous vous remercions de la confiance que vous nous accordez " +
      "cette année.</p>" +
    
      "~SCA ❄️ 🏔️ ⛷️ 🏂",
    attachments: attachments
  }

  // Add CC if defined.
  var cc_to = getStringAt(coord_cc)
  if (cc_to != "") {
    email_options.cc = isProd() ? cc_to : email_dev
  }

  if (send_email) {
    // Send the email  
    MailApp.sendEmail(email_options)
    maybeEmailLicenseSCA([attachments[0]]);

    setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                      coord_status, 
                      "✅ Dossier envoyé", "green")  
  } else {
    setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                      coord_status, 
                      "✅ Dossier généré", "green")  
  }
  displayPDFLink(pdf_file)
  SpreadsheetApp.flush()  
}

