// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.
// This script runs with duplicates of the following shared doc: 
// shorturl.at/dkmrH

// Dev or prod?
var dev_or_prod = "dev"

// Seasonal parameters
var season = "2021/2022"
var season_web = "saison-2021-2022"

// Spreadsheet parameters (row, columns, etc...)
var coord_family_civility = [8, 3]
var coord_family_name = [8, 4]
var coord_family_email = [11, 3]
var coord_cc = [11, 5]
var coord_personal_message = [83, 3]
var coord_family_phone = [84, 7]
var coord_timestamp = [84, 2]
var coord_parental_consent = [84, 5]
var coord_status = [86, 4]
var coord_generated_pdf = [86, 6]

var coords_identity_lines = [16, 17, 18, 19, 20]
var coords_identity_cols  = [2, 3, 4, 5, 6]
var coords_pdf_row_column_ranges = {'start': [1, 0], 'end': [84, 7]}

// Some globals defined here to make changes easy:
var parental_consent_pdf = '1LaWS0mmjE8GPeendM1c1mCGUcrsBIUFc'
var rules_pdf = '1tNrkUkI2w_DYgXg9dRXPda_25gBh1TAG'
var parents_note_pdf = '1zkI5NapvYyLn_vEIxyejKJ4WzAvEco6z'
var pass_pdf = '1fsjge7JAuV3PTPXBnLbX9PGkSGW3nVHL'

var db_folder = '1UmSA2OIMZs_9J7kEFu0LSfAHaYFfi8Et'
var allowed_user = 'inscriptions.sca@gmail.com'
var email_loisir = 'sca.loisir@gmail.com'
var email_comp = 'skicluballevardin@gmail.com'
var email_dev = 'apbianco@gmail.com'
var email_license = (isProd() ?
                     'licence.sca@gmail.com': email_dev)

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
  for (var index in coords_identity_lines) {
    if (getStringAt([coords_identity_lines[index], 4]) != '') {
      var birth = new Date(getStringAt([coords_identity_lines[index], 4]))
      birth = Utilities.formatDate(birth, "GMT", "dd/MM/yyyy")
    } else {
      birth = "??/??/????"
    }
    var sex = getStringAt([coords_identity_lines[index], 6])
    if (sex == "") {
      sex = "?"
    }
    var city = getStringAt([coords_identity_lines[index], 5])
    if (city != "") {
      city = "(" + city + ")";
    }
    family.push({'first': getStringAt([coords_identity_lines[index], 2]),
                 'last': getStringAt([coords_identity_lines[index], 3]),
                 'birth': birth,
                 'city': city,
                 'sex': sex})
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

function setRangeTextColor(sheet, coord, text, color) {
  var x = coord[0];
  var y = coord[1];
  sheet.getRange(x,y).setValue(text);
  sheet.getRange(x,y).setFontColor(color);
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

function isEmpty(obj) {
  return Object.keys(obj).length === 0;
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
  // Make sure only an allowed user runs this.
  if (Session.getEffectiveUser() != allowed_user) {
    displayErrorPannel(
      SpreadsheetApp.getActiveSheet(),
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
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var operator = spreadsheet.getName().toString().split(':')[0]
  var family_name = spreadsheet.getName().toString().split(':')[1]
  if (operator == "TEST") {
    return
  }
  var family_dict = getFamilyDictionary() 
  var string_family_members = "<blockquote>\n"
  for (var index in family_dict) {
    if (family_dict[index]['last'] == "") {
      continue
    }
    string_family_members += ("<tt><b>" +
                              family_dict[index]['last'].toUpperCase() + "</b> " +
                              family_dict[index]['first'] + " " +
                              family_dict[index]['birth'] + " M/F=" +
                              family_dict[index]['sex'] + " " +
                              family_dict[index]['city'] + "</tt><br>\n")
  }
  string_family_members += "</blockquote>\n"
  
  var email_options = {
    name: family_name + ": nouvelle inscription",
    to: email_license,
    subject: family_name + ": nouvelle inscription",
    htmlBody:
      "<p>Licences nécessaires pour:</p>" +
      string_family_members +
      "<p>Dossier saisi par: " + operator + "</p>" +
      "<p>Facture en attachement</p>" +
      "<p>Merci!</p>",
    attachments: invoice,
  } 
  MailApp.sendEmail(email_options)
}

// This is what the [generate and send email] button runs.
function GeneratePDFAndSendEmailButton() {
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
                    "⏳ Préparation et envoit du dossier...", "orange")
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
  var phone = getStringAt(coord_family_phone)
  if (phone != 'Aucun') {
    phone = '<p>Merci de contacter ' + phone + '</p>'
  } else {
    phone = ''
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var email_options = {
    name: 'Ski Club Allevardin, gestion des inscriptions',
    to: mail_to,
    subject: subject,
    htmlBody:
      "<h3>Bonjour " + civility + " " + family_name + "</h3>" +
    
      personal_message_text + 
      phone +

      "<p>Votre facture pour la saison " + season + " " +
      "est disponible en attachement. Veuillez contrôler " +
      "qu\'elle correspond à vos besoins.</p>" +
    
      "<p>Si nécessaire, votre règlement est à retourner à:" +
    
      "<blockquote>" +
      "  <b>Marie-Pierre Béranger</b>,<br>" +
      "  44 Grange Merle.<br>" +
      "  <b><u>38580 Allevard</u></b><br>" +
      "</blockquote>" +
    
      "<p>Pour faciliter la gestion des inscriptions nous vous invitons " +
      "à accompagner vos chèques d\'une copie de la facture en attachement " +
      "ou de mentionner au dos de chacun de vos chèques " +
      "la référence suivante: " +
      "<b><code>" + spreadsheet.getName() + "</code></b></p>" +
    
      "<p>Certificats médicaux et photos nécessaires à l\'établissement " +
      "de l\'inscription sont à faire parvenir à " + allowed_user + "</p>" +
    
      parental_consent_text +
    
      "<p>Vous trouverez également en attachement une note adressée aux " +
      "parents, merçi de la lire attentivement.</p>" +
    
      "<p>Nous vous remercions de la confiance que vous nous accordez " +
      "cette année.</p>" +
    
      "<p>Des questions concernant cette facture? Répondez directement " +
      "à ce mail. Des questions concernant la saison " + season + " ? " +
      "Envoyez un mail à " + email_loisir + "(ski loisir) " +
      "ou à " + email_comp + " (ski compétition)</p>" +
    
      "~SCA ❄️ 🏔️ ⛷️ 🏂",
    attachments: attachments
  }
   
  // Add CC if defined.
  var cc_to = getStringAt(coord_cc)
  if (cc_to != "") {
    email_options.cc = cc_to
  }

  // Send the email  
  MailApp.sendEmail(email_options)
  // If this isn't a test, send a mail to Marie-Pierre. A test is something ran by
  // the ALEX or TEST trigger.
  maybeEmailLicenseSCA([attachments[0]]);
  
  setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                    coord_status, 
                    "✅ Dossier envoyé", "green")  
  displayPDFLink(pdf_file)
  SpreadsheetApp.flush()  
}
