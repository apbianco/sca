// This validates the invoice sheet (more could be done BTW) and
// sends an email with the required attached pieces.
// This script runs with duplicates of the following shared doc: 
// shorturl.at/dkmrH

// Seasonal parameters
var season = "2021/2022"
var season_web = "saison-2021-2022"

// Spreadsheet parameters (row, columns, etc...)
coord_family_civility = [8, 3]
coord_family_name = [8, 4]
coord_family_email = [11, 3]
coord_cc = [11, 5]
coord_personal_message = [80, 3]
coord_family_phone = [81, 7]
coord_timestamp = [81, 2]
coord_parental_consent = [81, 5]
coord_status = [82, 3]
coord_generated_pdf = [82, 4]

// Some globals defined here to make changes easy:
var parental_consent_pdf = '1F4pfeJbiNB1VQrjPHAJbo0il1WEUTuZB'
var rules_pdf = '1tPia7eLaoUamKl9iulbEADoI_lwo5pz_'
var payment_refund = '1A3HNJ1iH2A0G9WiwJlBZxk788f-9QdaB' // Not used yet.
var parents_note_pdf = '1_DG8n1HdldZTrc5XX1b5ESqxr7pbP9yv'
var db_folder = '1UmSA2OIMZs_9J7kEFu0LSfAHaYFfi8Et'
var allowed_user = 'inscriptions.sca@gmail.com'
var email_loisir = 'sca.loisir@gmail.com'

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

function savePDF(blob, fileName) {
  blob = blob.setName(fileName)
  var file = DriveApp.createFile(blob);
  pdfId = DriveApp.createFile(blob).getId()
  DriveApp.getFileById(file.getId()).moveTo(DriveApp.getFolderById(db_folder))
  return file;
}

function createPDF(url) {
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=7'
      + '&portrait=true'
      + '&fitw=true'       
      + '&top_margin=0.1'              
      + '&bottom_margin=0.1'          
      + '&left_margin=0.1'             
      + '&right_margin=0.1'           
      + '&sheetnames=false'
      + '&printtitle=true'
      + '&pagenum=CENTER'
      + '&gridlines=false'
      + '&fzr=FALSE'
      // Note: this cell range doesn't seem to work.
      + '&r1=2'
      + '&r2=80'
      + '&c1=2'
      + '&c2=6'

  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { 
      Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
    },
  })
  return response.getBlob()
}

function getStringAt(coord) {
  var x = coord[0];
  var y = coord[1];
  return SpreadsheetApp.getActiveSheet().getRange(x,y).getValue().toString()
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
  var blob = createPDF(spreadsheet.getUrl())
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
          'mail_to': mail_to,
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

// Opening or reloading the sheet triggers this function.
// Clear the invoice/email status and file indicators
function onOpen() {
  clearRange(SpreadsheetApp.getActiveSheet(), coord_status);  
  clearRange(SpreadsheetApp.getActiveSheet(), coord_generated_pdf);  
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
  attachments = [pdf.getAs(MimeType.PDF)]

  setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                    coord_status, 
                    "⏳ Préparation et envoit du dossier...", "orange")
  SpreadsheetApp.flush()
  
  var civility = validation['civility'];
  var family_name = validation['family_name'];
  var mail_to = validation['mail_to'];
  var consent = validation['consent'];
  
  // Determine whether parental consent needs to be generated. If
  // that's the case, we generate additional attachment content.
  var rules, parents_note, parental_consent, parental_consent_text = ''
  if (consent = 'Nécessaire') {
    var consent_file = DriveApp.getFileById(parental_consent_pdf)
    var rules_file = DriveApp.getFileById(rules_pdf)
    var notes_file = DriveApp.getFileById(parents_note_pdf)
    attachments.push(consent_file.getAs(MimeType.PDF))
    attachments.push(rules_file.getAs(MimeType.PDF))
    attachments.push(notes_file.getAs(MimeType.PDF))
    
    parental_consent_text = (
      "<p>Il vous faut également compléter et signer l'autorisation " +
      "parentale fournie en attachment, couvrant le droit à l'image, le " +
      "règlement intérieur et les interventions médicales.</p>" +
      "<p>Le règlement intérieur mentionné dans l'autorisation parentale " +
      "est joint à ce message en attachement.</p>" +
      "<p>Vous trouverez également en attachement une note adressée aux " +
      "parents, merçi de la lire attentivement.</p>")
  }

  var subject = ("[Incription Ski Club Allevardin] " +
                 civility + ' ' + family_name +
		         ": Facture pour la saison " + season)

  // Collect the personal message and add it to the mail
  var personal_message_text = getStringAt(coord_personal_message)
  if (personal_message_text != '') {
      personal_message_text = '<p><b>Message personnel:</b>' +
      personal_message_text + '</p>'
  }
  
  // Fetch a possible phone number
  var phone = getStringAt(coord_family_phone)
  if (phone != 'Aucun') {
    phone = '<p>Merci de contacter ' + phone + '</p>'
  } else {
    phone = ''
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  email_options = {
    name: 'Ski Club Allevardin, gestion des inscriptions',
    to: mail_to,
    subject: subject,
    htmlBody:
      "<h3>Bonjour " + civility + " " + family_name + "</h3>" +
    
      personal_message_text + 
      phone +

      "<p>Votre facture pour la saison " + season +
      "est disponible en attachement. Veuillez contrôler " +
      "qu\'elle correspond à vos besoins.</p>" +
    
      "<p>Votre règlement est à retourner à:" +
    
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
    
      "<p>Nous vous remercions de la confiance que vous nous accordez " +
      "cette année.\n" +
    
      "<p>Des questions concernant cette facture? Répondez directement " +
      "à ce mail. Des questions concernant la saison " + season + " ? " +
      "Envoyez un mail à " + email_loisir + "</p>" +
    
      "~SCA",
    attachments: attachments
  }
   
  // Add CC if defined.
  var cc_to = getStringAt(coord_cc)
  if (cc_to != "") {
    email_options.cc = cc_to
  }
  
  // If this isn't a test, BCC Marie-Pierre. A test is something ran by
  // the ALEX or TEST trigger.
  if (spreadsheet.getName().toString().substring(0,5) != "TEST:" &&
      spreadsheet.getName().toString().substring(0,5) != "ALEX:") {
    email_options.bcc = "licence.sca@gmail.com"
  }
  
  MailApp.sendEmail(email_options)
  setRangeTextColor(SpreadsheetApp.getActiveSheet(),
                    coord_status, 
                    "✅ Dossier envoyé", "green")  
  displayPDFLink(pdf_file)
  SpreadsheetApp.flush()  
}
