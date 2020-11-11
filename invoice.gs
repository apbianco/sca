// This code validates a newly generation subscription sheet and sends
// the invoice email, BCCing Marie-Pierre.

// Some globals defined here to make changes easy:
var parental_consent_pdf = '1F4pfeJbiNB1VQrjPHAJbo0il1WEUTuZB'
var rules_pdf = '1tPia7eLaoUamKl9iulbEADoI_lwo5pz_'
var payment_refund = '1A3HNJ1iH2A0G9WiwJlBZxk788f-9QdaB' // Not used yet.
var parents_note_pdf = '1_DG8n1HdldZTrc5XX1b5ESqxr7pbP9yv'
var db_folder = '1BDEV3PQULwsrqG3QjsTj1EGcnqFp6N6r'

function savePDF(blob, fileName) {
  blob = blob.setName(fileName)
  pdfId = DriveApp.createFile(blob).getId()
  DriveApp.getFileById(pdfId).moveTo(DriveApp.getFolderById(db_folder))
  return pdfId
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

function setRangeTextColor(sheet, x, y, text, color) {
  sheet.getRange(x,y).setValue(text);
  sheet.getRange(x,y).setFontColor(color);
}


function displayErrorPanel(message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(message, ui.ButtonSet.OK);
}

function getStringAt(x, y) {
  return SpreadsheetApp.getActiveSheet().getRange(x,y).getValue().toString()
}

// Validate a cell at (x, y) whose value is set via a drop-down menu. We rely
// on the fact that a cell not yet set a proper value always has the same value.
// When the value is valid, it is returned. Otherwise, '' is returned.
function validateAndReturnDropDownValue(x, y, message) {
  var value = getStringAt(x, y)
  if (value == 'Choix non renseigné' || value == '') {
    displayErrorPanel(message)
    return ''
  }
  return value
}

// Runs when the [secure authorization] button is pressed.
function GetAuthorization() {
  ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
}

// This is what the [generate and send email] button runs.
function GeneratePDFAndSendEmail() {
  // Validation: proper civility
  var civility = validateAndReturnDropDownValue(
    8, 3,
    "Vous n'avez pas renseigné de civilité")
  if (civility == '') {
    return
  }
  
  // Validation: a family name
  var family_name = getStringAt(8, 4)
  if (family_name == '') {
    displayErrorPanel("Vous n'avez pas renseigné de nom de famille ou vous avez oublié \n" +
                       "de valider le nom de famille par [return] ou [enter]...")
    return
  }

  // Validation: proper email adress.
  var mail_to = getStringAt(11, 3)
  if (mail_to == '') {
    displayErrorPanel("Vous n'avez pas saisi d'adresse email principale ou vous avez oublié \n" +
                       "de valider l'adresse email par [return] ou [enter]...")
    return
  }
  
  // Validation: parental consent set.
  var consent = validateAndReturnDropDownValue(
    81, 5,
    "Vous n'avez pas renseigné la nécessitée ou non de devoir " +
    "fournir une autorisatin parentalle.");
  if (consent == '') {
    return
  }
  
  // Update the timestamp. 
  setRangeTextColor(SpreadsheetApp.getActiveSheet(), 81, 3,
                    Utilities.formatDate(new Date(),
                                         Session.getScriptTimeZone(),
                                         "dd-MM-yyyy à HH:mm:ss"), 'black')
  SpreadsheetApp.flush()
  
  // Create the invoice as a PDF: first create a blob and then save the blob
  // as a PDF and move it to the db/ directory. Add it to the attachment array.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var blob = createPDF(spreadsheet.getUrl())
  var pdf = DriveApp.getFileById(savePDF(blob, spreadsheet.getName()+'.pdf'))
  attachments = [pdf.getAs(MimeType.PDF)]
  
  // Determine whether parental consent needs to be generated. If that's the case,
  // we generate additional attachment content.
  var rules, parents_note, parental_consent, parental_consent_text = ''
  if (consent = 'Nécessaire') {
    var consent = DriveApp.getFileById(parental_consent_pdf)
    var rules = DriveApp.getFileById(rules_pdf)
    var notes = DriveApp.getFileById(parents_note_pdf)
    attachments.push(consent.getAs(MimeType.PDF))
    attachments.push(rules.getAs(MimeType.PDF))
    attachments.push(notes.getAs(MimeType.PDF))
    
    parental_consent_text = (
      "<p>Il vous faut également compléter et signer l'autorisation " +
      "parentale fournie en attachment, couvrant le droit à l'image, le règlement " +
      "intérieur et les interventions médicales.</p>" +
      "<p>Le règlement intérieur mentionné dans l'autorisation parentale est joint à " +
      "ce message en attachement.</p>" +
      "<p>Vous trouverez également en attachement une note adressée aux parents, merçi " +
      "de la lire attentivement.</p>")
  }

  var subject = ('[Incription Ski Club Allevardin] ' +
                 civility + ' ' + family_name + ': Facture pour la saison 2020/2021')

  // Collect the personal message and add it to the mail
  var personal_message_text = getStringAt(80, 3)
  if (personal_message_text != '') {
    personal_message_text = '<p>' + personal_message_text + '</p>'
  }
  
  email_options = {
    name: 'Ski Club Allevardin, gestion des inscriptions',
    to: mail_to,
    subject: subject,
    htmlBody:
      '<h3>Bonjour ' + civility + ' ' + family_name + '</h3>' +

      '<p>Votre facture pour la saison 2020/2021 est disponible en attachement. Veuillez contrôler ' +
      'qu\'elle correspond à vos besoins.</p>' +
    
      '<p>Votre règlement est à retourner à:' +
    
      '<blockquote>' +
      '  <b>Marie-Pierre Béranger</b>,<br>' +
      '  44 Grange Merle.<br>' +
      '  <b><u>38580 Allevard</u></b><br>' +
      '</blockquote>' +
    
      '<p>Pour faciliter la gestion des inscriptions nous vous invitons à accompagner vos chèques ' +
      'd\'une copie de la facture en attachement ou de mentionner au dos de chacun de vos chèques ' +
      'la référence suivante: <b><code>' + spreadsheet.getName() + '</code></b></p>' +
    
      parental_consent_text +
      personal_message_text + 
    
      '<p>Nous vous remercions de la confiance que vous nous accordez cette année.\n' +
      'En ces temps incertains, la perception de votre règlement et le remboursement des frais ' +
      'engagés pourrons se faire sous certaines conditions. Retrouvez les en détails sur ' +
      'skicluballevard.fr/conditions-2020-2021</p>' +
    
      '<p>Des questions concernant cette facture? Répondez directement à ce mail. ' +
      'Des questions concernant la saison 2020/2021? Envoyez un mail à skicluballevardin@gmail.com</p>' +
    
      '~SCA',
    attachments: attachments
  }
   
  // Add CC if defined.
  var cc_to = getStringAt(11, 5)
  if (cc_to != '') {
    email_options.cc = cc_to
  }
  
  // FIXME: BCC inscriptions
  // email_options.bcc = 'blah'
  
  MailApp.sendEmail(email_options)
}
