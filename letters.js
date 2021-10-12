// Email RSVP confirmations and mail PDFs for faculties
// Create and email PDFs

// PDF templates can be placed anywhere
// Search and replace specified below

function RSVP_each_row() {
    my_rows = SpreadsheetApp.getActiveSheet().getActiveRange().getNumRows()
    my_first_row = SpreadsheetApp.getActiveSheet().getCurrentCell().getRow()
    interations = my_first_row + my_rows

    for (let i = my_first_row; i < interations; i++) {
	SpreadsheetApp.getActiveSheet().setActiveSelection("A".concat(i.toString()))
	emailRSVPconfirmation()
    }
}

function createPDF(first, last, school, dob) {

    const pdfFolder = DriveApp.getFolderById("1jTK5e8N7NdRSxW-vdnmF2Nh5fofuPlwe")
    const tempFolder = DriveApp.getFolderById("14YJD5PjU6_FUpGey8ulkgDBCELqEDsyX")
    const templateDoc = DriveApp.getFileById("1IrPF78v6jp_6cD-CsU4D6FFsfnq7vBbqhHqAGyz9zS4")

    const newTempFile = templateDoc.makeCopy(tempFolder)

    const openDoc = DocumentApp.openById(newTempFile.getId())
    const body = openDoc.getBody()
    body.replaceText("{fn}", first)
    body.replaceText("{ln}", last)
    body.replaceText("{mf}", school)
    body.replaceText("{dob}", dob)
    openDoc.saveAndClose()

    const blobPDF = newTempFile.getAs(MimeType.PDF)
    const pdfFile = pdfFolder.createFile(blobPDF).setName(last + "-" + first + "-BPC")
    //tempFolder.removeFile(newTempFile)

    return pdfFile

}

function emailRSVPconfirmation() {
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1")
    var rowIndex = spreadsheet.getCurrentCell().getRow()
    var data = spreadsheet.getRange(rowIndex, 1, 1, 13).getValues()[0]
    
    last = data[0]
    first = data[1]
    team = data[3]
    email = data[4]
    hostel = data[8]
    social = data[11]
    paid = data[13]

    payStatus = "Paid"
    if(paid === "No"){payStatus = "<span style=\"color: #ff0000\"><strong>Not paid</strong></span>"}

    if(social === "N/A"){social = "None"}

    console.log(data)

    // Grab raw data spreadsheet
    const spreadsheet_2 = SpreadsheetApp.openById('1XTYhwY1WoHGFK1ChXyh5sY746TXZ1UGcfBUlcohGGYc');
    const student_sheet = spreadsheet_2.getSheetByName('Risposte del modulo 1');
    var raw_data = student_sheet.getDataRange().getValues();
    var RSVP_raw = raw_data.filter(function(row){
	if (row[2] === last && row[1] === first) {
	    return row;
	}
    });

    console.log(RSVP_raw)

    school = RSVP_raw[0][7]
    dob = Utilities.formatDate(RSVP_raw[0][3], "GMT+2", "dd/MM/yyyy")

    pdfFile = createPDF(first, last, school, dob)

    // Set up email text

    var htmlBody = '<p>';
    htmlBody += "Dear " + first + ",<br>";
    htmlBody += '<br>';
    htmlBody += "This is confirmation that you have RSVPd to BPC. We are excited to see you in Sarajevo! Please read the notes below and make sure the information we have about you at the bottom of this email is correct.<br>";
    htmlBody += '<br>';

    htmlBody += "<strong>Faculty letters:</strong> You will find a PDF attached to this email that you can give to your faculty if you need to provide confirmation of your participation in BPC.<br>"
    htmlBody += '<br>';

    htmlBody += '<strong>Payments:</strong> Below you will find your payment status. If your faculty is covering your fees and they have not yet paid, your payment status will reflect that. Once we have confirmed your payment, you will get a separate payment confirmation email. Payments can be made through bank transfer and information about that is <a href="https://balkanphys.com/registration#fees">on our website</a>. We have extended the early payment deadline to Oct 1. After this fees will go up.<br>';
    htmlBody += '<br>';

    htmlBody += "<strong>COVID:</strong> Remember, you must provide either proof of recovery, full vaccination, or testing to participate in BPC. <strong>Please bring this proof with you to Sarajevo.</strong> Please also make sure to check the entry requirements into BiH from your country. Currently, travel from the EU requires proof of vaccination or testing done within 48 hours of arriving. This can, however, change. We are not responsible for BiH travel restrictions, so please make sure you understand what those restrictions are before traveling to Sarajevo.<br>"
    htmlBody += '<br>';

    htmlBody += "<strong>Hostels:</strong> If you requested hostel nights, you will be provided with hostel information next week. Please check the number of nights you requested below. <br>"
    htmlBody += '<br>';

    htmlBody += "<strong>Tara:</strong> Not enough people requested rafting on Tara (!!!), so if you did rank Tara first, we currently have your second choice listed as your Saturday social event. If you would like to change this, please let us know.<br>"
    htmlBody += '<br>';

    htmlBody += '<strong>WhatsApp group:</strong> Please join our WhatsApp group for BPC information! <a href="https://chat.whatsapp.com/DSBNsCpvGSo5ihbgiKzvZk">BPC WhatsApp Invitation</a><br>'
    htmlBody += '<br>';

    htmlBody += "<strong>Competition:</strong> Finally, the written portion of the competition will be held online Friday morning. Please make sure to have access to a computer or smartphone while you are in Sarajevo. The oral quizzes will be held live at the medical faculty. Before the competition, you will be provided with usernames and passwords that you will use to access the written quiz. Please see our website for the quiz agenda/schedule.<br>"
    htmlBody += '<br>';

    htmlBody += "<strong>REGISTRATION INFORMATION</strong><br>"
    htmlBody += '<br>';

    htmlBody += "<strong>Participant:</strong> " + first + " " + last + '<br>';
    htmlBody += "<strong>DOB:</strong> " + dob + "<br>";
    htmlBody += "<strong>Institution:</strong> " + school + "<br>";
    htmlBody += "<strong>Team:</strong> " + team + '<br>';
    htmlBody += "<strong>Saturday social event:</strong> " + social + '<br>';
    htmlBody += "<strong>Number of nights at hostel requseted:</strong> " + hostel + '<br>';
    htmlBody += "<strong>Payment status:</strong> " + payStatus + '<br>';
    htmlBody += '<br>';

    htmlBody += "Again, thank you so much and let us know if you have questions.<br>";
    htmlBody += '<br>';
    htmlBody += '<b>Team BPC</b><br>';
    htmlBody += '<a href="mailto: info@balkanphys.com">info@balkanphys.com</a><br>';
    htmlBody += '</p>';

    MailApp.sendEmail({
        to: email,
        subject: "BPC21 Confirmation and faculty letters",
        htmlBody: htmlBody,
        attachments: [pdfFile],
        name: "Team BPC"
    });

}
