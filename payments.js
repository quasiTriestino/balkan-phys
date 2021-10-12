// Send payment confirmation
// Check for "Yes" in paid column and "No" in conf column
// This script loops through every line in the file

function sendStudentConfirmations() {
    const spreadsheet = SpreadsheetApp.getActive();
    const payments_sheet = SpreadsheetApp.getActiveSheet();

    const lastrow = payments_sheet.getLastRow();
    const endloop = lastrow + 1;

    var i;
    for (i = 2; i < endloop; i++) {

	var range = payments_sheet.getRange(i, 1, 1, 13);
	var all_data = range.getValues(); // grab a row of data (as 2D array)

	if (all_data[0][11] === 'Yes' && all_data[0][12] === 'No') // if the participant paid and not received confirmation
	{
	    // send then a payment confirmation email
	    var htmlBody = '<p>';
	    htmlBody += "Dear " + all_data[0][1] + ",<br>";
	    htmlBody += '<br>';
	    htmlBody += "This is confirmation that your amount due of <b>" + all_data[0][9] + " €</b> for BPC 2021 was received. Please keep this email as confirmation that you've paid. Thank you!<br>";
	    htmlBody += '<br>';
	    htmlBody += 'To get updates, see example quiz questions, and connect with other competitors, <a href="https://www.facebook.com/balkanphysiologycompetition">visit and like our Facebook page</a>.<br>';
	    htmlBody += '<br>';
	    htmlBody += "Again, thank you so much and we're looking forward to seeing you in Sarajevo!<br>";
	    htmlBody += '<br>';
	    htmlBody += '<b>Team BPC</b><br>';
	    htmlBody += '<a href="mailto: info@balkanphys.com">info@balkanphys.com</a><br>';
	    htmlBody += '</p>';

	    GmailApp.sendEmail(all_data[0][4],'BPC Payment Confirmation','', {htmlBody:htmlBody});

	    // and update spreadsheet

	    all_data[0][12] = 'Conf';
	    range.setValues(all_data);
	} // close if

    } // close for lopp

}


// Add a menu item to Google spreasheet

function onOpen() {
   var menu = SpreadsheetApp.getUi().createMenu("BPC functions");
   menu.addItem("Send confirmations" , "sendStudentConfirmations");
   menu.addToUi();
}
