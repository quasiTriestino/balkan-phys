// Confirm registration on submission of Google form
// Student confirmation followed by faculty confirmation below

function StudentRegistrationTrigger() {

    // Trigger that will happen at each submit into spreasheet
    // run function onStudentSubmit

    ScriptApp.newTrigger('onStudentSubmit')
        .forSpreadsheet('1XTYhwY1WoHGFK1ChXyh5sY746TXZ1UGcfBUlcohGGYc')
        .onFormSubmit()
        .create();
}

function onStudentSubmit(e) {

    var values = e.namedValues; // grab all values

    // email mismatch?
    var emailmismatch = '';
    var email1 = values['E-mail'].toString();
    var email2 = values['Please re-enter your email here'].toString();
    if (email1 === email2) {emailmismatch = 'No';}
    else {emailmismatch = 'Yes';}

    var hostels = values['Would you like us to make hostel reservations for you (Thursday and Friday night) in one of the two hostels listed on our website?'].toString();

    var hostelExtra = values['Would you like us to make a hostel reservation for you for Saturday night after the excursion?'].toString();
    if (hostelExtra.length < 3) {hostelExtra = 'No'};

    var hosteldays = 0;
    if (hostels === 'Yes') {hosteldays = 2};
    if (hostelExtra === 'Yes') {hosteldays = hosteldays + 1};

    var hostelmessage = '';
    if (hosteldays == 1) {hostelmessage = '(check-in Sat afternoon, check-out Sun morning)';}
    if (hosteldays == 2) {hostelmessage = '(check-in Thu afternoon, check-out Sat morning)';}
    if (hosteldays == 3) {hostelmessage = '(check-in Thu afternoon, check-out Sun morning)';}

    var saturdayevents = values['Would you like to attend one of the social events on Saturday?'].toString();

    var tara = values['Please rank your choices for each Saturday excursion. [Tara rafting ]'].toString();
    var neretva = values['Please rank your choices for each Saturday excursion. [Neretva rafting ]'].toString();
    var mostar = values['Please rank your choices for each Saturday excursion. [Mostar and Fortica Zipline ]'].toString();
    var paintball = values['Please rank your choices for each Saturday excursion. [Escape room and Paintball ]'].toString();

    var firstchoice = '';
    if (tara === '1st choice') {firstchoice = 'Tara rafting';}
    else if (neretva === '1st choice') {firstchoice = 'Neretva rafting';}
    else if (mostar === '1st choice') {firstchoice = 'Mostar ziplining';}
    else {firstchoice = 'Paintball and escape room';}

    var secondchoice = '';
    if (tara === '2nd choice') {secondchoice = 'Tara rafting';}
    else if (neretva === '2nd choice') {secondchoice = 'Neretva rafting';}
    else if (mostar === '2nd choice') {secondchoice = 'Mostar ziplining';}
    else {secondchoice = 'Paintball and escape room';}

    var thirdchoice = '';
    if (tara === '3rd choice') {thirdchoice = 'Tara rafting';}
    else if (neretva === '3rd choice') {thirdchoice = 'Neretva rafting';}
    else if (mostar === '3rd choice') {thirdchoice = 'Mostar ziplining';}
    else {thirdchoice = 'Paintball and escape room';}

    var fourthchoice = '';
    if (tara === '4th choice') {fourthchoice = 'Tara rafting';}
    else if (neretva === '4th choice') {fourthchoice = 'Neretva rafting';}
    else if (mostar === '4th choice') {fourthchoice = 'Mostar ziplining';}
    else {fourthchoice = 'Paintball and escape room';}

    var choices = '';
    choices = '<ol><li>' + firstchoice + '</li><li>' + secondchoice + '</li><li>' + thirdchoice + '</li><li>' + fourthchoice + '</li></ol>';

    var eventtransport = '';
    if (values['Will you need transportation to and from the Saturday events?'].toString() === 'Yes, I will need transportation.'){
        eventtransport = 'Yes';}
    else {eventtransport = 'No';}

    var total = 10;
    if (hostels === 'Yes') {total = total + 20};
    if (hostelExtra === 'Yes') {total = total + 10};

    var totalreg = total + 5;

    var comments = values['If you have any other questions, comments, or concerns, please write them here.'];

    // send registration email to info@balkanphys.com

    var htmlBody = '<p>';
    if (emailmismatch === 'Yes')
    {
        htmlBody += '<b><font color="red">BE AWARE:</font></b> This person entered two slightly different email addresses while registering. ';
        htmlBody += 'If an email bounces back to the BPC email account, please try sending a confimation to the second email.<br>';
        htmlBody += '<hr>';
        htmlBody += '<br>';
    }
    htmlBody += 'The following <b>student</b> just registered for BPC:' + '<br>';
    htmlBody += '<br>'
    htmlBody += "Name: " + values['First name'] + ' ' + values['Last name'] + '<br>';
    htmlBody += "School: " + values['Medical School or University'] + '<br>';
    htmlBody += "Email: " + '<a href="mailto:' + values['E-mail'] + '">' + values['E-mail'] + '</a><br>';
    if (emailmismatch === 'Yes')
    {
        htmlBody += 'Second entered email: ' + email2 + '<br>';
    }
    htmlBody += '<br>';
    if (String(comments).length > 0)
    {
        htmlBody += 'This person left the following comments while registering: <br>';
        htmlBody += '<br>';
        htmlBody += '<font color="red"><i>' + comments + '</i></font><br>';
        htmlBody += '<br>';
    }
    htmlBody += 'Registration details for this person can be found at the student registration spreadsheet.<br>';
    htmlBody += '<br>';
    htmlBody += 'Love,<br>';
    htmlBody += 'Your friendly BPC email robot';
    htmlBody += '</p>';

    GmailApp.sendEmail('info@balkanphys.com','Student Registration BPC','', {htmlBody:htmlBody});
    GmailApp.sendEmail('michael.a.chirillo@gmail.com','Student Registration BPC','', {htmlBody:htmlBody});

    // create PDF registration confirmation

    // template = 1uqcfJtP9-pnZ1XDrCJAZw-6MumXQkemwGIUhyLUhBtc
    // TEMP folder = 1q1pD4gyfF6y3ASjOT4A1wnpTF5SJIrJa
    // PDF folder = 1pLZygghTZIvw6yE0Sx-wHJahIETrUySN

    //  const docFile = DriveApp.getFileById("1uqcfJtP9-pnZ1XDrCJAZw-6MumXQkemwGIUhyLUhBtc");
    //  const tempFolder = DriveApp.getFolderById('1q1pD4gyfF6y3ASjOT4A1wnpTF5SJIrJa');
    //  const pdfFolder = DriveApp.getFolderById('1pLZygghTZIvw6yE0Sx-wHJahIETrUySN');
    //
    //  const tempFile = docFile.makeCopy(tempFolder);
    //  const tempDocFile = DocumentApp.openById(tempFile.getId());
    //  const body = tempDocFile.getBody();
    //  var currentDate = Utilities.formatDate(new Date(), "GMT", "EEEE dd MMM yyyy")
    //  body.replaceText("{date}", currentDate);
    //  body.replaceText("{first}", values['First name']);
    //  body.replaceText("{last}", values['Last name']);
    //  tempDocFile.saveAndClose();
    //  const pdfContentBlob = tempFile.getAs(MimeType.PDF);
    //  pdfFolder.createFile(pdfContentBlob).setName(values['First name']);
    //  tempFolder.removeFile(tempFile);

    // send email to registered student with payment information

    var htmlBody2 = '<p>';
    htmlBody2 += "Thank you for registering for BPC! We're looking forward to seeing you in Sarajevo. ";
    htmlBody2 += "A summary of your registration as well as payment information is listed below. Please make sure this information is correct.<br>";
    htmlBody2 += '<br>';
    htmlBody2 += "<b>REGISTRATION INFORMATION</b><br>";
    htmlBody2 += '<br>';
    htmlBody2 += values['First name'] + ' ' + values['Last name'] + ' (DOB: ' + values['Date of birth '] + ')<br>';
    htmlBody2 += values['Address'] + ', ' + values['City or Town'] + ', ' + values['Country '] + '<br>';
    htmlBody2 += values['Phone number'] + '<br>';
    htmlBody2 += values['E-mail'] + '<br>';
    htmlBody2 += '<br>';
    htmlBody2 += 'Emergency contact: ' + values['Emergency contact information (name and number)'] + '<br>';
    htmlBody2 += '<br>';
    htmlBody2 += "School: " + values['Medical School or University'] + '<br>';
    htmlBody2 += "Year: " + values['Year of study'] + '<br>';
    htmlBody2 += "Team name: " + values['Team name'] + '<br>';
    htmlBody2 += "Number of people: " + values['Number of people in your team'] + '<br>';
    htmlBody2 += "Team members: " + values['Names of people in your team'] + '<br>';
    htmlBody2 += "<br>";
    htmlBody2 += "Hostel nights requested: " + hosteldays + ' ' + hostelmessage + '<br>';
    htmlBody2 += "Attending Saturday events: " + values['Would you like to attend one of the social events on Saturday?'] + '<br>';
    if (values['Would you like to attend one of the social events on Saturday?'].toString() === 'Yes'){
        htmlBody2 += "Transportation to Saturday events: " + eventtransport + '<br>';
        htmlBody2 += "Event preferences (high to low):";
        htmlBody2 += choices;
    }
    htmlBody2 += "Request official invitation: " + values['Do you need an official letter of invitation?'] + '<br>';
    htmlBody2 += '<br>';
    if (String(comments).length > 0)
    {
        htmlBody2 += 'Additional comments / questions: <i>';
        htmlBody2 += comments + '</i><br>';
        htmlBody2 += '<br>';
    }
    htmlBody2 += "<b>Total due: " + total + ' €</b><br>';
    htmlBody2 += '<br>';
    htmlBody2 += "<b>PAYMENT INFORMATION</b><br>";
    htmlBody2 += '<br>';
    htmlBody2 += 'Your amount due of <b>' + total + ' €</b> can be transfered to the bank account of our student organization, USMF, via IBAN. Please remember this price will increase the closer we get to the competition. Here are the account details:<br>';
    htmlBody2 += '<br>';
    htmlBody2 += "<strong>Beneficiary Bank:</strong>  Raiffeisen Bank dd Bosna i Hercegovina <br>";
    htmlBody2 += "<strong>Swift Code:</strong>  RZBABA2S <br>";
    htmlBody2 += "<strong>Address:</strong>    Zmaja od Bosne bb Sarajevo BiH <br>";
    htmlBody2 += "<strong>IBAN Code:</strong>   BA391610000030570046 <br>";
    htmlBody2 += "<strong>Full Beneficiary Name:</strong>  USMF Udruzenje <br>";
    htmlBody2 += '<strong>Purpose of payment:</strong>    "Balkan Physiology Competition" <br>';
    htmlBody2 += '<br>';
    htmlBody2 += "Once your payment has been made and confirmed, you will receive a payment confirmation email. ";
    htmlBody2 += "Please contact us directly if you have requested an invitation letter. <br>";
    htmlBody2 += '<br>';
    // htmlBody2 += 'In case you or your team members would like to study some more physiology or improve your general medical skills, our partner, McGraw Hill, is giving you free access <strong>from February 8 until March 30</strong> to their comprehensive online platform for medical education <strong>AccessMedicine</strong>! Visit the <a href="https://accessmedicine.mhmedical.com/">AccessMedicine website</a> using the following username and password:<br>';
    // htmlBody2 += '<br>';
    // htmlBody2 += "<strong>username:</strong>  usarajevo <br>";
    // htmlBody2 += "<strong>password:</strong>  Medicine2020 <br>";
    // htmlBody2 += '<br>';
    htmlBody2 += 'To get updates, see example quiz questions, and connect with other competitors, <a href="https://www.facebook.com/balkanphysiologycompetition">visit and like our Facebook page</a>!<br>';
    htmlBody2 += '<br>';
    htmlBody2 += "For any other questions or comments, please do not hesitate to email us by replying to this message. Finally, please encourage your teammates to register soon! Thank you again and we're looking forward to meeting you!<br>";
    htmlBody2 += '<br>';
    htmlBody2 += '<b>Team BPC</b><br>';
    htmlBody2 += '<a href="mailto: info@balkanphys.com">info@balkanphys.com</a><br>';
    htmlBody2 += '</p>';

    GmailApp.sendEmail(values['E-mail'],'BPC Registration confirmation','', {htmlBody:htmlBody2});
    // GmailApp.sendEmail('michael.a.chirillo@gmail.com','BPC Registration confirmation','', {htmlBody:htmlBody2});

    // update payments spreadsheet to include new data
    const id = '1gDyGVxar6aohSzbpdOpbyP-Keum3rrNK0ByoI4vV6Vk';
    const spreadsheet = SpreadsheetApp.openById(id);
    const payments_sheet = spreadsheet.getSheetByName('Payments');

    payments_sheet.getRange(payments_sheet.getLastRow() + 1, 1, 1, payments_sheet.getLastColumn())
        .setValues([[values['Last name'], values['First name'], values['Country '], values['Medical School or University'], values['E-mail'], values['Phone number'], values['Informazioni cronologiche'], values['Team name'], hosteldays, total, totalreg, 'No', 'No']]);
}

function FacultyRegistrationTrigger() {
    ScriptApp.newTrigger('onFacultySubmit')
	.forSpreadsheet('1PZaDsgRc0Sa7B2QV0FHDM1_Vw7VadCQWA8QVY1DvWTY')
	.onFormSubmit()
	.create();
}

function onFacultySubmit(e) {
    // send notification email to info@balkanphys.com

    var values = e.namedValues;

    // email mismatch?
    var emailmismatch = '';
    var email1 = values['E-mail'].toString();
    var email2 = values['Please re-enter your email here'].toString();
    if (email1 === email2) {emailmismatch = 'No';}
    else {emailmismatch = 'Yes';}

    var saturdayevents = values['Would you like to attend one of the social events on Saturday?'].toString();

    var tara = values['Please rank your choices from 1 to 4 for each Saturday excursion. [Tara rafting ]'].toString();
    var neretva = values['Please rank your choices from 1 to 4 for each Saturday excursion. [Neretva rafting ]'].toString();
    var mostar = values['Please rank your choices from 1 to 4 for each Saturday excursion. [Mostar and Fortica Zipline ]'].toString();
    var paintball = values['Please rank your choices from 1 to 4 for each Saturday excursion. [Escape room and Paintball ]'].toString();

    var firstchoice = '';
    if (tara === '1st choice') {firstchoice = 'Tara rafting';}
    else if (neretva === '1st choice') {firstchoice = 'Neretva rafting';}
    else if (mostar === '1st choice') {firstchoice = 'Mostar ziplining';}
    else {firstchoice = 'Paintball and escape room';}

    var secondchoice = '';
    if (tara === '2nd choice') {secondchoice = 'Tara rafting';}
    else if (neretva === '2nd choice') {secondchoice = 'Neretva rafting';}
    else if (mostar === '2nd choice') {secondchoice = 'Mostar ziplining';}
    else {secondchoice = 'Paintball and escape room';}

    var thirdchoice = '';
    if (tara === '3rd choice') {thirdchoice = 'Tara rafting';}
    else if (neretva === '3rd choice') {thirdchoice = 'Neretva rafting';}
    else if (mostar === '3rd choice') {thirdchoice = 'Mostar ziplining';}
    else {thirdchoice = 'Paintball and escape room';}

    var fourthchoice = '';
    if (tara === '4th choice') {fourthchoice = 'Tara rafting';}
    else if (neretva === '4th choice') {fourthchoice = 'Neretva rafting';}
    else if (mostar === '4th choice') {fourthchoice = 'Mostar ziplining';}
    else {fourthchoice = 'Paintball and escape room';}

    var choices = '';
    choices = '<ol><li>' + firstchoice + '</li><li>' + secondchoice + '</li><li>' + thirdchoice + '</li><li>' + fourthchoice + '</li></ol>';

    var eventtransport = '';
    if (values['Will you need transportation to and from the Saturday events?'].toString() === 'Yes, I will need transportation.'){
	eventtransport = 'Yes';}
    else {eventtransport = 'No';}

    var comments = values['Do you have any other questions, comments or concerns?'];

    var htmlBody = '<p>';
    if (emailmismatch === 'Yes')
    {
	htmlBody += '<b><font color="red">BE AWARE.</font></b> This person entered two slightly different email addresses while registering. ';
	htmlBody += 'If an email bounces back to the BPC email account, please try sending a confimation to the second email.<br>';
	htmlBody += '<hr>';
	htmlBody += '<br>';
    }
    htmlBody += 'The following <b>faculty member</b> has just registered for BPC:' + '<br>';
    htmlBody += '<br>'
    htmlBody += "Name: " + values['First name'] + ' ' + values['Last name'] + '<br>';
    htmlBody += "School: " + values['Medical School or University'] + '<br>';
    htmlBody += "Academic position: " + values['Please enter your academic / faculty title.'] + '<br>';
    htmlBody += "Email: " + '<a href="mailto:' + values['E-mail'] + '">' + values['E-mail'] + '</a><br>';
    if (emailmismatch === 'Yes')
    {
	htmlBody += 'Second entered email: ' + email2 + '<br>';
    }
    htmlBody += '<br>';
    if (String(comments).length > 0)
    {
	htmlBody += 'This person left the following comments while registering: <br>';
	htmlBody += '<br>';
	htmlBody += '<font color="red"><i>' + comments + '</i></font><br>';
	htmlBody += '<br>';
    }
    htmlBody += 'Registration details for this person can be found at the faculty registration spreadsheet.<br>';
    htmlBody += '<br>';
    htmlBody += 'Love,<br>';
    htmlBody += 'Your friendly BPC email robot';
    htmlBody += '</p>';
    GmailApp.sendEmail('info@balkanphys.com','Faculty Registration BPC','', {htmlBody:htmlBody});
    GmailApp.sendEmail('michael.a.chirillo@gmail.com','Faculty Registration BPC','', {htmlBody:htmlBody});

    // send confirmation email to faculty member

    var htmlBody2 = '<p>';
    htmlBody2 += "Thank you for registering for BPC! We're looking forward to seeing you in Sarajevo. ";
    htmlBody2 += "Please contact us with specifics if you need a formal invitation letters. "
    htmlBody2 += "We will respond to questions left in the registration form as quickly as possible. For any other questions or comments, please do not hesitate to email us by replying to this message. <br>";
    htmlBody2 += '<br>';
    htmlBody2 += "A short summary of your registration is found below this message. Please be on the look out for reminder emails in the summer.<br>";
    htmlBody2 += '<br>';
    // htmlBody2 += 'In case you or your teams would like to study some more physiology or improve your general medical skills, our partner, McGraw Hill, is giving you free access <strong>from February 8 until March 30</strong> to their comprehensive online platform for medical education <strong>AccessMedicine</strong>! Visit the <a href="https://accessmedicine.mhmedical.com/">AccessMedicine website</a> using the following username and password:<br>';
    // htmlBody2 += '<br>';
    // htmlBody2 += "<strong>username:</strong>  usarajevo <br>";
    // htmlBody2 += "<strong>password:</strong>  Medicine2020 <br>";
    // htmlBody2 += '<br>';
    htmlBody2 += 'Thank you again!<br>';
    htmlBody2 += '<br>';
    htmlBody2 += '<b>Team BPC</b><br>';
    htmlBody2 += '<a href="mailto: info@balkanphys.com">info@balkanphys.com</a><br>';
    htmlBody2 += '<hr>';
    htmlBody2 += '<br>';
    htmlBody2 += "<b>REGISTRATION SUMMARY</b><br>";
    htmlBody2 += '<br>';
    htmlBody2 += values['First name'] + ' ' + values['Last name'] + '<br>';
    htmlBody2 += values['Address'] + ', ' + values['City or Town'] + ', ' + values['Country'] + '<br>';
    htmlBody2 += values['Phone number'] + '<br>';
    htmlBody2 += values['E-mail'] + '<br>';
    htmlBody2 += '<br>';
    htmlBody2 += "School: " + values['Medical School or University'] + '<br>';
    htmlBody2 += "Attending Saturday events: " + saturdayevents + '<br>';
    if (saturdayevents === 'Yes'){
	htmlBody2 += "Transportation to Saturday events: " + eventtransport + '<br>';
	htmlBody2 += "Event preferences (high to low):";
	htmlBody2 += choices;
    }
    htmlBody2 += "Request official invitation: " + values['Do you need an official letter of invitation?'] + '<br>';
    if (String(comments).length > 0)
    {
	htmlBody2 += '<br>';
	htmlBody2 += 'Additional comments / questions: <i>';
	htmlBody2 += comments + '</i><br>';
    }
    htmlBody2 += '</p>';
    GmailApp.sendEmail(values['E-mail'],'BPC Registration Confirmation','', {htmlBody:htmlBody2});
}
