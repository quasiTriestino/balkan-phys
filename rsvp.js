// Code the generates RSVP from each row of a spreasheets
// Data is sourced from a payments spreadsheet
// and from the raw responses

function RSVP_each_row() {

    // Lopps through each row of active selection
    // Can only select and drag downward
    
    my_rows = SpreadsheetApp.getActiveSheet().getActiveRange().getNumRows()
    my_first_row = SpreadsheetApp.getActiveSheet().getCurrentCell().getRow()
    interations = my_first_row + my_rows
    console.log(my_rows, my_first_row)

    for (let i = my_first_row; i < interations; i++) {
	SpreadsheetApp.getActiveSheet().setActiveSelection("A".concat(i.toString()))
	createRVSP()
  }
}

function createRVSP() {
  
  // Current row index
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getCurrentCell().getRow();

  // Grab info from registration ("payments") spreadsheet
  const spreadsheet = SpreadsheetApp.openById('1gDyGVxar6aohSzbpdOpbyP-Keum3rrNK0ByoI4vV6Vk');
  const payments_sheet = spreadsheet.getSheetByName('Payments');
  var all_data = payments_sheet.getDataRange().getValues();

  // Grab info from raw student data spreadshet 
  const spreadsheet_2 = SpreadsheetApp.openById('1XTYhwY1WoHGFK1ChXyh5sY746TXZ1UGcfBUlcohGGYc');
  const student_sheet = spreadsheet_2.getSheetByName('Risposte del modulo 1');
  var raw_data = student_sheet.getDataRange().getValues(); 

  // grab name of selected participant 
  var name = getName();
  console.log(name);

  // Grab data from payment sheet for this person
  var RSVP_data = all_data.filter(function(row){
    if (row[0] === name[0] && row[1] === name[1]) {
    return row;
    }
  });

  var email = RSVP_data[0][4]
  var country = RSVP_data[0][2]
  var school = RSVP_data[0][3]
  var team = RSVP_data[0][7]
  var paid = RSVP_data[0][11]

  // Grab data from raw data sheet for this person
  var RSVP_raw = raw_data.filter(function(row){
    if (row[2] === name[0] && row[1] === name[1]) {
    return row;
    }
  });

  console.log(RSVP_raw)

  // Get hostel and nights information
  var hostel = RSVP_raw[0][16]
  var nights;
  if (hostel == "No"){nights = 0}
  else if (hostel == "Yes"){nights = 2}
  if (RSVP_raw[0][23] === "Yes"){nights = nights + 1}

// Get Social event choice #1 and #2
// Figure out where they're going 
tara = RSVP_raw[0][18]
neretva = RSVP_raw[0][19]
mostar = RSVP_raw[0][21]

var firstchoice = '';
if (tara === '1st choice') {firstchoice = 'Tara';}
else if (neretva === '1st choice') {firstchoice = 'Neretva';}
else if (mostar === '1st choice') {firstchoice = 'Mostar';}
else {firstchoice = 'Paintball/Escape';}

var secondchoice = '';
if (tara === '2nd choice') {secondchoice = 'Tara';}
else if (neretva === '2nd choice') {secondchoice = 'Neretva';}
else if (mostar === '2nd choice') {secondchoice = 'Mostar';}
else {secondchoice = 'Paintball/Escape';}

var thirdchoice = '';
if (tara === '3rd choice') {thirdchoice = 'Tara';}
else if (neretva === '3rd choice') {thirdchoice = 'Neretva';}
else if (mostar === '3rd choice') {thirdchoice = 'Mostar';}
else {thirdchoice = 'Paintball/Escape';}

var our_choice = firstchoice;
if (firstchoice == "Tara"){our_choice = secondchoice}; 

var rain = firstchoice
if (firstchoice == "Tara" || firstchoice == "Neretva"){
  if (secondchoice != "Tara" && secondchoice != "Neretva"){
    rain = secondchoice
  } else {
    rain = thirdchoice
  }
};

console.log(firstchoice)
console.log(secondchoice)
console.log(our_choice)
console.log(rain)

if (RSVP_raw[0][17] === "No"){
  firstchoice = "N/A"
  secondchoice = "N/A"
  our_choice = "N/A"
  rain = "N/A"
}

// Set values in RSVP spreasheet
sheet.getRange(rowIndex,4,1,11)
  .setValues([[team, email, country, school, hostel, nights, firstchoice, secondchoice, our_choice, rain, paid]]);

}

