// Add options to Google spreadsheet meny on open

function onOpen() {
    // Name the menu item
    var menu = SpreadsheetApp.getUi().createMenu("BPC functions");

    // Provide option and its function
    menu.addItem("Collect BPC data" , "Collect_data");

    // Add menu to user interface
    menu.addToUi();
}

function Collect_data() {
  RSVP_each_row();
}

function getName() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getCurrentCell().getRow();
  var rowValues = sheet.getRange(rowIndex, 1, 1, 2).getValues()[0];
  return(rowValues);
}
