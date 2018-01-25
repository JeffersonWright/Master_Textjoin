function ResetSheet() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
  sheet.getRange('A:D').clearContent(); //clears the content in columns A-D
  sheet.getRange('A:D').clearFormat(); //clears the formatting in columns A-D
  sheet.getRange('F4').setValue(0); //resets the skip line field to 0, not skipping any lines
  sheet.getRange('F1:F2').clearContent(); //clears the content in the remove, replace, and starts with fields
  sheet.getRange('H1:H2').clearContent();
  sheet.getRange('J1:J2').clearContent();
  sheet.getRange('L1:L2').clearContent();
  sheet.getRange('N3').clearContent();
}
