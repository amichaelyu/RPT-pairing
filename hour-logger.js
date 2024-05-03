function hours() {
  var form = FormApp.getActiveForm();
  var len = form.getResponses().length;
  var tutorSheet = SpreadsheetApp.openById("").getSheets()[1];
  
  // add hours calculator
  var hoursMASTERSheet = SpreadsheetApp.openById("").getSheets()[3];
  hoursMASTERSheet.getRange(len+1,8).getCell(1,1).setValue("=F"+(len+1)+"-E"+(len+1));
  hoursMASTERSheet.getRange(len+1,8).getCell(1,1).setNumberFormat("HH:MM");

  // find tutor's hours sheet index
  var row = 2;
  var response = form.getResponses()[len-1];
  while (tutorSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    if (tutorSheet.getRange(row,2).getCell(1,1).getValue() == response.getRespondentEmail()) {
      break;
    }
    row++;
  }
  if (tutorSheet.getRange(row,1).getCell(1,1).getValue() == '') {
    throw "Not a registered tutor! Email: " + response.getRespondentEmail();
  }
  var sheetNum = tutorSheet.getRange(row,24).getCell(1,1).getValue();
  var hoursSheet = SpreadsheetApp.openById("").getSheets()[sheetNum-1];
  
  // find which form response is theirs
  var row = 2;
  var response = form.getResponses()[len-1];
  while (hoursSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    row++;
  }

  var timestamp = hoursSheet.getRange(row,1).getCell(1,1);
  var email = hoursSheet.getRange(row,2).getCell(1,1);
  var tutee = hoursSheet.getRange(row,3).getCell(1,1);
  var date = hoursSheet.getRange(row,4).getCell(1,1);
  var start = hoursSheet.getRange(row,5).getCell(1,1);
  var end = hoursSheet.getRange(row,6).getCell(1,1);
  var hours = hoursSheet.getRange(row,7).getCell(1,1);

  timestamp.setNumberFormat("MM/DD/YYYY HH:MM:SS");
  date.setNumberFormat("MM/DD/YYYY");
  start.setNumberFormat("HH:MM AM/PM");
  end.setNumberFormat("HH:MM AM/PM");
  hours.setNumberFormat("[HH]:MM");
  timestamp.setValue(hoursMASTERSheet.getRange(len+1,1).getCell(1,1).getValue());
  email.setValue(hoursMASTERSheet.getRange(len+1,2).getCell(1,1).getValue());
  tutee.setValue(hoursMASTERSheet.getRange(len+1,3).getCell(1,1).getValue());
  date.setValue(hoursMASTERSheet.getRange(len+1,4).getCell(1,1).getValue());
  start.setValue(hoursMASTERSheet.getRange(len+1,5).getCell(1,1).getValue());
  end.setValue(hoursMASTERSheet.getRange(len+1,6).getCell(1,1).getValue());
  hours.setValue("=F"+row+"-E"+row);
}
