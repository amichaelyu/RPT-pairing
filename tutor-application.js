var sendEmail = true;

function submitting() {
  var form = FormApp.getActiveForm();
  var response = form.getResponses()[form.getResponses().length-1];
  var sheet = SpreadsheetApp.openById("");
  var tutorSheet = SpreadsheetApp.openById("").getSheets()[1];
  var webhookPairing = "";
  var webhookError = "";
  var name = response.getItemResponses()[0].getResponse().toString();
  var pastTutor = response.getItemResponses()[1].getResponse().toString();
  var school = response.getItemResponses()[2].getResponse().toString();
  var phone = response.getItemResponses()[3].getResponse().toString();
  var sheetName = name + "_" + school;
  var newSheet = sheet.insertSheet()
  try {
    newSheet.setName(sheetName).setFrozenRows(1);
    sheet.setActiveSheet(sheet.getSheetByName(sheetName));
  }
  catch {
    try {
      newSheet.setName(sheetName+"_"+phone).setFrozenRows(1);
      sheet.setActiveSheet(sheet.getSheetByName(sheetName+"_"+phone));
    }
    catch {
      sheet.deleteSheet(newSheet);
      var payload = {
        "error": "We have a clone! Email: " + response.getRespondentEmail()
      }
      var options = {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : JSON.stringify(payload)
      };
      UrlFetchApp.fetch(webhookError, options);
      throw "Tutor Application\nWe have a clone! Email: " + response.getRespondentEmail();
    }
  }
  var row = 2;
  while (tutorSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    if (tutorSheet.getRange(row,2).getCell(1,1).getValue() == response.getRespondentEmail()) {
      break;
    }
    row++;
  }
  if (tutorSheet.getRange(row,1).getCell(1,1).getValue() == '') {
    var payload = {
      "error": "Tutor Application\nNot a registered tutor! Email: " + response.getRespondentEmail()
    }
    var options = {
      "method" : "post",
      "contentType" : "application/json",
      "payload" : JSON.stringify(payload)
    };
    UrlFetchApp.fetch(webhookError, options);
    throw "Not a registered tutor! Email: " + response.getRespondentEmail();
  }
  tutorSheet.getRange(row,24).getCell(1,1).setValue(sheet.getSheets().length);
  sheet.moveActiveSheet(sheet.getSheets().length);
  var row1 = newSheet.getRange('A1:H2');
  row1.getCell(1,1).setValue("Timestamp");
  row1.getCell(1,2).setValue("Email Address");
  row1.getCell(1,3).setValue("Who are you tutoring?");
  row1.getCell(1,4).setValue("Date");
  row1.getCell(1,5).setValue("Start Time");
  row1.getCell(1,6).setValue("End Time");
  row1.getCell(1,7).setValue("Hours");
  row1.getCell(1,8).setValue("Total Hours");
  row1.getCell(2,8).setValue("=SUM(G:G)");
  newSheet.getRange('A:A').setNumberFormat("MM/DD/YYYY HH:MM:SS");
  newSheet.getRange('D:D').setNumberFormat("MM/DD/YYYY");
  newSheet.getRange('E:E').setNumberFormat("HH:MM AM/PM");
  newSheet.getRange('F:F').setNumberFormat("HH:MM AM/PM");
  newSheet.getRange('G:G').setNumberFormat("[HH]:MM");
  row1.getCell(2,8).setNumberFormat("[HH]:MM");

  if (tutorSheet.getRange(row,21).getCell(1,1).getValue() != '') {
    var payload = {
      "request": "A tutor requested to be paired with a specific tutee!\n" + response.getItemResponses()[0].getResponse().toString() + " (" + response.getRespondentEmail() + ") requested to be paired with " + tutorSheet.getRange(row,21).getCell(1,1).getValue()
    }
    var options = {
      "method" : "post",
      "contentType" : "application/json",
      "payload" : JSON.stringify(payload)
    };
    UrlFetchApp.fetch(webhookPairing, options);
  }

  if (sendEmail) {
    if (pastTutor == "Yes, I tutored with RPT last school year.") {
      MailApp.sendEmail(
        response.getRespondentEmail(),
        "Welcome Back to Rochester Peer Tutoring!",
        "Dear " + response.getItemResponses()[0].getResponse().toString() + ", \n\nThank you for your interest in the organization - we are excited to have you back! We have received your application and are in the process of reviewing it. Before being paired up, you will be expected to attend a \"Returning Tutor Workshop\", where we will be sharing new updates for the upcoming school year. \n\nPlease join our Slack where a majority of our communication will be this year: https://join.slack.com/t/rochesterpeertutoring/shared_invite/zt-22csb23uv-wNOf6LTQVQOdpNfloQXhMw\n\nWe will have more details on Slack in the next couple days regarding the time and date of this workshop. In the meantime, please let us know if you have any questions or concerns - we look forward to having you back on our team!\n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please contact us if you encounter any problems!"
      );
    }
    else {
      MailApp.sendEmail(
          response.getRespondentEmail(),
          "Welcome to Rochester Peer Tutoring!",
          "Dear " + response.getItemResponses()[0].getResponse().toString() + ", \n\nWelcome to Rochester Peer Tutoring! Thank you for your interest in our organization. We have received your application and are in the process of reviewing it.\n\nPlease join our Slack where a majority of our communication will be this year: https://join.slack.com/t/rochesterpeertutoring/shared_invite/zt-22csb23uv-wNOf6LTQVQOdpNfloQXhMw\n\nIn the next couple days, we will reach out on Slack requesting an interview with you so that we can get to know you and your tutoring style better. Afterwards, you will be expected to attend a training workshop so that we can let you know how our tutoring process works and answer any questions you may have. Please let us know if you have any questions or concerns - we look forward to having you on our team!\n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please contact us if you encounter any problems!"
      );
    }
  }
}
