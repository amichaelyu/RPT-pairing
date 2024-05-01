/**
 * 
 * Removes tutees who requested a tutor from the automated system by filling out their paired square in the sheets
 * Auto sends email 
 * 
 */

var sendEmail = true;

function email() {
    var form = FormApp.getActiveForm();
    var response = form.getResponses()[form.getResponses().length-1];
    var tuteeSheet = SpreadsheetApp.openById("").getSheets()[2];
    var webhook = "";
    var row = 2;
    while (tuteeSheet.getRange(row,1).getCell(1,1).getValue() != '') {
      if (tuteeSheet.getRange(row,2).getCell(1,1).getValue() == response.getRespondentEmail()) {
        break;
      }
      row++;
    }
    if (tuteeSheet.getRange(row,1).getCell(1,1).getValue() == '') {
      throw "Not a tutee? Don't know how this happened! Email: " + response.getRespondentEmail();
    }
    
    if (tuteeSheet.getRange(row,18).getCell(1,1).getValue() != '') {
      var rangeTuteeMath = tuteeSheet.getRange(row,12).getCell(1,1);
      var rangeTuteeScience = tuteeSheet.getRange(row,13).getCell(1,1);
      var rangeTuteeHistory = tuteeSheet.getRange(row,14).getCell(1,1);
      var rangeTuteeLanguage = tuteeSheet.getRange(row,15).getCell(1,1);
      var rangeTuteeECs = tuteeSheet.getRange(row,16).getCell(1,1);
      var rangeTuteeSubjectPaired = tuteeSheet.getRange(row,21).getCell(1,1);
      
      var allTuteeSubjects = [];
      allTuteeSubjects = allTuteeSubjects.concat(rangeTuteeMath.getValue().toString().split(", "), rangeTuteeScience.getValue().toString().split(", "), rangeTuteeHistory.getValue().toString().split(", "), rangeTuteeLanguage.getValue().toString().split(", "), rangeTuteeECs.getValue().toString().split(", "));
      allTuteeSubjects = allTuteeSubjects.filter(n => n);
      rangeTuteeSubjectPaired.setValue(allTuteeSubjects.join(','));
            var payload = {
        "request": "TUTOR PAIRING REQUEST:\n" + response.getItemResponses()[1].getResponse().toString() + " (" + response.getRespondentEmail() + ") requested to be paired with " + tuteeSheet.getRange(row,18).getCell(1,1).getValue()
      }
      var options = {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : JSON.stringify(payload)
      };
      UrlFetchApp.fetch(webhook, options);
    }
    
    var parentEmail = response.getItemResponses()[0].getResponse().toString();
    if (false) {
      MailApp.sendEmail(
      "itsmichaelyu@gmail.com",
      "Welcome to Rochester Peer Tutoring!",
      "Dear " + response.getItemResponses()[1].getResponse().toString() + ", \n\Welcome to Rochester Peer Tutoring! Anything can go here! Thank you for your interest in our organization. We have received your application and are in the process of reviewing it. Please note that our program is first come first serve and we have different numbers of tutors available for each subject, so finding a suitable tutor match may take anywhere from a couple days to a couple weeks. Upon being paired, you will receive an email from us detailing the tutoring process and providing more information about your tutor. In the meantime, feel free to reach out to us with any questions or concerns.\n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please reply to this email with any problems or concerns!", {
          // cc: parentEmail,
      });
    }
}