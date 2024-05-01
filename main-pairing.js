var sendEmail = true;


/**
 * GOOGLE SHEET MUST BE IN THIS ORDER!!!!
 * 
 * PAIRED SHEET               0
 * TUTOR APPLICATIONS         1
 * TUTEE APPLICATIONS         2
 * HOUR LOGGER                3
 * TUTOR HOUR PAGES           4
 * ...
 * 
 */
function autoPairer() {
  var form = FormApp.getActiveForm();
  // take the id which is the part after the "/d"
  // https://docs.google.com/spreadsheets/d/Nu7OpaMgqbDyO62L7P3K1r8kpQ7L9L4rXowgKJ8hCyCv
  var sheet = SpreadsheetApp.openById("").getSheets()[2];

  // console.log(form.getTitle()); // check that we have the right form
  // console.log(sheet.getName()); // check that we have the right sheet

  // grabbing sheets data
  var priorityTutees = [];
  var tutees = [];
  var row = 2;
  // checks row to see if there is data
  while (sheet.getRange(row,1).getCell(1,1).getValue() != '') {
    var tuteeFullProfile = [];
    var tuteeSubjects = [];

    var submissionDate = new Date(sheet.getRange(row,1).getCell(1,1).getValue());
    
    var rangeGrade = sheet.getRange(row,7).getCell(1,1);
    var rangeVirtual = sheet.getRange(row,10).getCell(1,1);
    var rangeTime = sheet.getRange(row,11).getCell(1,1);

    var rangeMath = sheet.getRange(row,12).getCell(1,1);
    var rangeScience = sheet.getRange(row,13).getCell(1,1);
    var rangeHistory = sheet.getRange(row,14).getCell(1,1);
    var rangeLanguage = sheet.getRange(row,15).getCell(1,1);
    var rangeECs = sheet.getRange(row,16).getCell(1,1);

    var rangePaired = sheet.getRange(row,21).getCell(1,1);
    var cellArr = rangePaired.getValue().split(',');

    tuteeFullProfile.push(row.toString());

    tuteeFullProfile.push(rangeGrade.getValue().toString());

    var tuteeLocation = [];
    if (rangeVirtual.getValue().toString().indexOf("Virtual Tutoring via Zoom") != -1) {
      tuteeLocation.push("Virtual");
    } 
    if (rangeVirtual.getValue().toString().indexOf("In-person Tutoring at Pittsford Community Library") != -1) {
      tuteeLocation.push("In-Person at Pittsford Library");
    } 
    if (rangeVirtual.getValue().toString().indexOf("In-person Tutoring at Brighton Memorial Library") != -1) {
      tuteeLocation.push("In-Person at Brighton Library");
    }
    tuteeFullProfile.push(tuteeLocation.join(","));

    tuteeFullProfile.push(rangeTime.getValue().toString());

    var allTuteeSubjects = [];
    allTuteeSubjects = allTuteeSubjects.concat(rangeMath.getValue().toString().split(", "), rangeScience.getValue().toString().split(", "), rangeHistory.getValue().toString().split(", "), rangeLanguage.getValue().toString().split(", "), rangeECs.getValue().toString().split(", "));
    allTuteeSubjects = allTuteeSubjects.filter(n => n);
    for (var i = 0; i < allTuteeSubjects.length; i++) {
      var equal = false;
      for (var z = 0; z < cellArr.length; z++) {
        if (allTuteeSubjects[i] == cellArr[z]) {
          equal = true;
        }
      }
      if (!equal) {
        tuteeSubjects.push(allTuteeSubjects[i]);
      }
    }

    tuteeFullProfile.push(tuteeSubjects);

    if (tuteeSubjects.length > 0) {
      if (diffWeeks(submissionDate) > 2) {
        priorityTutees.push(tuteeFullProfile);
      }
      else {
        tutees.push(tuteeFullProfile);
      }
    }

    row++;
  }

  // set up tutee selection question
  if (form.getItems().length == 0) {
    form.addMultipleChoiceItem().setTitle("Which tutee would you like to pair with?").setRequired(true);
  }

  var tuteeBios = "";
  var choices = [];
  // set up priority tutees bio and pages
  var num = 0;
  for (var i = 0; i < priorityTutees.length; i++) {
    if (i == 0) {
      tuteeBios += "PRIORITY TUTEES:\nPlease try to pair with these tutees first!\n\n";
    }
    if (i != 0) {
      tuteeBios += "\n\n\n";
    }
    tuteeBios += "(PRIORITY) Tutee " + priorityTutees[i][0] + ":\n";
    var location = priorityTutees[i][2].split(",");
    var bio = "Grade: " + priorityTutees[i][1] + "\nLocation: " + location.join(", ") + "\n\nTime: " + priorityTutees[i][3] + "\n\nSubjects:";
    var subjects = [];
    for (var z = 0; z < priorityTutees[i][4].length; z++) {
      bio += "\n- " + priorityTutees[i][4][z];
      subjects.push(priorityTutees[i][4][z]);
    }
    tuteeBios += bio;
    if (form.getItems().length > 3 * (i+1)) {
      var tuteePage = form.getItems()[1 + 3 * i].asPageBreakItem().setTitle("(PRIORITY) Tutee " + priorityTutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      form.getItems()[2 + 3 * i].asCheckboxItem().setTitle("Which subjects would you like to tutor?").setChoiceValues(subjects).setRequired(true);
      form.getItems()[3 + 3 * i].asMultipleChoiceItem().setTitle("Would you like to tutor In-Person or Virtual?").setChoiceValues(location).setRequired(true);
    }
    else {
      var tuteePage = form.addPageBreakItem().setTitle("(PRIORITY) Tutee " + priorityTutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      box = form.addCheckboxItem();
      choice = form.addMultipleChoiceItem();
      box.setTitle("Which subjects would you like to tutor?");
      box.setChoiceValues(subjects).setRequired(true);
      choice.setTitle("Would you like to tutor In-Person or Virtual?");
      choice.setChoiceValues(location).setRequired(true);
    }

    choices.push(form.getItems()[0].asMultipleChoiceItem().createChoice("(PRIORITY) Tutee " + priorityTutees[i][0], tuteePage));
    num = i + 1;
    if (i == priorityTutees.length - 1) {
      tuteeBios += "\n\n\n\n";
    }
  }

  // set up non-priority tutees bio and pages
  for (var i = 0; i < tutees.length; i++) {
    if (i == 0) {
      tuteeBios += "NON-PRIORITY TUTEES:\n";
    }
    if (i != 0) {
      tuteeBios += "\n\n\n";
    }
    tuteeBios += "Tutee " + tutees[i][0] + ":\n";
    var bio = "Grade: " + tutees[i][1] + "\nLocation: " + tutees[i][2].split(",").join(", ") + "\n\nTime: " + tutees[i][3] + "\n\nSubjects:";
    var subjects = [];
    for (var z = 0; z < tutees[i][4].length; z++) {
      bio += "\n- " + tutees[i][4][z];
      subjects.push(tutees[i][4][z]);
    }
    var location = tutees[i][2].split(",");
    tuteeBios += bio;
    if (form.getItems().length > 3 * (i + 1 + num)) {
      var tuteePage = form.getItems()[1 + 3 * (i + num)].asPageBreakItem().setTitle("Tutee " + tutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      form.getItems()[2 + 3 * (i + num)].asCheckboxItem().setTitle("Which subjects would you like to tutor?").setChoiceValues(subjects).setRequired(true);
      form.getItems()[3 + 3 * (i + num)].asMultipleChoiceItem().setTitle("Would you like to tutor In-Person or Virtual?").setChoiceValues(location).setRequired(true);
    }
    else {
      var tuteePage = form.addPageBreakItem().setTitle("Tutee " + tutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      box = form.addCheckboxItem();
      choice = form.addMultipleChoiceItem();
      box.setTitle("Which subjects would you like to tutor?");
      box.setChoiceValues(subjects).setRequired(true);
      choice.setTitle("Would you like to tutor In-Person or Virtual?");
      choice.setChoiceValues(location).setRequired(true);
    }

    choices.push(form.getItems()[0].asMultipleChoiceItem().createChoice("Tutee " + tutees[i][0], tuteePage));
  }
  // push data to form
  if (choices.length == 0) {
    for (var i = 0; i < form.getItems().length; i++) {
      form.deleteItem(i);
    }
    form.setDescription("There are no tutees available right now!");
  }
  else {
    form.setDescription(tuteeBios + "\n\n\nPlease use the same email you signed up with!");
    form.getItems()[0].asMultipleChoiceItem().setChoices(choices);
  }
}

function paired() {
  var form = FormApp.getActiveForm();
  var tuteeSheet = SpreadsheetApp.openById("").getSheets()[2];
  var tutorSheet = SpreadsheetApp.openById("").getSheets()[1];
  var MASTERSheet = SpreadsheetApp.openById("").getSheets()[0];
  var tutorAgreement = DriveApp.getFileById("");

  var response = form.getResponses()[form.getResponses().length-1];
  var tuteeRow = parseInt(response.getItemResponses()[0].getResponse().toString().split(" ")[0] == "(PRIORITY)" ? response.getItemResponses()[0].getResponse().toString().split(" ")[2] : response.getItemResponses()[0].getResponse().toString().split(" ")[1]);

  var rangeTuteeSubjectPaired = tuteeSheet.getRange(tuteeRow,21).getCell(1,1);
  var rangeTuteeName = tuteeSheet.getRange(tuteeRow,4).getCell(1,1);
  var rangeTuteeEmail = tuteeSheet.getRange(tuteeRow,2).getCell(1,1);
  var rangeTuteeEmailParent = tuteeSheet.getRange(tuteeRow,3).getCell(1,1);
  var rangeTuteePhone = tuteeSheet.getRange(tuteeRow,6).getCell(1,1);

  var row = 2;
  while (MASTERSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    row++;
  }
  var rangeMASTERTuteeName = MASTERSheet.getRange(row,1).getCell(1,1);
  var rangeMASTERTuteeEmail = MASTERSheet.getRange(row,2).getCell(1,1);
  var rangeMASTERTuteeEmailParent = MASTERSheet.getRange(row,3).getCell(1,1);
  var rangeMASTERTuteePhone = MASTERSheet.getRange(row,4).getCell(1,1);
  var rangeMASTERTutorName = MASTERSheet.getRange(row,5).getCell(1,1);
  var rangeMASTERTutorEmail = MASTERSheet.getRange(row,6).getCell(1,1);
  var rangeMASTERTutorPhone = MASTERSheet.getRange(row,7).getCell(1,1);
  var rangeMASTERSubject = MASTERSheet.getRange(row,8).getCell(1,1);
  var rangeMASTERLocation = MASTERSheet.getRange(row,9).getCell(1,1);

  var row = 2;
  while (tutorSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    if (tutorSheet.getRange(row,2).getCell(1,1).getValue() == response.getRespondentEmail()) {
      break;
    }
    row++;
  }
  if (tutorSheet.getRange(row,1).getCell(1,1).getValue() == '') {
    autoPairer();
    throw "Not a registered tutor! Email: " + response.getRespondentEmail();
  }
    
  var cellVal = rangeTuteeSubjectPaired.getValue();
  var subjectsForEmail = "";
  allSubjects = cellVal.toString().split(',');
  subjects = [];
  for (var i = 0; i < response.getItemResponses()[1].getResponse().length; i++) {
    if (i == response.getItemResponses()[1].getResponse().length - 1 && i != 0) {
      subjectsForEmail += ", and "  
    }
    else if (i != 0) {
      subjectsForEmail += ", ";
    }
    for (let z in allSubjects) {
      if (allSubjects[z] == response.getItemResponses()[1].getResponse()[i].toString()) {
        if (sendEmail) {
          MailApp.sendEmail(
            response.getRespondentEmail(),
            "Rochester Peer Tutoring Pairing Failure!",
            "Dear " + tutorSheet.getRange(row,3).getCell(1,1).getValue().trim() + ", \n\nUnfortunately, your tutee was already paired with someone else. Feel free to choose another tutee using the same form! \n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please contact us if you encounter any problems!"
          );
        }
        autoPairer();
        throw "ALREADY PAIRED!";
      }
    }
    subjects.push(response.getItemResponses()[1].getResponse()[i].toString());
    subjectsForEmail += response.getItemResponses()[1].getResponse()[i];
  }
  function removeDuplicates(arr) {
    return arr.filter((item,
        index) => arr.indexOf(item) === index);
  }
  subjects = subjects.filter(n => n);
  subjects = removeDuplicates(subjects);
  allSubjects = allSubjects.concat(subjects);
  allSubjects = allSubjects.filter(n => n);
  allSubjects = removeDuplicates(allSubjects);
  rangeTuteeSubjectPaired.setValue(allSubjects.join(','));

  autoPairer();
  var virtual;
  if (response.getItemResponses()[2].getResponse() == "Virtual") {
    virtual = "virtually";
  } 
  else if (response.getItemResponses()[2].getResponse() == "In-Person at Pittsford Library") {
    virtual = "in-person at Pittsford Community Library";
  } 
  else if (response.getItemResponses()[2].getResponse() == "In-Person at Brighton Library") {
    virtual = "in-person at Brighton Memorial Library";
  } 
  else {
    virtual = "unknown location";
    throw "ERROR: failed to process virtual/in person";
  }
  rangeMASTERTuteeName.setValue(rangeTuteeName.getCell(1,1).getValue());
  rangeMASTERTuteeEmail.setValue(rangeTuteeEmail.getCell(1,1).getValue());
  rangeMASTERTuteeEmailParent.setValue(rangeTuteeEmailParent.getCell(1,1).getValue());
  rangeMASTERTuteePhone.setValue(rangeTuteePhone.getCell(1,1).getValue());
  rangeMASTERTutorEmail.setValue(response.getRespondentEmail());
  rangeMASTERTutorName.setValue(tutorSheet.getRange(row,3).getCell(1,1).getValue());
  rangeMASTERTutorPhone.setValue(tutorSheet.getRange(row,6).getCell(1,1).getValue());
  rangeMASTERSubject.setValue(subjects);
  rangeMASTERLocation.setValue(virtual);

  if (sendEmail) {
    MailApp.sendEmail(
      rangeTuteeEmail.getValue(),
      "Rochester Peer Tutoring Tutor Match: " + tutorSheet.getRange(row,3).getCell(1,1).getValue().trim(),
      "Dear " + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim() + ", \n\nCongratulations! You've been paired with a tutor for " + subjectsForEmail + ". You will be working " + virtual + " with your tutor. Please discuss with your tutor to see what times can work. If there are conflicts or a different time works better, that is something you can discuss individually with your tutor. However, once you do so, please let us know so that we can keep track of the sessions.\n\nYour tutor is " + tutorSheet.getRange(row,3).getCell(1,1).getValue().trim() + ", and is in " + tutorSheet.getRange(row,7).getCell(1,1).getValue().trim() + " at " + tutorSheet.getRange(row,5).getCell(1,1).getValue().trim() + ". They will contact you soon to touch base and know you better. Below is their contact information: \n\n" +
      tutorSheet.getRange(row,3).getCell(1,1).getValue().trim() + ": \nEmail: " + response.getRespondentEmail() + "\nPhone Number: " + tutorSheet.getRange(row,6).getCell(1,1).getValue() + "\n\nAttached is a document highlighting all tutor and tutee expectations - if you and your tutor have read through the document together and agree to all the terms, please sign it and send it back to us as soon as possible. Once we receive the document, you will be officially registered into our system as a tutor-tutee pair. Both you and your tutor are responsible for abiding by these terms - if these rules are breached by either party, please let us know so we can help resolve the issue. If the behavior continues after two warnings, the non-abiding party will be suspended from our program for the remainder of the school year.\n\nIf you have any questions, issues, or concerns, please email us. We look forward to hearing about your progress together in the coming weeks! \n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please reply to this email with any problems or concerns!", {
          cc: response.getRespondentEmail() + "," + rangeTuteeEmailParent.getValue(),
          attachments: tutorAgreement
      });

    MailApp.sendEmail(
    response.getRespondentEmail(),
    "Rochester Peer Tutoring Tutee Info: " + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim(),
    "Dear " + tutorSheet.getRange(row,3).getCell(1,1).getValue().trim() + ", \n\nCongratulations! You've been paired with a tutee, " + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim() + " a " + tuteeSheet.getRange(tuteeRow,7).getCell(1,1).getValue().trim() + "r at " + tuteeSheet.getRange(tuteeRow,5).getCell(1,1).getValue().trim() + ", for " + subjectsForEmail + ". You will be working " + virtual + " with your tutee. Below is their contact information: \n\n" + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim() + "\nEmail: " + tuteeSheet.getRange(tuteeRow,2).getCell(1,1).getValue().trim() + "\nParent's Email: " + tuteeSheet.getRange(tuteeRow,3).getCell(1,1).getValue().trim() + "\nPhone Number: " + tuteeSheet.getRange(tuteeRow,6).getCell(1,1).getValue() + "\n\nHere is their introduction:\n" + tuteeSheet.getRange(tuteeRow,8).getCell(1,1).getValue().trim() + "\n\nHere are any specific learning needs they might have:\n" + tuteeSheet.getRange(tuteeRow,9).getCell(1,1).getValue().trim() + "\n\n\nHere is the form where you can log your hours:\nhttps://forms.gle/YVLpZBtPKSVnpacu7\n\nIf you have any questions, issues, or concerns, please email us. Good luck with tutoring!\n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please reply to this email with any problems or concerns!");
  }
}

function numUnpaired() {
  var form = FormApp.getActiveForm();
  // take the id which is the part after the "/d"
  // https://docs.google.com/spreadsheets/d/Nu7OpaMgqbDyO62L7P3K1r8kpQ7L9L4rXowgKJ8hCyCv
  var sheet = SpreadsheetApp.openById("").getSheets()[2];
  var webhookReminder = "";

  // console.log(form.getTitle()); // check that we have the right form
  // console.log(sheet.getName()); // check that we have the right sheet

  // grabbing sheets data
  var priorityTutees = [];
  var tutees = [];
  var row = 2;

  while (sheet.getRange(row,1).getCell(1,1).getValue() != '') {
    var tuteeFullProfile = [];
    var tuteeSubjects = [];

    var submissionDate = new Date(sheet.getRange(row,1).getCell(1,1).getValue());
    
    var rangeGrade = sheet.getRange(row,7).getCell(1,1);
    var rangeVirtual = sheet.getRange(row,10).getCell(1,1);
    var rangeTime = sheet.getRange(row,11).getCell(1,1);

    var rangeMath = sheet.getRange(row,12).getCell(1,1);
    var rangeScience = sheet.getRange(row,13).getCell(1,1);
    var rangeHistory = sheet.getRange(row,14).getCell(1,1);
    var rangeLanguage = sheet.getRange(row,15).getCell(1,1);
    var rangeECs = sheet.getRange(row,16).getCell(1,1);

    var rangePaired = sheet.getRange(row,21).getCell(1,1);
    var cellArr = rangePaired.getValue().split(',');

    tuteeFullProfile.push(row.toString());

    tuteeFullProfile.push(rangeGrade.getValue().toString());

    var tuteeLocation = [];
    if (rangeVirtual.getValue().toString().indexOf("Virtual Tutoring via Zoom") != -1) {
      tuteeLocation.push("Virtual");
    } 
    if (rangeVirtual.getValue().toString().indexOf("In-person Tutoring at Pittsford Community Library") != -1) {
      tuteeLocation.push("In-Person at Pittsford Library");
    } 
    if (rangeVirtual.getValue().toString().indexOf("In-person Tutoring at Brighton Memorial Library") != -1) {
      tuteeLocation.push("In-Person at Brighton Library");
    }
    tuteeFullProfile.push(tuteeLocation.join(","));

    tuteeFullProfile.push(rangeTime.getValue().toString());

    var allTuteeSubjects = [];
    allTuteeSubjects = allTuteeSubjects.concat(rangeMath.getValue().toString().split(", "), rangeScience.getValue().toString().split(", "), rangeHistory.getValue().toString().split(", "), rangeLanguage.getValue().toString().split(", "), rangeECs.getValue().toString().split(", "));
    allTuteeSubjects = allTuteeSubjects.filter(n => n);
    for (var i = 0; i < allTuteeSubjects.length; i++) {
      var equal = false;
      for (var z = 0; z < cellArr.length; z++) {
        if (allTuteeSubjects[i] == cellArr[z]) {
          equal = true;
        }
      }
      if (!equal) {
        tuteeSubjects.push(allTuteeSubjects[i]);
      }
    }

    tuteeFullProfile.push(tuteeSubjects);

    if (tuteeSubjects.length > 0) {
      if (diffWeeks(submissionDate) > 2) {
        priorityTutees.push(tuteeFullProfile);
      }
      else {
        tutees.push(tuteeFullProfile);
      }
    }

    row++;
  }

  var payload = {
    "num": (priorityTutees.length + tutees.length).toString()
  }
  var options = {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : JSON.stringify(payload)
  };
  if ((priorityTutees.length + tutees.length) > 0) {
    UrlFetchApp.fetch(webhookReminder, options);
  }
}

function diffWeeks(dt) {
  var today = new Date();
  var year = today.getFullYear();
  var month = today.getMonth() + 1;
  var day = today.getDate();
  var day = 24 * 60 * 60 * 1000;
  var weeks = (today - dt) / day / 7;
  return weeks;
}