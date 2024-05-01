/**
 * Allows for the pairing of any tutor or tutee no matter the pairing status
 * 
 * 
 * AutoPairer:
 *  Auto pulls data from Google Sheet to automatically update every minute
 * 
 * Paired:
 *  Runs when a tutor fills out the form
 *  Updates excel to track paired tutors
 * 
 * DiffWeeks:
 *  Finds the difference in weeks between a JS date and today, returning that value
 */

var sendEmail = true;

function autoPairer() {
  var form = FormApp.getActiveForm();
  var tutorSheet = SpreadsheetApp.openById("").getSheets()[1];
  var tuteeSheet = SpreadsheetApp.openById("").getSheets()[2];

  // console.log(form.getTitle()); // check that we have the right form
  // console.log(sheet.getName()); // check that we have the right sheet

  // grabbing sheets data
  var priorityTutees = [];
  var tutees = [];
  var row = 2;
  // checks row to see if there is data
  while (tuteeSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    var tuteeFullProfile = [];
    var tuteeSubjects = [];

    var submissionDate = new Date(tuteeSheet.getRange(row,1).getCell(1,1).getValue());
    
    var rangeTuteeGrade = tuteeSheet.getRange(row,7).getCell(1,1);
    var rangeTuteeVirtual = tuteeSheet.getRange(row,10).getCell(1,1);
    var rangeTuteeTime = tuteeSheet.getRange(row,11).getCell(1,1);
    var rangeTuteeName = tuteeSheet.getRange(row,4).getCell(1,1).getValue();
    var rangeTuteeEmail = tuteeSheet.getRange(row,2).getCell(1,1).getValue();
    var rangeTuteeParentEmail = tuteeSheet.getRange(row,3).getCell(1,1).getValue();
    var rangeTuteeSchool = tuteeSheet.getRange(row,5).getCell(1,1).getValue();
    var rangeTuteePhone = tuteeSheet.getRange(row,6).getCell(1,1).getValue();

    var rangeTuteeMath = tuteeSheet.getRange(row,12).getCell(1,1);
    var rangeTuteeScience = tuteeSheet.getRange(row,13).getCell(1,1);
    var rangeTuteeHistory = tuteeSheet.getRange(row,14).getCell(1,1);
    var rangeTuteeLanguage = tuteeSheet.getRange(row,15).getCell(1,1);
    var rangeTuteeECs = tuteeSheet.getRange(row,16).getCell(1,1);

    var rangeTuteePaired = tuteeSheet.getRange(row,21).getCell(1,1);
    var cellArr = rangeTuteePaired.getValue().split(',');

    tuteeFullProfile.push(row.toString());

    tuteeFullProfile.push(rangeTuteeGrade.getValue().toString());

    var tuteeLocation = [];
    if (rangeTuteeVirtual.getValue().toString().indexOf("Virtual Tutoring via Zoom") != -1) {
      tuteeLocation.push("Virtual");
    } 
    if (rangeTuteeVirtual.getValue().toString().indexOf("In-person Tutoring at Pittsford Community Library") != -1) {
      tuteeLocation.push("In-Person at Pittsford Library");
    } 
    if (rangeTuteeVirtual.getValue().toString().indexOf("In-person Tutoring at Brighton Memorial Library") != -1) {
      tuteeLocation.push("In-Person at Brighton Library");
    }
    tuteeFullProfile.push(tuteeLocation.join(","));

    tuteeFullProfile.push(rangeTuteeTime.getValue().toString());

    var allTuteeSubjects = [];
    allTuteeSubjects = allTuteeSubjects.concat(rangeTuteeMath.getValue().toString().split(", "), rangeTuteeScience.getValue().toString().split(", "), rangeTuteeHistory.getValue().toString().split(", "), rangeTuteeLanguage.getValue().toString().split(", "), rangeTuteeECs.getValue().toString().split(", "));
    allTuteeSubjects = allTuteeSubjects.filter(n => n);
    var num = allTuteeSubjects.length;
    for (var i = 0; i < allTuteeSubjects.length; i++) {
      var equal = false;
      for (var z = 0; z < cellArr.length; z++) {
        if (allTuteeSubjects[i] == cellArr[z]) {
          tuteeSubjects.push(allTuteeSubjects[i] + " (PAIRED)");
          equal = true;
          num--;
        }
      }
      if (!equal) {
        tuteeSubjects.push(allTuteeSubjects[i]);
      }
    }

    tuteeFullProfile.push(tuteeSubjects);
    tuteeFullProfile.push(rangeTuteeName);
    tuteeFullProfile.push(rangeTuteeEmail);
    tuteeFullProfile.push(rangeTuteeParentEmail);
    tuteeFullProfile.push(rangeTuteeSchool);
    tuteeFullProfile.push(rangeTuteePhone);
    tuteeFullProfile.push(num == 0);

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
  var tutors = [];
  var row = 2;
  while (tutorSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    var rangeTutorName = tutorSheet.getRange(row,3).getCell(1,1).getValue();
    var rangeTutorGrade = tutorSheet.getRange(row,7).getCell(1,1).getValue();
    var rangeTutorEmail = tutorSheet.getRange(row,2).getCell(1,1).getValue();
    var rangeTutorSchool = tutorSheet.getRange(row,5).getCell(1,1).getValue();

    tutors.push(form.getItems()[1].asMultipleChoiceItem().createChoice(rangeTutorName + " (" + rangeTutorGrade + ", " + rangeTutorSchool + ", " + rangeTutorEmail+ ")"));

    row++;
  }

  // tutor and tutee selection question
  {
    if (form.getItems().length == 0) {
      form.addMultipleChoiceItem().setTitle("Which tutee would you like to pair?").setRequired(true);
      form.addMultipleChoiceItem().setTitle("Which tutor would you like to pair?").setRequired(true);
    }
    else if (form.getItems().length == 1) {
      form.getItems()[0].asMultipleChoiceItem().setTitle("Which tutee would you like to pair?").setRequired(true);
      form.addMultipleChoiceItem().setTitle("Which tutor would you like to pair?").setRequired(true);
    }
    else {
      form.getItems()[0].asMultipleChoiceItem().setTitle("Which tutee would you like to pair?").setRequired(true);
      form.getItems()[1].asMultipleChoiceItem().setTitle("Which tutor would you like to pair?").setRequired(true);
    }
  }

  var tuteeBios = "";
  var choices = [];
  /*
  form all the tutees bios for tutors to view
  general format:
  name
  in person or zoom
  times
  subjects
  */
  var num = 0;
  for (var i = 0; i < priorityTutees.length; i++) {
    if (i == 0) {
      tuteeBios += "PRIORITY TUTEES:\nPlease try to pair with these tutees first!\n\n";
    }
    if (i != 0) {
      tuteeBios += "\n\n\n";
    }
    tuteeBios += (priorityTutees[i][10] ? "(FULLY PAIRED) " : "(PRIORITY) ") + "Tutee " + priorityTutees[i][0] + ":\n";
    var bio = "Name: " + priorityTutees[i][5] + "\nEmail: " + priorityTutees[i][6] + "\nParent Email: " + priorityTutees[i][7] + "\nPhone Number: " + priorityTutees[i][9] + "\nSchool: " + priorityTutees[i][8] + "\nGrade: " + priorityTutees[i][1] + "\nLocation: " + priorityTutees[i][2].split(",").join(", ") + "\n\nTime: " + priorityTutees[i][3] + "\n\nSubjects:";
    var subjects = [];
    for (var z = 0; z < priorityTutees[i][4].length; z++) {
      bio += "\n- " + priorityTutees[i][4][z];
      subjects.push(priorityTutees[i][4][z]);
    }
    var location = [];
    if (priorityTutees[i][2].toString().indexOf("Virtual") != -1) {
      location.push("Virtual (PREFERED)");
    } 
    else {
      location.push("Virtual (OVERRIDE)");
    }
    if (priorityTutees[i][2].toString().indexOf("In-Person at Pittsford Library") != -1) {
      location.push("In-Person at Pittsford (PREFERED)");
    } 
    else {
      location.push("In-Person at Pittsford (OVERRIDE)");
    }
    if (priorityTutees[i][2].toString().indexOf("In-Person at Brighton Library") != -1) {
      location.push("In-Person at Brighton (PREFERED)");
    }
    else {
      location.push("In-Person at Brighton (OVERRIDE)");
    }
    tuteeBios += bio;
    if (form.getItems().length > 3 * (i+2)) {
      var tuteePage = form.getItems()[2 + 3 * i].asPageBreakItem().setTitle((priorityTutees[i][10] ? "(FULLY PAIRED) " : "(PRIORITY) ") + "Tutee " + priorityTutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      form.getItems()[3 + 3 * i].asCheckboxItem().setTitle("Which subjects would you like to tutor?").setChoiceValues(subjects).setRequired(true);
      form.getItems()[4 + 3 * i].asMultipleChoiceItem().setTitle("Would you like to tutor In-Person or Virtual?").setChoiceValues(location).setRequired(true);
    }
    else {
      var tuteePage = form.addPageBreakItem().setTitle((priorityTutees[i][10] ? "(FULLY PAIRED) " : "(PRIORITY) ") + "Tutee " + priorityTutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      box = form.addCheckboxItem();
      choice = form.addMultipleChoiceItem();
      box.setTitle("Which subjects would you like to tutor?");
      box.setChoiceValues(subjects).setRequired(true);
      choice.setTitle("Would you like to tutor In-Person or Virtual?");
      choice.setChoiceValues(location).setRequired(true);
    }

    choices.push(form.getItems()[0].asMultipleChoiceItem().createChoice((priorityTutees[i][10] ? "(FULLY PAIRED) " : "(PRIORITY) ") + "Tutee " + priorityTutees[i][0] + " ("+ priorityTutees[i][5] + ", " + priorityTutees[i][6] +")", tuteePage));
    num = i + 1;
    if (i == priorityTutees.length - 1) {
      tuteeBios += "\n\n\n\n";
    }
  }
  for (var i = 0; i < tutees.length; i++) {
    if (i == 0) {
      tuteeBios += "NON-PRIORITY TUTEES:\n";
    }
    if (i != 0) {
      tuteeBios += "\n\n\n";
    }
    tuteeBios += (tutees[i][10] ? "(FULLY PAIRED) " : "") + "Tutee " + tutees[i][0] + ":\n";
    var bio = "Grade: " + tutees[i][1] + "\nLocation: " + tutees[i][2].split(",").join(", ") + "\n\nTime: " + tutees[i][3] + "\n\nSubjects:";
    var subjects = [];
    for (var z = 0; z < tutees[i][4].length; z++) {
      bio += "\n- " + tutees[i][4][z];
      subjects.push(tutees[i][4][z]);
    }
    var location = [];
    if (tutees[i][2].toString().indexOf("Virtual") != -1) {
      location.push("Virtual (PREFERED)");
    } 
    else {
      location.push("Virtual (OVERRIDE)");
    }
    if (tutees[i][2].toString().indexOf("In-Person at Pittsford Library") != -1) {
      location.push("In-Person at Pittsford (PREFERED)");
    } 
    else {
      location.push("In-Person at Pittsford (OVERRIDE)");
    }
    if (tutees[i][2].toString().indexOf("In-Person at Brighton Library") != -1) {
      location.push("In-Person at Brighton (PREFERED)");
    }
    else {
      location.push("In-Person at Brighton (OVERRIDE)");
    }
    tuteeBios += bio;
    if (form.getItems().length > 3 * (i+1 + num)) {
      var tuteePage = form.getItems()[2 + 3 * (i+num)].asPageBreakItem().setTitle((tutees[i][10] ? "(FULLY PAIRED) " : "") + "Tutee " + tutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      form.getItems()[3 + 3 * (i + num)].asCheckboxItem().setTitle("Which subjects would you like to tutor?").setChoiceValues(subjects).setRequired(true);
      form.getItems()[4 + 3 * (i + num)].asMultipleChoiceItem().setTitle("Would you like to tutor In-Person or Virtual?").setChoiceValues(location).setRequired(true);
    }
    else {
      var tuteePage = form.addPageBreakItem().setTitle((tutees[i][10] ? "(FULLY PAIRED) " : "" + "Tutee ") + tutees[i][0]).setHelpText(bio).setGoToPage(FormApp.PageNavigationType.SUBMIT);
      box = form.addCheckboxItem();
      choice = form.addMultipleChoiceItem();
      box.setTitle("Which subjects would you like to tutor?");
      box.setChoiceValues(subjects).setRequired(true);
      choice.setTitle("Would you like to tutor In-Person or Virtual?");
      choice.setChoiceValues(location).setRequired(true);
    }

    choices.push(form.getItems()[0].asMultipleChoiceItem().createChoice((tutees[i][10] ? "(FULLY PAIRED) " : "") + "Tutee " + tutees[i][0] + " ("+ tutees[i][5] + ", " + tutees[i][6] + ")", tuteePage));
  }
  form.setDescription(tuteeBios);
  form.getItems()[0].asMultipleChoiceItem().setChoices(choices);
  form.getItems()[1].asMultipleChoiceItem().setChoices(tutors);
}

function paired() {
  var form = FormApp.getActiveForm();
  var tuteeSheet = SpreadsheetApp.openById("").getSheets()[2];
  var tutorSheet = SpreadsheetApp.openById("").getSheets()[1];
  var MASTERSheet = SpreadsheetApp.openById("").getSheets()[0];
  var tutorAgreement = DriveApp.getFileById("");

  var response = form.getResponses()[form.getResponses().length-1];
  var split = response.getItemResponses()[0].getResponse().toString().split(" ");
  var tuteeRow = split[0] == "(PRIORITY)" ? split[2] : (split[0] == "(FULLY" ? split[3] : split[1]);

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

  var split = response.getItemResponses()[1].getResponse().toString().split(", ");
  var tutorEmail = split[2].substring(0,(split[2].length-1));

  var row = 2;
  while (tutorSheet.getRange(row,1).getCell(1,1).getValue() != '') {
    if (tutorSheet.getRange(row,2).getCell(1,1).getValue() == tutorEmail) {
      break;
    }
    row++;
  }
  if (tutorSheet.getRange(row,1).getCell(1,1).getValue() == '') {
    autoPairer();
    throw "Not a registered tutor! Email: " + tutorEmail;
  }
    
  var cellVal = rangeTuteeSubjectPaired.getValue();
  var subjectsForEmail = "";
  allSubjects = cellVal.toString().split(',');
  subjects = [];
  for (var i = 0; i < response.getItemResponses()[2].getResponse().length; i++) {
    if (i == response.getItemResponses()[2].getResponse().length - 1 && i != 0) {
      subjectsForEmail += ", and "  
    }
    else if (i != 0) {
      subjectsForEmail += ", ";
    }
    if (response.getItemResponses()[2].getResponse()[i].toString().indexOf('(PAIRED)') != -1) {
      subjects.push(response.getItemResponses()[2].getResponse()[i].toString().split(' (PAIRED)')[0]);
      subjectsForEmail += response.getItemResponses()[2].getResponse()[i].toString().split(' (PAIRED)')[0];
    }
    else {
      for (let z in allSubjects) {
        if (allSubjects[z] == response.getItemResponses()[2].getResponse()[i].toString()) {
          autoPairer();
          throw "ALREADY PAIRED!";
        }
      }
      subjects.push(response.getItemResponses()[2].getResponse()[i].toString());
      subjectsForEmail += response.getItemResponses()[2].getResponse()[i];
    }
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
  
  // updates the form
  autoPairer();

  // email code
  var virtual;
  
  if (response.getItemResponses()[3].getResponse().toString().indexOf("Virtual") != -1) {
    virtual = "virtually";
  } 
  else if (response.getItemResponses()[3].getResponse().toString().indexOf("In-Person at Pittsford Library") != -1) {
    virtual = "in-person at Pittsford Community Library";
  } 
  else if (response.getItemResponses()[3].getResponse().toString().indexOf("In-Person at Brighton Library") != -1) {
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
  rangeMASTERTutorEmail.setValue(tutorEmail);
  rangeMASTERTutorName.setValue(tutorSheet.getRange(row,3).getCell(1,1).getValue());
  rangeMASTERTutorPhone.setValue(tutorSheet.getRange(row,6).getCell(1,1).getValue());
  rangeMASTERSubject.setValue(subjects.join(','));
  rangeMASTERLocation.setValue(virtual);
  if (sendEmail) {
    MailApp.sendEmail(
      rangeTuteeEmail.getValue(),
      "Rochester Peer Tutoring Tutor Match: " + tutorSheet.getRange(row,3).getCell(1,1).getValue().trim(),
      "Dear " + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim() + ", \n\nCongratulations! You've been paired with a tutor for " + subjectsForEmail + ". You will be working " + virtual + " with your tutor. Please discuss with your tutor to see what times can work. If there are conflicts or a different time works better, that is something you can discuss individually with your tutor. However, once you do so, please let us know so that we can keep track of the sessions.\n\nYour tutor is " + tutorSheet.getRange(row,3).getCell(1,1).getValue().trim() + ", and is in " + tutorSheet.getRange(row,7).getCell(1,1).getValue().trim() + " at " + tutorSheet.getRange(row,5).getCell(1,1).getValue().trim() + ". They will contact you soon to touch base and know you better. Below is their contact information: \n\n" +
      tutorSheet.getRange(row,3).getCell(1,1).getValue().trim() + ": \nEmail: " + tutorEmail + "\nPhone Number: " + tutorSheet.getRange(row,6).getCell(1,1).getValue() + "\n\nAttached is a document highlighting all tutor and tutee expectations - if you and your tutor have read through the document together and agree to all the terms, please sign it and send it back to us as soon as possible. Once we receive the document, you will be officially registered into our system as a tutor-tutee pair. Both you and your tutor are responsible for abiding by these terms - if these rules are breached by either party, please let us know so we can help resolve the issue. If the behavior continues after two warnings, the non-abiding party will be suspended from our program for the remainder of the school year.\n\nIf you have any questions, issues, or concerns, please email us. We look forward to hearing about your progress together in the coming weeks! \n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please reply to this email with any problems or concerns!", {
          cc: tutorEmail + "," + rangeTuteeEmailParent.getValue(),
          attachments: tutorAgreement
      });

    MailApp.sendEmail(
      tutorEmail,
      "Rochester Peer Tutoring Tutee Info: " + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim(),
      "Dear " + tutorSheet.getRange(row,3).getCell(1,1).getValue().trim() + ", \n\nCongratulations! You've been paired with a tutee, " + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim() + " a " + tuteeSheet.getRange(tuteeRow,7).getCell(1,1).getValue().trim() + "r at " + tuteeSheet.getRange(tuteeRow,5).getCell(1,1).getValue().trim() + ", for " + subjectsForEmail + ". You will be working " + virtual + " with your tutee. Below is their contact information: \n\n" + tuteeSheet.getRange(tuteeRow,4).getCell(1,1).getValue().trim() + "\nEmail: " + tuteeSheet.getRange(tuteeRow,2).getCell(1,1).getValue().trim() + ":\nParent's Email: " + tuteeSheet.getRange(tuteeRow,3).getCell(1,1).getValue().trim() + "\nPhone Number: " + tuteeSheet.getRange(tuteeRow,6).getCell(1,1).getValue() + "\n\nHere is their introduction:\n" + tuteeSheet.getRange(tuteeRow,8).getCell(1,1).getValue().trim() + "\n\nHere are any specific learning needs they might have:\n" + tuteeSheet.getRange(tuteeRow,9).getCell(1,1).getValue().trim() + "\n\n\nHere is the form where you can log your hours:\nhttps://forms.gle/YVLpZBtPKSVnpacu7\n\nIf you have any questions, issues, or concerns, please email us. Good luck with tutoring!\n\nSincerely,\nRochester Peer Tutoring\n\n\nThis email was sent automatically, please reply to this email with any problems or concerns!");
  }
}

function diffWeeks(dt) {
  var today = new Date();
  var day = today.getDate();
  var day = 24 * 60 * 60 * 1000;
  var weeks = (today - dt) / day / 7;
  return weeks;
}