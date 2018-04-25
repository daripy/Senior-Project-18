function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Run Scripts')

      .addSubMenu(ui.createMenu('Send Emails')
          .addItem('Brainstorm Approval', 'brainstormApproval')
          .addItem('Proposal Approval', 'proposalApproval')
          .addItem('Mentor Agreement Recieved', 'MentorAgreementRecieved')
          .addItem('Service Hour Update', 'ServiceHours')
          .addItem('Brainstorm Update', 'brainstormupdate')
                 )
      .addToUi();
}
function sheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}
function brainstormApproval() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 // ss.setActiveSheet(ss.getSheetByName("MASTER"));
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var numCol = 11;
  var numRows = sheet.getLastRow();
  var EMAIL_SENTY = "APPROVED Sent";
  var EMAIL_SENTN = "NOT APPROVED Sent";
  var dataRange = sheet.getRange(startRow, 1, numRows, numCol);
  var data = dataRange.getValues();
    for (i=0;i<data.length;i++) {
      var row = data[i];
      var emailAddress = row[1];
      var approved = row[6];
      var approvecell = sheet.getRange(startRow + i, 7);
      var comments = approvecell.getNote();
      var recipient = row[2];
      var yessubject = 'Senior Project Brainstorm: Approved';
      var nosubject = 'Senior Project Brainstorm: Not Approved';
      var nomessage = 'Dear ' + recipient + ',\n\n' + 'Your Senior Project Brainstorm has been reviewed but has not been approved, please reach out to Stacie or Rick to discuss next steps.' ;
      var html = 
                'Dear ' + recipient + ',' + '<br />'+'<br />' + 'Your Senior Project Brainstorm has been reviewed and approved. Please click on <a href="http://seniorproject.seattle.academy/">this link</a> for a list of past project sites that have hosted a SAAS student that you can review to determine if there is one that is of interest to you. If so, click on the site to inform us of your interest in that career field and site, then we will forward contact information to you.'
                 + '<br />'+'<br />' + 'To help ensure we do not have multiple students contacting the same site, please inform us of your interest and do not make contact without our approval so that we can best manage the process and support you. If you already have a direct contact and are ready to proceed please fill out this <a href="https://goo.gl/forms/HrzWHbZI5IqV4BmZ2">Project Proposal Form</a>'
                 + '<br />'+'<br />' + 'Thank you and please reach out to Stacie or Rick if you have any questions.' 
                 + '<br />'+'<br />' +'-Rick' ;    
   if(approved=="Yes"){
          MailApp.sendEmail(emailAddress, yessubject,'body',{ htmlBody: html}); 
          approvecell.setValue(EMAIL_SENTY); 
              comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
      if(approved=="No"){
          MailApp.sendEmail(emailAddress, nosubject, nomessage);
          approvecell.setValue(EMAIL_SENTN);
        comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
    }
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert('Emails Sent!');
}

function brainstormupdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 // ss.setActiveSheet(ss.getSheetByName("MASTER"));
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var numCol = 11;
  var numRows = sheet.getLastRow();
  var EMAIL_SENTY = "UPDATE Sent";
  var EMAIL_SENTN = "NOT APPROVED Sent";
  var dataRange = sheet.getRange(startRow, 1, numRows, numCol);
  var data = dataRange.getValues();
    for (i=0;i<data.length;i++) {
      var row = data[i];
      var emailAddress = row[1];
      var approved = row[6];
      var approvecell = sheet.getRange(startRow + i, 7);
      var comments = approvecell.getNote();
      var recipient = row[2];
      var yessubject = 'Senior Project Propsal Link';
      var nosubject = 'Senior Project Brainstorm: Not Approved';
      var nomessage = 'Dear ' + recipient + ',\n\n' + 'Your Senior Project Brainstorm has been reviewed but has not been approved, please reach out to Stacie or Rick to discuss next steps.' ;
      var html = 
                'Dear ' + recipient + ',' + '<br />'+'<br />' + 'If have spoken with your proposed site mentor and have a good idea of what your project will entail please fill out this <a href="https://goo.gl/forms/HrzWHbZI5IqV4BmZ2">Project Proposal Form.</a>'
                 + '<br />'+'<br />' + 'Thank you and please reach out to Stacie or Rick if you have any questions.' 
                 + '<br />'+'<br />' +'-Rick' ;    
   if(approved=="APPROVED Sent"){
          MailApp.sendEmail(emailAddress, yessubject,'body',{ htmlBody: html}); 
          approvecell.setValue(EMAIL_SENTY); 
              comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
      if(approved=="No"){
          MailApp.sendEmail(emailAddress, nosubject, nomessage);
          approvecell.setValue(EMAIL_SENTN);
        comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
    }
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert('Emails Sent!');
}


function proposalApproval() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 // ss.setActiveSheet(ss.getSheetByName("MASTER"));
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var numCol = 11;
  var numRows = sheet.getLastRow();
  var EMAIL_SENTY = "APPROVED Sent";
  var EMAIL_SENTN = "NOT APPROVED Sent";
  var dataRange = sheet.getRange(startRow, 1, numRows, numCol);
  var data = dataRange.getValues();
    for (i=0;i<data.length;i++) {
      var row = data[i];
      var emailAddress = row[1];
      var approved = row[8];
      var approvecell = sheet.getRange(startRow + i, 9);
      var comments = approvecell.getNote();
      var recipient = row[2];
      var mentor = row[4];
      var company = [6];
      var yessubject = 'Senior Project Proposal: Approved';
      var nosubject = 'Senior Project Proposal: Not Approved';
      var nomessage = 'Dear ' + recipient + ',\n\n' + 'Your project proposal has been denied. Please resubmit your project proposal survey with more specific details in the project description and goals/objectives sections. Also please make sure to include all of the contact information for your site and project mentor. If you have any questions, please contact Rick DuPree, rdupree@seattleacademy.org.';
      var html = 
                "Hello " + recipient + "," + '<br />'+'<br />' + "Your Senior Project proposal has been approved. Your next step will be to get your mentor to fill out the  " 
                + '<a href=\"https://goo.gl/forms/7QbHEpY1bmBEOKSz2">Mentor Agreement Form</a>' + ".  You can copy the link and then send to it to your site mentor to complete." 
                + '<br />'+'<br />'+ "Please reach out if you have any questions." +'<br />'+'<br />'+ "Stacie and Rick";     
   if(approved=="Yes"){
          MailApp.sendEmail(emailAddress, yessubject,'body',{ htmlBody: html}); 
          approvecell.setValue(EMAIL_SENTY); 
              comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
      if(approved=="No"){
          MailApp.sendEmail(emailAddress, nosubject, nomessage);
          approvecell.setValue(EMAIL_SENTN);
        comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
    }
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert('Emails Sent!');
}

function MentorAgreementRecieved(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var numCol = 14;
  var numRows = sheet.getLastRow();
  var EMAIL_SENTY = "APPROVED Sent";
  var EMAIL_SENTN = "NOT APPROVED Sent";
  var dataRange = sheet.getRange(startRow, 1, numRows, numCol);
  var data = dataRange.getValues();
    for (i=0;i<data.length;i++) {
      var row = data[i];
      var studentemailAddress = row[1];
      var mentoremailAddress = row[13];
      var approved = row[10];
      var approvecell = sheet.getRange(startRow + i, 11);
      var comments = approvecell.getNote();
      var recipient = row[2];
      var mentor = row[12];
      var mentorbody = "Thank you for agreeing to participate in the Seattle Academy Senior Project program as a site mentor. Attached is a letter with specific project details, if you have any questions please contact Rick DuPree, rdupree@seattleacademy.org, (206) 676-6893.";
      var file = DriveApp.getFileById("1BjDnWfl_ET4QWOvmUMSYzexEj7ajzIlCWiyjUvdhyfI");
      var pdffile = file.getAs(MimeType.PDF)
      var yessubject = 'Senior Project Mentor Agreement Received';
      var mentorsubject = 'Seattle Academy Senior Project Mentorship';
      var nosubject = 'Senior Project Mentor Agreement Not Received';
      var nomessage = 'Dear ' + recipient + ',\n\n' + 'Your project proposal has been denied. Please resubmit your project proposal survey with more specific details in the project description and goals/objectives sections. Also please make sure to include all of the contact information for your site and project mentor. If you have any questions, please contact Rick DuPree, rdupree@seattleacademy.org.';
      var html = 
                "Congratulations " + recipient + "," + '<br />'+'<br />' + " We have received your Mentor Agreement form from your Senior Project site mentor. You have completed all of the steps in the process, and you are now ready to begin your project on Monday April 23. " +'<br />'+'<br />'+ "Please make sure to confirm your work schedule with your mentor in advance of your start date and communicate any schedule conflicts to ensure that your mentor and site is aware and approves the conflicts. Please remember that the expectation is for you to average 25 hours per week at your site, and return to campus every Friday for A-D blocks."+ '<br />'+'<br />'+ " If you have any questions, or need additional information, contact Rick DuPree - rdupree@seattleacademy.org.";    
   
      if(approved=="Yes"){
          MailApp.sendEmail(studentemailAddress, yessubject,'body',{ htmlBody: html});
          MailApp.sendEmail(mentoremailAddress, mentorsubject, mentorbody, {attachments: [pdffile]});
   
        
          approvecell.setValue(EMAIL_SENTY); 
              comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
      if(approved=="No"){
          MailApp.sendEmail(studentemailAddress, nosubject, nomessage);
          approvecell.setValue(EMAIL_SENTN);
        comments = comments + "\nModified: " + (new Date());
        approvecell.setNote(comments);
        
      }
    }
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .alert('Emails Sent!');



}




function ServiceHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("A2:e118");
  var data = dataRange.getValues();
  
  for (i in data) {
    var rowData = data[i];  
      if(rowData[4] > 20){
    var emailAddress = rowData[1];
    var recipient = rowData[2];
    var service = rowData[4]
    var message = 'Dear ' + recipient + ',\n\n' + 'Our records indicate that you still need ' + service + ' service hours to meet the SAAS community service graduation requirement of 160 hours.' + '\n\n' + 'If this is incorrect, or you have completed hours that have not yet been submitted into X2Vol, please inform us immediately so that we can update your hours. If our records are accurate, please submit your plan to meet the requirement.' + '\n\n' +   'Students that have not met the requirement by June 1 will NOT receive their diploma at graduation and final grades will not be released to colleges until the requirement is met.  If you have any questions, please contact me directly.'+ '\n\n' +  'Rick' ;
    var subject = 'Community Service Graduation Requirement NOT MET';
    MailApp.sendEmail(emailAddress, subject, message);
      }
  }
  
  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Emails Sent!');
}