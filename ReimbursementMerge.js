// This code can be used to automatically generate PDF files from approved form submissions, and email those PDF files to the form submitter. This code is meant to be attached to a Google Spreadsheet as a Google Script.

// By Jeremy Hissong, NYU Shanghai
// jeremy.hissong@nyu.edu

function onEdit(event) {
  
  // These functions move rows back and forth between tabs of a Google Spreadsheet.
  // This helped our university's finance team maintain accurate records by keeping track of which requests had already been processed.
  // Our needs required that we separate out requests for amounts greater than ¥800, so the code below puts these requests on a different spreadsheet tab.
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
  var d = new Date();
  var date = d.getDate();
  date = ("0" + date).slice(-2);
  var month = d.getMonth() + 1;
  month = ("0" + month).slice(-2);
  var year = d.getFullYear();
  var today = year.toString() + "-" + month + "-" + date;
  
  if(s.getName() == "Pending" && r.getColumn() == 13 && r.getValue() == "Yes") {
    var row = r.getRow();
    var numColumns = s.getLastColumn();
    var range = s.getRange(row, 4);
    var amount = range.getValue();
    if(amount <= 800) { var targetSheet = ss.getSheetByName("Forms Received (under ¥800)");}
    if(amount > 800) { var targetSheet = ss.getSheetByName("Forms Received (over ¥800)");}
    range = s.getRange(row, 11);
    range.setValue(today); // Write date fapiaos and form were received
    var numRows = targetSheet.getLastRow();
    targetSheet.insertRowsAfter(numRows, 1);
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);    
  }
    
  if(s.getName() == "Forms Received (under ¥800)" && r.getColumn() == 13 && r.getValue() == "No") {
    var row = r.getRow();
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Pending");
    var numRows = targetSheet.getLastRow();
    targetSheet.insertRowsAfter(numRows, 1);
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);  
  }
  
  if(s.getName() == "Forms Received (over ¥800)" && r.getColumn() == 13 && r.getValue() == "No") {
    var row = r.getRow();
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Pending");
    var numRows = targetSheet.getLastRow();
    targetSheet.insertRowsAfter(numRows, 1);
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);  
  }
  
  if(s.getName() == "Forms Received (under ¥800)" && r.getColumn() == 14 && r.getValue() == "Yes") {
    var row = r.getRow();
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Completed");
    var numRows = targetSheet.getLastRow();
    range = s.getRange(row, 12);
    range.setValue(today); // Write date uploaded to Workday
    targetSheet.insertRowsAfter(numRows, 1);
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);
  }
  
  if(s.getName() == "Forms Received (over ¥800)" && r.getColumn() == 14 && r.getValue() == "Yes") {
    var row = r.getRow();
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Completed");
    var numRows = targetSheet.getLastRow();
    range = s.getRange(row, 12);
    range.setValue(today); // Write date uploaded to Workday
    targetSheet.insertRowsAfter(numRows, 1);
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);  
  }
  
  if(s.getName() == "Completed" && r.getColumn() == 14 && r.getValue() == "No") {
    var row = r.getRow();
    var numColumns = s.getLastColumn();
    var range = s.getRange(row, 4);
    var amount = range.getValue();
    if(amount <= 800) { var targetSheet = ss.getSheetByName("Forms Received (under ¥800)");}
    if(amount > 800) { var targetSheet = ss.getSheetByName("Forms Received (over ¥800)");}
    var numRows = targetSheet.getLastRow();
    targetSheet.insertRowsAfter(numRows, 1);
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);
  }
}


function whatShouldIProcess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[4]; // REMEMBER TO CHANGE THIS NUMBER IF ANYTHING CHANGES WITH THE SHEET LAYOUT -- THIS NUMBER (4) referrs to the VIEW ONLY tab, which is the 5th tab of the sheet
  var lastRow = sheet.getLastRow();
  var numAlreadyProcessed = (lastRow-4);
  var alreadyProcessed = sheet.getSheetValues(5, 1, numAlreadyProcessed, 1);
  
  // This part of the code goes through our first OrgSync Form, and makes an array of the submission IDs that need to be processed.
  // You must alter the form code (after "forms/" in the line below) to correspond to the form from which you'd like to pull responses.
  
  var response = UrlFetchApp.fetch("https://api.orgsync.com/api/v2/forms/127013.json?key=PutApiKeyHere&status=Approved"); // <-- it's gotta have that sneaky little .json after the form code
  var json = response.getContentText();
//  Logger.log(response);
  var formData = JSON.parse(json);
  var numApproved = formData.submission_count;
  var needToProcessClubs = [];
  for (var i=0; i<numApproved; i++) {
    needToProcessClubs[needToProcessClubs.length] = formData.submission_ids[i]; // makes a Javascript array of all the approved submission ids from the Clubs form
  }
  var numRemoved = 0;
  for (var i=0; i<numApproved; i++) {
    for (var t=0; t<numAlreadyProcessed; t++) {
      if (alreadyProcessed[t] == ('RR' + formData.submission_ids[i])) {
        needToProcessClubs.splice(i-numRemoved,1);
        numRemoved = numRemoved + 1;
      } 
    }  
  }
  
  // This part of the code goes through our second OrgSync Form, and makes an array of the submission IDs that need to be processed.
  // You must alter the form code (after "forms/" in the line below) to correspond to the form from which you'd like to pull responses.
  
  var response = UrlFetchApp.fetch("https://api.orgsync.com/api/v2/forms/130515.json?key=PutApiKeyHere&status=Approved");
  var json = response.getContentText();
  var formData = JSON.parse(json);
  var numApproved = formData.submission_count;
  var needToProcessResLife = [];
  for (var i=0; i<numApproved; i++) {
    needToProcessResLife[needToProcessResLife.length] = formData.submission_ids[i]; // makes a Javascript array of all the approved submission ids from the Res Life form
  }
  var numRemoved = 0;
  for (var i=0; i<numApproved; i++) {
    for (var t=0; t<numAlreadyProcessed; t++) {
      if (alreadyProcessed[t] == ('RR' + formData.submission_ids[i])) {
        needToProcessResLife.splice(i-numRemoved,1);
        numRemoved = numRemoved + 1;
      } 
    }  
  }
  
  var needToProcess = needToProcessClubs.concat(needToProcessResLife);
  return needToProcess;  
}


function apiPullAndProcess() {
  var needToProcess = whatShouldIProcess();
  
  for (var j=0; j<needToProcess.length; j++) {
    var response = UrlFetchApp.fetch("https://api.orgsync.com/api/v2/form_submissions/" + needToProcess[j] + ".json?key=PutApiKeyHere");
    var json = response.getContentText();
    var submission = JSON.parse(json);
    
    var d = new Date();
    var date = d.getDate();
    date = ("0" + date).slice(-2);
    var month = d.getMonth() + 1;
    month = ("0" + month).slice(-2);
    var year = d.getFullYear();
    var today = year.toString() + "-" + month + "-" + date;
    
    var eventDate = submission.responses[5].data;
    if (submission.responses[5].data == null) {
      eventDate = 'N/A';
    }
    
    var onBehalf;
    if (submission.on_behalf_of) {
      onBehalf = submission.on_behalf_of.name;
    }
    else {
      onBehalf = 'N/A';
    }
    
    var description = ("Club/Organization Name: " + onBehalf + "\n" + "Event Date: " + eventDate + "\n" + "Expense: " + submission.responses[4].data + "\n" + "Request ID: RR" + submission.id);
    var cc1 = submission.responses[34].data.name;
    var fund1 = submission.responses[35].data.name;
    var program1 = submission.responses[36].data.name;
    var cc2 = '', fund2 = '', program2 = '';
    var cc3 = '', fund3 = '', program3 = '';
    var cc4 = '', fund4 = '', program4 = '';
    var cc5 = '', fund5 = '', program5 = '';
    var item1 = submission.responses[8].data.name;
    var item2 = '', item3 = '', item4 = '', item5 = '';
    
    var totalAmount = 0.0;
    var amount1 = parseFloat(submission.responses[9].data);
    totalAmount += amount1;
    amount1 = '¥' + amount1.toFixed(2);
    
    var amount2 = '';
    if (submission.responses[14].data != '') {
      amount2 = parseFloat(submission.responses[14].data);
      totalAmount += amount2;
      amount2 = '¥' + amount2.toFixed(2);
      item2 = submission.responses[13].data.name;
      cc2 = cc1;
      fund2 = fund1;
      program2 = program1;
    }
    
    var amount3 = '';
    if (submission.responses[19].data != '') {
      amount3 = parseFloat(submission.responses[19].data);
      totalAmount += amount3;
      amount3 = '¥' + amount3.toFixed(2);
      item3 = submission.responses[18].data.name;
      cc3 = cc1;
      fund3 = fund1;
      program3 = program1;
    }
    
    var amount4 = '';
    if (submission.responses[24].data != '') {
      amount4 = parseFloat(submission.responses[24].data);
      totalAmount += amount4;
      amount4 = '¥' + amount4.toFixed(2);
      item4 = submission.responses[23].data.name;
      cc4 = cc1;
      fund4 = fund1;
      program4 = program1;
    }
    
    var amount5 = '';
    if (submission.responses[29].data != '') {
      amount5 = parseFloat(submission.responses[29].data);
      totalAmount += amount5;
      amount5 = '¥' + amount5.toFixed(2);
      item5 = submission.responses[28].data.name;
      cc5 = cc1;
      fund5 = fund1;
      program5 = program1;
    }
    
    
    // Consider expanding this code to all fields if issue keeps occuring with things getting mixed up in the JSON from the API pull.
    var approverPhone = '';
	var submissionLength = (submission.responses).length;
    for (r=0; r<submissionLength; r++) {
    	if (submission.responses[r].element.name == 'Approver phone number:') {
          approverPhone = submission.responses[r].data;
        }
    }
    
        
    var data = {
      PayeeName:submission.responses[0].data,
      NYUID:submission.responses[1].data,
      TotalAmount:totalAmount,
      ID:'RR' + submission.id,
      
      Item1:item1,
      Amount1:amount1,
      CC1:cc1,
      Fund1:fund1,
      Program1:program1,
      
      Item2:item2,
      Amount2:amount2,
      CC2:cc2,
      Fund2:fund2,
      Program2:program2,
      
      Item3:item3,
      Amount3:amount3,
      CC3:cc3,
      Fund3:fund3,
      Program3:program3,
      
      Item4:item4,
      Amount4:amount4,
      CC4:cc4,
      Fund4:fund4,
      Program4:program4,
      
      Item5:item5,
      Amount5:amount5,
      CC5:cc5,
      Fund5:fund5,
      Program5:program5,
      
      Words:submission.responses[31].data,
      Description:description,
      PayeeEmail:submission.responses[2].data,
      PayeePhone:submission.responses[3].data,
      ApproverName:submission.responses[37].data,
      ApproverSignature:"OrgSync Approved: RR" + submission.id,
      ApproverPhone:approverPhone,
      Date:today
    } // end of data object
    mergeAndMail(data);
  }
}


function mergeAndMail(data) {

// The information in this section must be mapped to the Google Doc that you want to use as a merge template.

  var folder = DriveApp.getFolderById("0B25oPMdA2ye-YVNqejdKM0ZRRWc"),
      copyFile = DriveApp.getFileById("1DHZP3WM57Vvk0oGST7uok8iWRn9qeE7UtG1x8p0yhIU").makeCopy(data.ID),
      copyId = copyFile.getId(),
      copyDoc = DocumentApp.openById(copyId),
      copyBody = copyDoc.getBody(),
      stylizedTotal = ('¥' + (data.TotalAmount).toFixed(2));
  
// This chunk of code replaces merge parameters in the Google Doc. You may add or take away from this list as needed to accommodate the merge fields in your Google Doc.

  copyBody.replaceText('%PayeeName%', data.PayeeName);
  copyBody.replaceText('%NYUID%', data.NYUID);
  copyBody.replaceText('%Item1%', data.Item1);
  copyBody.replaceText('%Item2%', data.Item2);
  copyBody.replaceText('%Item3%', data.Item3);
  copyBody.replaceText('%Item4%', data.Item4);
  copyBody.replaceText('%Item5%', data.Item5);
  copyBody.replaceText('%Amount1%', data.Amount1);
  copyBody.replaceText('%Amount2%', data.Amount2);
  copyBody.replaceText('%Amount3%', data.Amount3);
  copyBody.replaceText('%Amount4%', data.Amount4);
  copyBody.replaceText('%Amount5%', data.Amount5);
  copyBody.replaceText('%CC1%', data.CC1);
  copyBody.replaceText('%CC2%', data.CC2);
  copyBody.replaceText('%CC3%', data.CC3);
  copyBody.replaceText('%CC4%', data.CC4);
  copyBody.replaceText('%CC5%', data.CC5);
  copyBody.replaceText('%Fund1%', data.Fund1);
  copyBody.replaceText('%Fund2%', data.Fund2);
  copyBody.replaceText('%Fund3%', data.Fund3);
  copyBody.replaceText('%Fund4%', data.Fund4);
  copyBody.replaceText('%Fund5%', data.Fund5);
  copyBody.replaceText('%Program1%', data.Program1);
  copyBody.replaceText('%Program2%', data.Program2);
  copyBody.replaceText('%Program3%', data.Program3);
  copyBody.replaceText('%Program4%', data.Program4);
  copyBody.replaceText('%Program5%', data.Program5);
  copyBody.replaceText('%Total%', stylizedTotal);
  copyBody.replaceText('%Words%', data.Words);
  copyBody.replaceText('%Description%', data.Description);
  copyBody.replaceText('%PayeeEmail%', data.PayeeEmail);
  copyBody.replaceText('%PayeePhone%', data.PayeePhone);
  copyBody.replaceText('%ApproverName%', data.ApproverName);
  copyBody.replaceText('%ApproverSignature%', data.ApproverSignature);
  copyBody.replaceText('%ApproverPhone%', data.ApproverPhone);
  copyBody.replaceText('%Date%', data.Date);
  copyDoc.saveAndClose();
  
// This chunk of code is used to generate an email to send out the PDF document. Our institution required a different email to be sent for reimbursements under/over ¥800.
  
  var divStuff = '<div style="color: #000000; font-family: arial, sans-serif; font-size: 13px; font-weight: normal; background-color: #ffffff;"><div class="im" style="color: #000000;">';
  var greeting = 'Dear ' + data.PayeeName + ':<br /><br />';
  var beginningOfMsgOver = 'Your reimbursement request for ' + stylizedTotal + ' is now available for processing. Please print out the attached PDF and bring it to the Bursar window at Room 1062, along with your NYU ID card and all original fapiaos (发票) for these expenses.<br /><br />';
  var beginningOfMsgUnder = 'Your reimbursement request for ' + stylizedTotal + ' is now available for pickup. To receive your reimbursement, please print out the attached PDF and bring it to the Bursar window at Room 1062, along with your NYU ID card and all original fapiaos (发票) for these expenses.<br /><br />'
  var over800 = 'Please note that since these expenses exceed ¥800, the Bursar is unable to provide cash reimbursement for this request. Instead, these funds will be deposited directly into your bank account within 10 working days from the time you submit your fapiaos.<br /><br />';
  var ifQuestions = 'If you have any questions, please visit the Student Life help desk in Room 110.<br /><br />';
  var signature = 'Office of Student Life<br /><br />';
  var emailSig = '<u><font color="#cccccc">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</font></u>'
    + '<u><font color="#cccccc">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</font></u><u><font color="#cccccc">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</font></u></div></div>'
    + '<div style="color: #000000; font-family: arial, sans-serif; font-size: 13px; font-weight: normal; background-color: #ffffff;"><br />'
    + '<div class="im" style="color: #000000; line-height: 145%;"><img class="CToWUd" style="float: left; padding-right: 12pt; padding-top: 2pt;" src="https://ci6.googleusercontent.com/proxy/hNFJDHcfyVFJKwqFUYAolRBwEYWnW0k8YfOgDSAqI0N2MJrt3cl8VEcDvNxwBdP36XxGyz81FBP1QAaWNWchIQJbq5JRbISSe6I6uATphN77GAyxUMFW6No=s0-d-e1-ft#https://dl.dropboxusercontent.com/u/83124818/nyu_shanghai_logo.jpeg" alt="" />'
    + '<strong>Office of Student Life &nbsp;| &nbsp;学生事务部</strong><br />NYU Shanghai&nbsp;&nbsp;|&nbsp;&nbsp;上海纽约大学<br />'
    + '1555 Century Avenue, Room 110 &nbsp;|&nbsp;&nbsp;世纪大道1555号110室<br />'
    + 'Pudong New Area, Shanghai, China&nbsp;&nbsp;|&nbsp;&nbsp;中国上海浦东新区<br />'
    + 'Office (电话):&nbsp;+86 021 2059 5339&nbsp;&nbsp;|&nbsp;&nbsp;Email (邮箱):&nbsp;<a href="mailto:shanghai.involvement@nyu.edu">shanghai.involvement@nyu.edu</a></div>';
  
  var plainTextOver = 'Dear ' + data.PayeeName + ':\n\nYour reimbursement request for ' + stylizedTotal + ' is now available for processing. Please print out the attached PDF and bring it to the Bursar window at Room 1062, along with your NYU ID card and all original fapiaos (发票) for these expenses.\n\nPlease note that since these expenses exceed ¥800, the Bursar is unable to provide cash reimbursement for this request. Instead, these funds will be deposited directly into your bank account within 10 working days from the time you submit your fapiaos.\n\nIf you have any questions, please visit the Student Life help desk in Room 110.';
  var plainTextUnder = 'Dear ' + data.PayeeName + ':\n\nYour reimbursement request for ' + stylizedTotal + ' is now available for pickup. To receive your reimbursement, please print out the attached PDF and bring it to the Bursar window at Room 1062, along with your NYU ID card and all original fapiaos (发票) for these expenses.\n\nIf you have any questions, please visit the Student Life help desk in Room 110.';
  var plainText;
  
  var html;
  if (data.TotalAmount > 800) {
    html = '<html><body>' + divStuff + greeting + beginningOfMsgOver + over800 + ifQuestions + signature + emailSig + '</body></html>';
    plainText = plainTextOver;
  }
  else {
    html = '<html><body>' + divStuff + greeting + beginningOfMsgUnder + ifQuestions + signature + emailSig + '</body></html>';
    plainText = plainTextUnder;
  }
  
  // This part of the code should be amended to contain the information about the account that will send out the email with PDF attachment.
  
  GmailApp.sendEmail(data.PayeeEmail, 'Reimbursement Request for ' + stylizedTotal, plainText, {
     attachments: [copyFile.getAs(MimeType.PDF)],
     from: 'shanghai.student.reimbursements@nyu.edu',
     name: 'NYU Shanghai Student Reimbursements',
     replyTo: 'NYU Shanghai Student Involvement <shanghai.involvement@nyu.edu>',
     bcc:'NYU Shanghai Student Reimbursements Staff <shanghai.student.reimbursements-staff-group@nyu.edu>',
     htmlBody: html
  });

  // This part of the code adds information to our Google Sheet about the request that was just processed. In addition to presenting the data in an easily readable spreadsheet, the Google Sheet also serves as a "cache" for the script, allowing it to keep track of which approved submissions have already been processed.
  
  var dateFormat = 'yyyy-MM-dd';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheets()[4]; // this number [4] referrs to the fifth sheet (which is currently the VIEW ONLY sheet)
  sheet.appendRow([data.ID, data.PayeeName, data.NYUID, data.TotalAmount, data.CC1, data.Fund1, data.Program1, '', data.ApproverName, data.Date]);
  sheet.getRange('J2:J').setNumberFormat(dateFormat);
  
  var sheet = ss.getSheets()[0]; // this number [0] referrs to the first sheet (which is currently the Pending sheet)
  sheet.appendRow([data.ID, data.PayeeName, data.NYUID, data.TotalAmount, data.CC1, data.Fund1, data.Program1, '', data.ApproverName, data.Date, '', '', 'No', 'No', '']);
  sheet.getRange('J2:J').setNumberFormat(dateFormat);
  sheet.getRange('K2:K').setNumberFormat(dateFormat);
  sheet.getRange('L2:L').setNumberFormat(dateFormat);

  copyFile.setTrashed(true);
}


function main() {
  apiPullAndProcess();
}