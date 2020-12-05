const EMAIL_SENT = 'EMAIL_FORWARDED';

const help_categories = {
  'Plumbing' : 1,
  'Honeybees' : 2,
  'Electricity' : 3,
  'Weeds' : 4,
  'General Frustration with college life' : 5,
  'Snake spotted' : 6,
  'Someone broke my wall' : 7
}

function getEmailString(sheet,max,col) {
  let description ='';
  for(let j=2; j<max;j++){
     let currentcell = sheet.getRange(j,col).getValue();
      if(currentcell === '')
          break;
      else 
          description += ',' + currentcell;
  }
  return description;
}

function SendResponseEmails() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let responseSheet = ss.getSheetByName("Form Responses 1");
  let blockSheet = ss.getSheetByName("Block Wardens");
  let categorySheet = ss.getSheetByName("Help Categories");
  const startRow = 2;
  const numRows = responseSheet.getLastRow()+1;
  const dataRange = responseSheet.getRange(startRow,1,numRows,6);
  const data = dataRange.getValues();
  for(let i=0; i< data.length; ++i){
    let row = data[i];
    let isEmailSent = row[6];
    if(isEmailSent === EMAIL_SENT)
      continue;
    const block = row[2];
    const senderEmail = row[1];
    let cause = row[3];
    let wardenlist = 'rachitkdas@gmail.com';
    let description = 'Description Of The Problem:\n';
    description += "The following complaint was recieved:\n";
    let subject = 'Hostel Complaint reagrding :  ';
    let maxWardenRange = blockSheet.getLastRow();
    let maxCategoryRange = categorySheet.getLastRow();
    
    subject += row[3];
    description += row[4];
    if(block === '')
      return;
    if(block === 'P'){
         wardenlist += getEmailString(blockSheet,maxWardenRange,11);
      }
    else {
      wardenlist += getEmailString(blockSheet,maxWardenRange, block.charCodeAt()-64);
    }
    wardenlist += getEmailString(categorySheet, maxWardenRange, help_categories[cause] );
    description += "This complaint was sent by: ";
    description += senderEmail;
    MailApp.sendEmail(wardenlist, senderEmail, subject, description);
    responseSheet.getRange(startRow + i, 7).setValue(EMAIL_SENT);
    SpreadsheetApp.flush();
  }
}
