//
// SCRIPT WEBSITE: https://script.google.com/macros/d/1RM2O0kdr0RMspXE6vgu0yfrUWcLvqfWccAsfnn5ywLQDtd9jkMrU9rdG/edit?splash=yes
//
// FOLDER: https://drive.google.com/drive/folders/0BxVAKBzwjD0BMlVCTGhmelExdTg
//
// template: https://docs.google.com/document/d/1ZSqN-OOxvDV9YLpuDsXhqSWx4X3djUaVehuIlZbanxo/edit
//
// push to Drive with 'gapps upload'
//
function main(){

  // set by user  
  var inputSpreadsheet = "1mrqZ9GJctfWqqshAWcXGZzI2uo6nMLqcSk2-Mj96nRE";
  var subsheet = "ABCFinance";

  var templateID = "1ZSqN-OOxvDV9YLpuDsXhqSWx4X3djUaVehuIlZbanxo";
  var folderID = "0BxVAKBzwjD0BMlVCTGhmelExdTg";

  // where are the emails? set this according to sheet
  var emailsCol = 4
  var firstNameCol = 1

  // call the function that gets the emails from the ABC finance sheet
  var emailsList = extractDataFromSheet( inputSpreadsheet, subsheet, emailsCol );

  var namesList = extractDataFromSheet( inputSpreadsheet, subsheet, firstNameCol );

  // generate the certificate document
  generateDocumentSendMail(emailsList, namesList, templateID, folderID);

}




function extractDataFromSheet( INPUT_SPREADSHEET, SUBSHEET, INFO_INDEX ) {

  // create a data sctructure to hold emails
  var extract_data_container = new Array();

  // getting all data from this specific sheet and subsheet 
  var spreadsheet = SpreadsheetApp.openById( INPUT_SPREADSHEET );
  var sheet = spreadsheet.getSheetByName( SUBSHEET );

  // getting all the values across the sheet
  var values = sheet.getDataRange()
                    .getValues();

  //add emails from list to an array
  for(n=0;n<values.length;++n){

    // x is the index of the column starting from 0
    var cell_content = values[n][ INFO_INDEX ] ; 

    extract_data_container.push(cell_content);

  }

  return extract_data_container;
}



function generateDocumentSendMail(emails, names, templateID, folderID) {

   var copyId = DriveApp.getFileById( templateID )
                        .makeCopy('TSTNG_'+ emails[5] )
                        .getId();

   var copyDoc = DocumentApp.openById(copyId);
   var copyBody = copyDoc.getActiveSection();

   copyBody.replaceText("%WEEKNO%", names[5]);


   copyDoc.saveAndClose();


  var file = DriveApp.getFileById(copyId);

  //to put file in folder
  DriveApp.getFolderById(folderID)
          .addFile(file);


   var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");

   var participant_email = "s.matthew.english+187@gmail.com";

   var subject = "Here's your certificate";
   var body    = "Hello " + names[5] + " " + "<br /><br />" 
   + "Thank you for calling " + emails[5] + " Support. The attached document contains information " 
   + "for you to reference related to the credits we have issued back to your original form of payment." + "<br /><br />" 
   + "If you have any further questions or require additional assistance please let us know." + "<br /><br />" 
   + "Regards," + "<br /><br />" 
   + ", Payments Department" + "<br />" 
   MailApp.sendEmail(participant_email, subject, body, {htmlBody: body, attachments: pdf}); 

}



