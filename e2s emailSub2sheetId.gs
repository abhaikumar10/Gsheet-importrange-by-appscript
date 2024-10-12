//This returns 'spreadid' from mail contaning 'excel file'.
//emailSubject='Excel/texts of sheet'

function e2s(emailSubject) {
  //Email with exact same subject line latest mail
  const threads = GmailApp.search('subject:"' + emailSubject + '"');
  if (threads.length == 0) {
                                                var ee="No emails found with the subject: " + emailSubject
                                                Logger.log(ee);
                                                var dd={};
                                                dd.sourceID=emailSubject;
                                                logData(dd,ee,null);
                                                return;
                                          }
  const thread = threads[0];
  const messages = thread.getMessages();
  const latestMessage = messages[messages.length - 1];
 
  // Get attachments from the email
  const attachments = latestMessage.getAttachments();

  let excelFile;

  // Find the Excel file in the attachments
  for (const attachment of attachments) {
    if (attachment.getContentType() === MimeType.MICROSOFT_EXCEL || attachment.getContentType() === MimeType.MICROSOFT_EXCEL_LEGACY) {
      excelFile = attachment;
      break;
    }
  }  if (!excelFile) {
                                                    var ee="No Excel attachment found in the email. " + emailSubject
                                                    var dd={};
                                                    dd.sourceID=emailSubject;
                                                      Logger.log(ee); //logData(dd,ee,null);
                                                    return;
                                                  }
 
  // Create a new Google Sheet from the Excel file
  const blobb = excelFile.copyBlob().setContentType(MimeType.MICROSOFT_EXCEL);
  const tempFile = DriveApp.createFile(blobb); //Logger.log(tempFile.getId())
  var excelFileid = tempFile.getId(); //File name of the file to be converted
  var files = DriveApp.getFileById(excelFileid);
  var blob = files.getBlob();                                                                Logger.log(files.getName())
  var config = {
      title:  files.getName(), //sets the title of the converted file
    // parents: [{id: "1UBzAx4onU6IBQ3vtb1tCODzZO2v4rb_u"}], //Use this if wish to store in certain Folder.
      mimeType: MimeType.GOOGLE_SHEETS
  };
  var spreadsheet = Drive.Files.insert(config, blob);
  DriveApp.getFileById(excelFileid).setTrashed(true);                                         Logger.log(spreadsheet.id);
  return spreadsheet.id; //Displays the file ID
}
