/**
* Development Problems:
* For some reason, saving the attachments is causing a "you're using too much google drive" error.
* Just saving new messages at the same speed (currently set to 1 per second) does not.
* There are "serialization" errors at the end of the script when it is run and it actually saves something.
* The date system seems to be working, now.
* 
* After using it for a while, I think I'm happy to save the attachments myself, because they all need to be renamed in any case.
* Going to be removing the folder creation parts, and just having it save the emails as PDF.
* 
* I'm now having problems where it is regularly exceeding the allowed execution time. I need to examine and see if there is a way
* to prune the elements that it is dealing with.
**/

/**
 * The spreadsheet contains four columns.
 * The first column is the label in Gmail that should be acted on.  This is set by the user.
 * The second column is the folder to which messages filed under that label should be saved.  This is set by the user.
 * The third column is the email address to which messages should be forwarded.  This is set by the user.
 * The fourth column is the last time the script was run against that label name.  This is set by the code.
 * The fifth column is the last time an email was archived under that label.  This is set by the code.
 * The sixth column is the interval (in hours) that should have passed before the script checks that label again.  This is set by the code.
 */

/**
* There is now an "Emails to File" sheet in the file.
* it contains the data used to trigger a Zapier zap that creates a communications entry for the matter.
*/

/**
 * Create a menu item for running the script.
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "Archive Emails", functionName: "RunArchive"},
    {name: "Check Sheet", functionName: "CheckSheet"}
  ];
  ss.addMenu("Archiver", menuEntries);
}

/**
 * Get all Gmail threads for the specified label
 *
 * @param {GmailLabel} label GmailLabel object to get threads for
 * @return {GmailThread[]} an array of threads marked with this label
 */
function getThreadsForLabel(label) {
  var threads = label.getThreads();
  return threads;
}

/**
 * Get all Gmail messages for the specified Gmail thread
 *
 * @param {GmailThread} thread object to get messages for
 * @return {GmailMessage[]} an array of messages contained in the specified thread
 */
function getMessagesforThread(thread) {
  var messages = thread.getMessages();
  return messages;
}

/**
 * Create a Google Drive Folder
 *
 * @param {String} baseFolder name of the base folder
 * @param {String} folderName name of the folder
 * @return {Folder} the folder object created representing the new folder 
 */
function createDriveFolder(baseFolder, folderName) {
  var baseFolderObject = DocsList.getFolder(baseFolder);
  return baseFolderObject.createFolder(folderName);
}

/**
 * Check to see if the sheet has good data
 * (Taken out of main function for execution time reasons)
 */
function CheckSheet() {
  //Logger.log("Starting");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ActiveRange = ss.getActiveRange();
  var labels = ActiveRange.getValues();
  // test for valid data.
  for(var b=1; b<labels.length; b++) { // b=1 allows headers in the first row of spreadsheet.
    //Logger.log("Checking row " + b);
    var gmailLabel = labels[b][0];
    var driveFolder = labels[b][1];
    var labelObj = GmailApp.getUserLabelByName(gmailLabel);
    var folder = DocsList.getFolder(driveFolder);
    if(labelObj == null) { Browser.msgBox("Invalid from label: " + gmailLabel); return null;}
    if(folder == null) { Browser.msgBox("Invalid folder: " + driveFolder); return null;}
  } // end for loop
}


function testIDs() {
  var threads = GmailApp.search("in:inbox l:-Notifications");
  for (i in threads) {
    for (j in threads[i].getMessages()) {
      Logger.log("Date: " + threads[i].getMessages()[j].getDate() + " ID: " + threads[i].getMessages()[j].getId());
    }
  }
}

/**
 * Archive Emails
 * Go through labels.
 * For each row, confirm the label and folder exist.
 * *** This may be where I can punch in some efficiency, if I can grab labels with messages newer than X only, or grab emails that are newer.
 * Grab all the threads in the label.  For each thread where the last message is later than the last time the script ran
 *   Grab all the messages in that thread.  For each message where the date is later than the last time the script ran
 *     Put the message in the folder
 *     Email it to the address if there is one.
 *     *** To make better use of the Clio Communications, what I'd like to do is put an entry into a spreadsheet, and have zapier create a
 *     *** communication entry out of that spreadsheet, so that it doesn't look like it came from me.
 *     *** would need to get rid of spreadsheet entries on a new run, assuming zapier runs more frequently.
 * Set the date to now, and go to the next label.
 *
 *
 *
 * Refactoring:
 * Note Start Time
 * Get Data from Spreadsheet.
 * Delete the old emails to file data
 * For each row, calculate "next run" (Last Run + Interval (in hours))
 * Sort data by next run ascending.
 * For each row (label)
 *   check to see if we have run out of time, and if so, quit.
 *   Confirm label and folder exist
 *   Grab all threads
 *   for each thread
 *     if it has messages newer than last saved message
 *       grab all messages in the thread
 *       for each message
 *         if message is newer than last saved message
 *           add it to the list of messages to be saved
 *     sort messages to be saved by sent date
 *   if there are messages to be saved, decrement the interval (min 0)
 *   if there are no messages to be saved, increment the interval (max 24)
 *   for each message to be saved
 *     put it in the folder
 *     add it to the emails to file sheet
 *     update last saved message date.
 *   set last run date to now
 */
function RunArchive() {
  var endTime = new Date().getTime()+200000;  // set finish time for 200 seconds from now.
  var endTimeDate = new Date(endTime);
  var sleepTime = 10;
  var testing = true;
  Logger.log("Starting");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ActiveRange = ss.getDataRange();
  var labels = ActiveRange.getValues();
  var EmailsToFile = ss.getSheetByName("Emails to File"); // I'm assuming Zapier will have scanned this file 4 times since the last hourly run, and that it can be deleted.
  var lastRow = EmailsToFile.getLastRow();
  EmailsToFile.deleteRows(2,lastRow); // should just leave the header row
  
  labels.splice(0,1); // remove the header row

  
  // Fill in blanks with zeros.
  for(var q=0; q<labels.length; q++) {
    if (labels[q][3]=="") { labels[q][3]=new Date("January 1, 1970");}
    if (labels[q][4]=="") { labels[q][4]=labels[q][3];}
    if (labels[q][5]=="") { labels[q][5]="0";}
    //Set the spreadsheet row reference for each label
    labels[q][6] = q+2; // spreadsheets count from one, and I have removed the header row, so +2.
  }
  
  // Sort the labels by when they should run next.
  labels.sort(sortLabelsByNextRunDate);
  
  Logger.log("Processing Rows");
  for(var i=0; i<labels.length; i++) {
    //Check to see if we are out of time, and if so, quit.
    Logger.log("Checking time");
    var now = new Date().getTime();
    if ( now > endTime ) {Logger.log("Have run for more than 200 seconds, quitting."); return null;}
    Logger.log("Processing Row " + i);
    var gmailLabel = labels[i][0];
    Logger.log("gmailLabel is " + gmailLabel);
    var matterName = gmailLabel.substring(5);
    Logger.log("Matter Name is " + matterName);
    var driveFolder = labels[i][1];
    Logger.log("driveFolder is " + driveFolder);
    var emailAddress = labels[i][2];
    Logger.log("emailAddress is " + emailAddress);
    Logger.log("date is [" + labels[i][3] + "]");
    var lastRunDate = labels[i][3];
    Logger.log("lastRunDate is " + lastRunDate);
    var lastRunCell = ActiveRange.getCell(labels[i][6],4);
    Logger.log("Content of last run date cell is " + lastRunCell.getValue());
    var lastMessageDate = labels[i][4];
    var lastMessageCell = ActiveRange.getCell(labels[i][6],5);
    Logger.log("lastMessageDate is " + lastMessageDate);
    var interval = labels[i][5];
    var intervalCell = ActiveRange.getCell(labels[i][6],6);
    Logger.log("interval is " + interval);
    
    var labelObj = GmailApp.getUserLabelByName(gmailLabel);
    // Get all threads with from label
    if (labelObj != null) {
      var countThreads = getThreadsForLabel(labelObj);
    } else {
      MailApp.sendEmail("jason@roundtablelaw.ca", "Error in Archiver Script - no such label" + gmailLabel, Logger.getLog());
      var countThreads = null;
    }
    var msgsToStore = [];
    // for each thread
    Logger.log("Processing Threads");
    for(var j=0; j<countThreads.length; j++) {
      Logger.log("Processing Thread " + j);
      //check to see when the thread was last updated
      Logger.log("Comparing " + countThreads[j].getLastMessageDate() + " to lastMessageDate");
      if(countThreads[j].getLastMessageDate() > lastMessageDate) {
        Logger.log("Thread " + j + " has new messages");
        // get all the messages for that thread
        var messagesArr = getMessagesforThread(countThreads[j]);
        // for each message
        Logger.log("Processing Messages.");
        for(var k=0; k<messagesArr.length; k++) {
          Logger.log("Processing Message " + k);
          // check to see if it was sent after the last run
          if(messagesArr[k].getDate()>lastMessageDate) {
            msgsToStore.push(messagesArr[k]); // This adds the current message to the messages to be stored.            
          } // message date
        } // each message
      } // thread date
    } // Each thread
    
    // We should now have an array with all of the messages from this label that need to be stored from every thread.
    // It needs to be sorted by date ascending. I think so that if it fails midway, the most recent run date will catch the ones that weren't hit.
    Logger.log("Sorting " + msgsToStore.length + " Messages");
    msgsToStore.sort(sortMessagesByDate);
    
    // If there are messages to store
    if (msgsToStore.length > 0) {  // if there are messages to store
      var Folder = DocsList.getFolder(driveFolder); // get the drive folder
      if (interval > 0) {  // if the interval is not already zero
        interval--; // decrement the interval
        intervalCell.setValue(interval);
      }
    } else { // If there are no messages to store
      if (interval < 24) { // and the interval is not already 24
        interval++; // increment the interval
        intervalCell.setValue(interval);
      }
    }
        
    
    // now, for each message to store
    for(var l=0; l<msgsToStore.length; l++) {
      Logger.log("Storing message " + l);
      //store it
      var messageId = msgsToStore[l].getId();
      var messageDate = Utilities.formatDate(msgsToStore[l].getDate(), Session.getTimeZone(), "yyyy MM dd");
      var messageFrom = msgsToStore[l].getFrom();
      var messageSubject = msgsToStore[l].getSubject();
      var messageBody = msgsToStore[l].getRawContent();
      var messageAttachments = msgsToStore[l].getAttachments();
      var messageRecipient = msgsToStore[l].getTo();
          
      // Create the name and folder for the message
      var newMessageName = messageDate + " email to " + messageRecipient + " from " + messageFrom + " re " + messageSubject + " " + messageId;
        
      // Create the message PDF inside the folder
      var htmlBodyFile = Folder.createFile('body.html', messageBody, "text/html");
      var pdfBlob = htmlBodyFile.getAs('application/pdf');
      pdfBlob.setName(newMessageName + ".pdf");
      Folder.createFile(pdfBlob);
      Utilities.sleep(sleepTime);  // wait after creating something on the drive.
      htmlBodyFile.setTrashed(true);
    
      /*
            // Save attachments
            Logger.log("Saving Attachments");
            for(var i = 0; i < messageAttachments.length; i++) {
              Logger.log("Saving Attachment " + i);
              var attachmentBlob = messageAttachments[i].copyBlob();
              newFolder.createFile(attachmentBlob);
              Utilities.sleep(sleepTime);  // wait after creating something on the drive.
            } // each attachment
            */  
      
      
      // This deals with the situation where the message subject is very long, such as when I send faxes to rcfax.com
      trimmed_subject = msgsToStore[l].getSubject().substring(0, 250);
      
      // If there is an email address, forward the message to that address
      /* if(emailAddress != ""){
        Logger.log("There is an email address, forwarding message");
        msgsToStore[l].forward(emailAddress, {subject: trimmed_subject,}); // This uses the trimmed subject to avoid "too long" errors.
      }
      */ // Depreciated, now using the emails to file spreadsheet.
      
      // This is where it should create an entry in the emails to file sheet.
      // Column Order: Subject, Body, To, From, Date, Matter
      EmailsToFile.appendRow([messageSubject, messageBody, messageRecipient, messageFrom, messageDate, matterName]);
      
      // Now that it has been stored, update the spreadsheet with the date of that message.
      Logger.log("Updating last message saved to " + msgsToStore[l].getDate());
      lastMessageCell.setValue(msgsToStore[l].getDate());
    } // each message to store.
    
    Logger.log("Updating date label run to " + endTimeDate);
    lastRunCell.setValue(endTimeDate);

  } // Each label
  Logger.log("Looks like we're done.");
  return null;
} //RunArchive

function sortMessagesByDate(a, b) {
  return a.getDate() - b.getDate();
}

function sortLabelsByNextRunDate(a, b) {
  //Logger.log("Comparing time " + a[3] + " of type " + typeof a[3] + " and interval " + a[5] + " to time " + b[3] + " and interval " + b[5]);
  //Logger.log(a[3].getTime() + " " + a[5]*3600000 + " " + b[3].getTime() + " " + b[5]*3600000);
  return (a[3].getTime()+a[5]*3600000) - (b[3].getTime()+b[5]*3600000);
}
