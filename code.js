/**
 * This function checks for new emails in the Gmail inbox,
 * creates a subfolder for emails with attachments in a specified Google Drive folder,
 * and indexes their details in a specified Google Sheet, starting from the oldest new emails.
 */
function indexNewEmails() {
  // Prompt for Google Sheet ID and Drive folder ID
  const sheetId = Browser.inputBox('Enter the Google Sheet ID:');
  const sharedFolderId = Browser.inputBox('Enter the Google Drive folder ID:');

  // Open the Google Sheet
  const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  
  // Get the last indexed email date from the sheet
  const lastRow = sheet.getLastRow();
  let lastIndexedDate = new Date(0); // Start with a very old date
  if (lastRow > 1) {
    lastIndexedDate = new Date(sheet.getRange(lastRow, 1).getValue());
  }
  
  // Search for all emails and filter based on the last indexed date
  const threads = GmailApp.search(`after:${Utilities.formatDate(lastIndexedDate, Session.getScriptTimeZone(), 'yyyy/MM/dd')}`);
  
  // Check if there are any new emails
  if (threads.length === 0) {
    Logger.log('No new emails since the last indexed date.');
    return;
  }

  // Create an array to hold new emails
  let newEmails = [];

  // Loop through each thread and gather details
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // Only process messages newer than the last indexed date
      if (message.getDate() > lastIndexedDate) {
        newEmails.push(message);
      }
    });
  });

  // Sort new emails by date (oldest first)
  newEmails.sort((a, b) => a.getDate() - b.getDate());

  // Get the shared folder
  const sharedFolder = DriveApp.getFolderById(sharedFolderId);

  // Process the sorted emails
  newEmails.forEach(message => {
    const date = message.getDate();
    const from = message.getFrom();
    const clearEmail = from.replace(/.*<|>.*/g, '').trim(); // Extract clear email
    const subject = message.getSubject();
    const link = message.getThread().getPermalink();
    let folderLink = '';
    let attachmentDetails = '';

    // Check if the message has attachments
    if (message.getAttachments().length > 0) {
      // Format the date for the folder name
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
      const folderName = `${formattedDate} ${clearEmail}`;
      
      // Create a new subfolder for attachments
      const subfolder = sharedFolder.createFolder(folderName);
      folderLink = subfolder.getUrl();
      
      // Move attachments to the new subfolder
      message.getAttachments().forEach(attachment => {
        subfolder.createFile(attachment);
        attachmentDetails += `${attachment.getName()} (${attachment.getSize()} bytes), `;
      });
      // Remove trailing comma and space
      attachmentDetails = attachmentDetails.slice(0, -2);
    } else {
      attachmentDetails = 'No attachments';
    }
    
    // Append the email details to the sheet with date formatting
    sheet.appendRow([Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), from, clearEmail, subject, link, folderLink]);
    
    // Log details for each email
    Logger.log(`Indexed Email: Date: ${date}, From: ${from}, Clear Email: ${clearEmail}, Subject: ${subject}, Link: ${link}, Attachments: ${attachmentDetails}`);
  });

  Logger.log('Indexing complete. New emails indexed.');
}
