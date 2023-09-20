/**
 * Extracts Gmail threads that match the specified criteria and saves them as PDF files in Google Drive.
 *
 * @param {Object} e - Event object from the trigger.
 */
function extractAndSaveThreadsAsPDFs(e) {
  // Get the output folder ID and sender email from script properties
  const outputFolderId =
    PropertiesService.getScriptProperties().getProperty('outputFolderId');
  const senderEmail =
    PropertiesService.getScriptProperties().getProperty('senderEmail');

  if (!outputFolderId || !senderEmail) {
    console.error(
      'Output folder ID or sender email not found in script properties.'
    );
    return;
  }
  const outputFolder = DriveApp.getFolderById(outputFolderId);

  // Calculate the date range (1 day before and after today)
  const today = new Date();
  const oneDay = 24 * 60 * 60 * 1000; // 1 day in milliseconds
  const startDate = new Date(today.getTime() - oneDay);
  const endDate = new Date(today.getTime() + oneDay);
  const dateUtils = new DateUtils();

  // Define the search query
  const searchQuery = `from:${senderEmail} after:${dateUtils.formatDate(
    startDate
  )} before:${dateUtils.formatDate(endDate)}`;

  // Search for Gmail threads matching the criteria
  const threads = GmailApp.search(searchQuery);
  if (threads.length === 0) {
    return;
  }

  // Process each matching thread
  threads.forEach(thread => {
    // Get the thread title
    const threadTitle = thread.getFirstMessageSubject();

    // Generate a unique filename based on the thread ID and title
    const fileName = thread.getId() + '_' + threadTitle;
    deleteTargetFileIfExists_(outputFolder, fileName);
    const file = findGoogleDoc_(outputFolder, fileName);
    const document = DocumentApp.openById(file.getId());
    extractAndAppendToDocument_(thread, document);
    //convertDocumentToPDFAndDelete_(document, fileName, outputFolder);
  });
}

function appendTextToDocumentBody_(messages, document) {
  const heading1Style = {};
  heading1Style[DocumentApp.Attribute.HEADING] =
    DocumentApp.ParagraphHeading.HEADING1;
  const body = document.getBody();
  body.clear();
  messages.forEach(message => {
    body.appendParagraph(message.get('title')).setAttributes(heading1Style);
    body.appendParagraph(message.get('date'));
    body.appendParagraph(`From: ${message.get('from')}`);
    body.appendParagraph(message.get('body'));
  });
  document.saveAndClose();
}

function extractAndAppendToDocument_(thread, document) {
  const messages = thread.getMessages();
  const dateUtils = new DateUtils();
  const messageBodies = messages.map(message => {
    const body = message
      .getBody()
      .replace(/<br>/g, '\n')
      .replace(/<br[^>].+>/g, '\n')
      .replace(/<[^>]*>/g, '');
    const title = message.getSubject();
    const date = dateUtils.formatDateTime(message.getDate());
    const mailFrom = message.getFrom();
    return new Map([
      ['title', title],
      ['body', body],
      ['date', date],
      ['from', mailFrom],
    ]);
  });
  appendTextToDocumentBody_(messageBodies, document);
}
