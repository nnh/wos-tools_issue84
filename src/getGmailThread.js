const linkColNumber =
  Number(PropertiesService.getScriptProperties().getProperty('linkColIdx')) + 1;
const titleColNumber =
  Number(PropertiesService.getScriptProperties().getProperty('titleColIdx')) +
  1;

function getOutputSheet_() {
  const outputSheetName =
    PropertiesService.getScriptProperties().getProperty('outputSheetName');
  const outputSheet = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('outputSheetId')
  ).getSheetByName(outputSheetName);
  return outputSheet;
}
function getTitleList_(outputSheet) {
  const titleList = new Set(
    outputSheet
      .getDataRange()
      .getValues()
      .map(x =>
        x[
          PropertiesService.getScriptProperties().getProperty('titleColIdx')
        ].replace(/^[0-9a-z]+_/, '')
      )
  );
  titleList.delete('ドキュメントタイトル');
  return titleList;
}
function editDocumentByTitleMain() {
  const outputFolderId =
    PropertiesService.getScriptProperties().getProperty('outputFolderId');

  if (!outputFolderId) {
    console.error('Output folder ID not found in script properties.');
    return;
  }
  const outputFolder = DriveApp.getFolderById(outputFolderId);
  //  const targetTitle = "Fwd: Fw: 【水戸医療センター】NHO病院別英文論文リストについて";
  const outputSheet = getOutputSheet_();
  const titleList = getTitleList_(outputSheet);
  titleList.forEach(targetTitle =>
    editDocumentByTitle_(outputFolder, targetTitle, outputSheet)
  );
}
function getGmailThreadByTitleMain() {
  const targetTitle =
    'Fwd: Fw: 【水戸医療センター】NHO病院別英文論文リストについて';
  const titleList = new Set(targetTitle);
  titleList.forEach(targetTitle =>
    getGmailThreadByTitle_(targetTitle, outputSheet)
  );
}
function getThreadByTitle_(title) {
  // Define the search query
  const searchQuery = `subject:(${title})`;
  // Search for Gmail threads matching the criteria
  return GmailApp.search(searchQuery);
}
function createFileName_(thread) {
  const threadTitle = thread.getFirstMessageSubject();

  // Generate a unique filename based on the thread ID and title
  const fileName = thread.getId() + '_' + threadTitle;
  return fileName;
}
function getGmailThreadByTitle_(title, outputSheet) {
  const threads = getThreadByTitle_(title);
  if (threads.length === 0) {
    return;
  }
  const res = threads.map(thread => createFileName_(thread));
  if (res.length === 0) {
    return;
  }

  const existingTitles = outputSheet
    .getDataRange()
    .getValues()
    .map(
      x => x[PropertiesService.getScriptProperties().getProperty('titleColIdx')]
    );
  const outputTitles = res
    .map(x => (!existingTitles.includes(x) ? x : null))
    .filter(x => x !== null);
  if (outputTitles.length === 0) {
    return;
  }
  const outputRow = outputSheet.getLastRow() + 1;
  const fileNames = outputTitles.length === 1 ? [outputTitles] : outputTitles;
  outputSheet
    .getRange(outputRow, titleColNumber, fileNames.length, 1)
    .setValues(fileNames);
}
function editDocumentByTitle_(outputFolder, title, outputSheet) {
  const targetRows = outputSheet
    .getDataRange()
    .getValues()
    .map((x, idx) =>
      new RegExp(title.trim()).test(
        `^[0-9a-z]+_${x[
          PropertiesService.getScriptProperties().getProperty('titleColIdx')
        ].trim()}$`
      )
        ? idx + 1
        : null
    )
    .filter(x => x !== null);
  if (targetRows.length === 0) {
    return;
  }
  const threads = getThreadByTitle_(title);
  if (threads.length === 0) {
    return;
  }
  threads.forEach(thread => {
    const fileName = createFileName_(thread);
    deleteTargetFileIfExists_(outputFolder, fileName);
    const file = findGoogleDoc_(outputFolder, fileName);
    const document = DocumentApp.openById(file.getId());
    extractAndAppendToDocument_(thread, document);
    targetRows.forEach(row => {
      const linkCell = outputSheet.getRange(row, linkColNumber);
      linkCell.setValue(file.getUrl());
    });
  });
}
