/**
 * Gets all Google Document files (Google Docs) in the specified folder and returns their names and URLs in a 2D array.
 *
 * @param {Folder} folder - The Google Drive folder to search for Google Document files.
 * @return {Array[]} - A 2D array containing file names and URLs of Google Document files.
 */
function getAllGoogleDocsInFolder_(folder) {
  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  const docInfoArray = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileInfo = [file.getName(), file.getUrl()];
    docInfoArray.push(fileInfo);
  }

  return docInfoArray;
}
function processGoogleDocsInfoToSpreadsheet() {
  const columnIdx = new Map([['documentUrl', 1]]);
  const folderId =
    PropertiesService.getScriptProperties().getProperty('outputFolderId');
  const folder = DriveApp.getFolderById(folderId);
  const googleDocsInfo = getAllGoogleDocsInFolder_(folder);
  const outputSheet = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('outputSheetId')
  ).getSheets()[0];
  const beforeValues = outputSheet.getDataRange().getValues();
  const outputValues = extractUniqueRecords_(
    googleDocsInfo,
    beforeValues,
    columnIdx.get('documentUrl')
  );
  if (outputValues.length === 0) {
    return;
  }
  const outputStartRowNumber = beforeValues.length + 1;
  outputSheet
    .getRange(
      outputStartRowNumber,
      1,
      outputValues.length,
      outputValues[0].length
    )
    .setValues(outputValues);
}

/**
 * Extracts unique records from two-dimensional arrays A and B based on the specified key index.
 *
 * @param {Array} arrayA - The first two-dimensional array.
 * @param {Array} arrayB - The second two-dimensional array.
 * @param {number} keyIndex - The index of the key column to use for comparison.
 * @return {Array} - An array containing unique records from array A that are not in array B.
 */
function extractUniqueRecords_(arrayA, arrayB, keyIndex) {
  // Create a set of keys from two-dimensional array B
  const keySetB = new Set(arrayB.map(row => row[keyIndex]));
  // Extract records from two-dimensional array A that do not exist in B
  const uniqueRecords = arrayA.filter(row => !keySetB.has(row[keyIndex]));
  return uniqueRecords;
}
