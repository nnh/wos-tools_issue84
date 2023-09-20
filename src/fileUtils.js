function findGoogleDoc_(folder, fileName) {
  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);

  while (files.hasNext()) {
    const file = files.next();
    if (file.getName() === fileName) {
      // Return the Google Document if its name matches
      return file;
    }
  }
  const copyFile = DriveApp.getFileById(
    PropertiesService.getScriptProperties().getProperty('templateDocId')
  ).makeCopy(folder);
  copyFile.setName(fileName);
  return copyFile;
}

function deleteTargetFileIfExists_(folder, fileName) {
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    if (
      file.getName() === fileName &&
      (file.getMimeType() === 'application/pdf' ||
        file.getMimeType() === 'application/vnd.google-apps.document')
    ) {
      // Delete the PDF file if its name and MIME type match
      file.setTrashed(true);
    }
  }
}

function convertDocumentToPDFAndDelete_(document, filename, outputFolder) {
  // Check if the file already exists in the output folder
  const pdfFilename = `${filename}.pdf`;
  deleteTargetFileIfExists_(outputFolder, pdfFilename);
  const pdfBlob = document.getAs('application/pdf');
  const pdfFile = outputFolder.createFile(pdfBlob);
  pdfFile.setName(pdfFilename);
}
