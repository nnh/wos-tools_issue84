function sendSpreadsheetUrlToGoogleSpace() {
  const spreadsheetUrl = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('outputSheetId')
  ).getUrl();
  const webhookUrl =
    PropertiesService.getScriptProperties().getProperty('webhookUrl');

  if (webhookUrl) {
    const payload = {
      text: '進捗の確認: ' + spreadsheetUrl,
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
    };

    UrlFetchApp.fetch(webhookUrl, options);
  } else {
    console.log('Webhook URLが設定されていません。');
  }
}
