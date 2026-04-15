function doGet(e) {
  if (e && e.parameter && e.parameter.api === "true") {
    return ContentService
      .createTextOutput(JSON.stringify({ message: "API working" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Index');
}
function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  // ✅ LOGIN
  if (data.action === "login") {
    const result = loginUser(data.username, data.password);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ✅ PREVIEW REFERENCE
  if (data.action === "previewComplaintReference") {
    const result = previewComplaintReference(data.token, data.source);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: "Unknown action" }))
    .setMimeType(ContentService.MimeType.JSON);
}