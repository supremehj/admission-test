function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var username = String(body.username || "").trim();
  var password = String(body.password || "").trim();

  var cache = CacheService.getScriptCache();
  var cached = cache.get("credentials");
  var data;

  if (cached) {
    data = JSON.parse(cached);
  } else {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    var lastRow = sheet.getLastRow();
    data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
    cache.put("credentials", JSON.stringify(data), 300); // cache for 5 minutes
  }

  var valid = false;
  for (var i = 0; i < data.length; i++) {
    var sheetUser = String(data[i][0]).trim();
    var sheetPass = String(data[i][1]).trim();
    if (sheetUser === username && sheetPass === password) {
      valid = true;
      break;
    }
  }

  var result = { success: valid };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
