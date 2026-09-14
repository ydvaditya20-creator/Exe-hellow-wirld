function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheetName = data.sheetName; // 'basahi' ya 'natwa'
    var tableData = data.rows;      // Table ka data array
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": "Sheet not found: " + sheetName}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Har baar naya data aane par purana data delete karne ke liye (Optional)
    // sheet.clearContents(); 
    
    // Naya data rows ko sheet me append karne ke liye
    tableData.forEach(function(row) {
      sheet.appendRow(row);
    });
    
    return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var requestedSheet = e.parameter.sheet; 
  
  if (!requestedSheet) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Sheet name missing"}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  var sheet = ss.getSheetByName(requestedSheet);
  var data = sheet ? sheet.getDataRange().getValues() : [];
  
  // JSON Output bhejna
  return ContentService.createTextOutput(JSON.stringify(data))
                       .setMimeType(ContentService.MimeType.JSON);
}
