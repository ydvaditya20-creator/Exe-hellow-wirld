function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheetName = data.sheetName; // 'basahi' ya 'natwa'
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": "Sheet not found: " + sheetName}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var sheetData = sheet.getDataRange().getValues();
    var voucherNo = data.voucherNo;
    var targetRowIndex = -1;
    
    // Voucher No ki madad se sahi row index dhoondhna (Agar action delete ya update ho)
    if (data.action === "delete" || data.action === "update") {
      for (var i = 1; i < sheetData.length; i++) {
        // [0] ka matlab hai pehla column yaani Voucher No
        if (sheetData[i][0].toString().trim() === voucherNo.toString().trim()) {
          targetRowIndex = i + 1; // Apps Script me indexing 1 se shuru hoti hai
          break;
        }
      }
    }
    
    // ==========================================
    // 1. ACTION: DELETE
    // ==========================================
    if (data.action === "delete") {
      if (targetRowIndex !== -1) {
        sheet.deleteRow(targetRowIndex); // Sahi row ko delete karna
        return ContentService.createTextOutput(JSON.stringify({"result": "success", "message": "Deleted successfully"}))
                             .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": "Voucher not found"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // ==========================================
    // 2. ACTION: UPDATE (EDIT)
    // ==========================================
    if (data.action === "update") {
      if (targetRowIndex !== -1) {
        var updatedRowData = data.updatedRow; // Browser se aaya naya data array
        
        // Sahi row number par jaakar naya data set (overwrite) karna
        var range = sheet.getRange(targetRowIndex, 1, 1, updatedRowData.length);
        range.setValues([updatedRowData]);
        
        return ContentService.createTextOutput(JSON.stringify({"result": "success", "message": "Updated successfully"}))
                             .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": "Voucher not found"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // ==========================================
    // 3. NORMAL INSERTION (Aapka Purana Data Entry Logic)
    // ==========================================
    var tableData = data.rows;      // Table ka data array
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
  
  // JSON Output channels
  return ContentService.createTextOutput(JSON.stringify(data))
                       .setMimeType(ContentService.MimeType.JSON);
}
