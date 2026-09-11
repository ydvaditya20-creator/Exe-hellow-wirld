// ⚠️ APNI GOOGLE SHEET KI ID YAHAN DALEIN. https://docs.google.com/spreadsheets/u/0/1qAG5pvjpj0iLEi0RPjEMC_kV01bceCY7sBNys7-xitE /htmlview#gid=63118773
const SPREADSHEET_ID = "1qAG5pvjpj0iLEi0RPjEMC_kV01bceCY7sBNys7";

// Isse mobile par bilkul sahi aur fixed sheet target hogi
function getTargetSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Jab pehli baar chalega, yeh function automatic sabhi Tabs aur Sample Data bana dega
function setupDatabase() {
  var ss = getTargetSpreadsheet();
  
  var defaultParticulars = [
    ["cash_receipt", "cash_payment", "bank_payment", "bank_receipt", "journal_voucher"],
    ["School Fees Deposit", "Office Stationery Purchase", "Staff Salary Payout via NEFT", "Online Fees Transfer", "Being Salary Expense Recorded"],
    ["Bus Fees Receipt", "Tea & Refreshments", "Electricity Bill Online", "Bank Interest Received", "Being Office Expenses Recorded"],
    ["Miscellaneous Income Receipt", "Local Conveyance Expenses", "Rent Payment via Cheque", "Grant Received in Bank", "Being Electricity Expenses Recorded"]
  ];
  
  var defaultHeads = [
    ["cash_receipt", "cash_payment", "bank_payment", "bank_receipt", "journal_voucher"],
    ["Tuition Fee A/c", "Printing & Stationery A/c", "Salary & Wages A/c", "Tuition Fee A/c", "Tuition Fee A/c"],
    ["Admission Fee A/c", "Office Expense A/c", "Electricity A/c", "Bank Interest A/c", "Office Expense A/c"],
    ["General Income A/c", "Conveyance A/c", "Rent A/c", "Donations A/c", "Salary & Wages A/c"]
  ];

  var partSheet = ss.getSheetByName("Particulars");
  if (!partSheet) {
    partSheet = ss.insertSheet("Particulars");
    partSheet.getRange(1, 1, defaultParticulars.length, defaultParticulars[0].length).setValues(defaultParticulars);
    partSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");
  }

  var headsSheet = ss.getSheetByName("Heads");
  if (!headsSheet) {
    headsSheet = ss.insertSheet("Heads");
    headsSheet.getRange(1, 1, defaultHeads.length, defaultHeads[0].length).setValues(defaultHeads);
    headsSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");
  }

  var sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1) { ss.deleteSheet(sheet1); }
}

// GET API: Mobile/Web se data padhne ke liye
function doGet() {
  setupDatabase(); // Ensure tabs are created
  
  var ss = getTargetSpreadsheet();
  var result = { particulars: {}, heads: {} };
  
  // Particulars Read
  var partSheet = ss.getSheetByName("Particulars");
  var partData = partSheet.getDataRange().getValues();
  var partHeaders = partData[0];
  for (var i = 0; i < partHeaders.length; i++) {
    result.particulars[partHeaders[i]] = [];
    for (var j = 1; j < partData.length; j++) {
      if (partData[j][i] !== "") result.particulars[partHeaders[i]].push(partData[j][i]);
    }
  }
  
  // Heads Read
  var headsSheet = ss.getSheetByName("Heads");
  var headsData = headsSheet.getDataRange().getValues();
  var headsHeaders = headsData[0];
  for (var i = 0; i < headsHeaders.length; i++) {
    result.heads[headsHeaders[i]] = [];
    for (var j = 1; j < headsData.length; j++) {
      if (headsData[j][i] !== "") result.heads[headsHeaders[i]].push(headsData[j][i]);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST API: Mobile/Web se naya data jodne ke liye
function doPost(e) {
  try {
    var ss = getTargetSpreadsheet();
    var params = JSON.parse(e.postData.contents);
    
    var sheetName = params.category;     // "Particulars" ya "Heads"
    var voucherType = params.voucherType; // e.g., "cash_receipt"
    var newValue = params.newValue;
    
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Sheet tab not found"})).setMimeType(ContentService.MimeType.JSON);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIndex = headers.indexOf(voucherType) + 1;
    
    if (colIndex > 0) {
      var lastRow = sheet.getLastRow();
      var nextRow = 1;
      
      for (var i = 1; i <= lastRow; i++) {
        if (sheet.getRange(i, colIndex).getValue() !== "") {
          nextRow = i + 1;
        }
      }
      sheet.getRange(nextRow, colIndex).setValue(newValue);
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Voucher type column not found"})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
