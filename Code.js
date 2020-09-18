function onOpen(e) {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('sales-list');
    htmlOutput.setTitle("営業メール顧客リスト入力フォーム");
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    SpreadsheetApp.flush();

  } catch(e) {
    Logger.log(e);
  }
}

function onSelectionChange(e) {
  var salesCustomerFormTitle = "DEPO営業メール";
  var projectFormTitle = "DEPO顧客リスト";
  var reactedCustomerFormTitle = "DEPO_反応のあったクライアント";
  

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var sheetName = activeSheet.getName();

  try {
    if(sheetName === salesCustomerFormTitle) {
      openSalesSidebar();
    } else if(sheetName === projectFormTitle) {
      openCustomerSidebar();
    } else if(sheetName === reactedCustomerFormTitle) {
      openPendingCustomerSidebar();
    } else {
      Logger.log("else");
      return;
    }
  } catch(e) {
    Logger.log(e);
  } finally {
    SpreadsheetApp.flush();
  }
}

function openSalesSidebar() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('sales-list');
    htmlOutput.setTitle("営業メール");
    insertImage();
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log(e);
  }
}

function openCustomerSidebar() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('customer-list');
    htmlOutput.setTitle("顧客リスト");
    insertImage();
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log(e);
  }
}

function openPendingCustomerSidebar() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('reacted-customer-list');
    htmlOutput.setTitle("反応のあったクライアント");
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log(e);
  }
}

function insertImage() {
  try {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    var img = DriveApp.getFileById("1ksL7NVikTeYIb1TsI3wrRwmxZMP5e0oJ").getThumbnail();

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト");
    activeSheet.setActiveSheet(spreadsheet);
    activeSheet.insertImage(img, 4, 2).setHeight(45).setWidth(100);
    var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange('D2');
    activeSheet.setCurrentCell(cell).setFormula('=HYPERLINK("https://drive.google.com/drive/u/0/folders/0B_loR2s0kRusflEzaHEwaDB3UWF0TlBkU1lKRWVNWGVaZTBtQnpjR3FoN1BMSk5aXzhzMGs","Driveのフォルダ")');;
  } catch(e) {
    Logger.log(e);
  }
}