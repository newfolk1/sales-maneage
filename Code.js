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

function setStatusColor() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト");
  var datas = spreadsheet.getRange(2, 2, spreadsheet.getLastRow() - 1);

  Logger.log(spreadsheet.getRange(2, 2).getValue());

  for(var i = 2; i < 3; i++) {
    for(var j = 2; j <= datas.getNumRows(); j++) {
      var cellValue = spreadsheet.getRange(j, i).getValue();
      Logger.log(spreadsheet.getRange(j, i));
      Logger.log(cellValue);
      if(cellValue == "解約") {
        spreadsheet.getRange(j, i).setBackground("gray");
      } else if(cellValue == "進行中") {
        spreadsheet.getRange(j, i).setBackground("white");      
      }
    }
  }
}


function setDeadLineColor() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト");
  var datas = spreadsheet.getRange(2, 3, spreadsheet.getLastRow() - 1);

  Logger.log(spreadsheet.getRange(2, 3).getValue());

  for(var i = 3; i < 4; i++) {
    for(var j = 2; j <= datas.getNumRows(); j++) {
      var date = spreadsheet.getRange(j, i).getValue();
      var splittedDate = date.split('~')[1];
      var limitMonthToDate = new Date(Date.parse(splittedDate));
      var nowDate = new Date(Date.now());
      // Logger.log(limitMonthToDate.getMonth() + 1 - nowDate.getMonth() + 1 <= 1);
      // Logger.log(splittedDate);
      // Logger.log(limitMonthToDate);
      // Logger.log(nowDate.getMonth() + 1);

      try {
        Logger.log(limitMonthToDate.getMonth() + 1 - nowDate.getMonth() + 1);
        Logger.log(limitMonthToDate.getMonth() + 1);
        Logger.log(nowDate.getMonth() + 1);
        if(limitMonthToDate.getMonth() + 1 - nowDate.getMonth() + 1 <= 1) {
          spreadsheet.getRange(j, i).setBackground("red");
        } else if(limitMonthToDate.getMonth() + 1 == 12) {
          var nokottatime = limitMonthToDate - nowDate;
          var nokottatimeMonth = new Date(nokottatime).getMonth() + 1;
          if (nokottatimeMonth <= 1){
            spreadsheet.getRange(j, i).setBackground("red");
          }
        }
  
      } catch(e) {
        Logger.log(e);
      }
    }
  }
}

function doPost(e) {
  try {
    Logger.log(params);

    var params = JSON.parse(e.postData.getDataAsString());  // ※
    var value = params.value;
    Logger.log(value);
    //var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   

    // var postdata = JSON.parse(e.postData.getDataAsString()); 
    // var name = postdata.parameters.name;
    // var url = postdata.parameters.url;
    // var mail = postdata.parameters.mail;
    // var category = postdata.parameters.category;
    // var charge = postdata.parameters.charge;
    // var arr = [name, url, mail, category, charge];
    // spreadsheet.appendRow(arr);
    SpreadsheetApp.flush();
  } catch(ee) {
    Logger.log(ee);
  }
}