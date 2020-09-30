function onOpen(e) {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('sales-list');
    htmlOutput.setTitle("営業メール");
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    SpreadsheetApp.flush();
    // var _activeSheet = document.querySelector('.goog-inline-block.docs-sheet-tab.docs-material.docs-sheet-active-tab');
    Logger.log(document);
  } catch(e) {
    Logger.log(e);
  }
}

// function onSelectionChange(e) {
//   Logger.log(e);
//   var salesCustomerFormTitle = "DEPO営業メール";
//   var projectFormTitle = "DEPO顧客リスト";
//   var reactedCustomerFormTitle = "DEPO_反応のあったクライアント";
  

//   var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
//   var sheetName = activeSheet.getName();

//   try {
//     if(sheetName === salesCustomerFormTitle) {
//       openSalesSidebar();
//     } else if(sheetName === projectFormTitle) {
//       openCustomerSidebar();
//     } else if(sheetName === reactedCustomerFormTitle) {
//       openPendingCustomerSidebar();
//     } else {
//       Logger.log("else");
//       return;
//     }
//   } catch(e) {
//     Logger.log(e);
//   } finally {
//     SpreadsheetApp.flush();
//   }
// }

function onEdit() {
  var salesCustomerFormTitle = "DEPO営業メール";
  var projectFormTitle = "DEPO顧客リスト";
  var reactedCustomerFormTitle = "DEPO_反応のあったクライアント";
  

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
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

function setFormForSalesList(e) {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('sales-list');
    htmlOutput.setTitle("営業メール");
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    SpreadsheetApp.flush();
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO営業メール");
    activeSheet.setActiveSheet(spreadsheet);
  } catch(e) {
    Logger.log(e);
  }
}

function openSalesSidebar() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('sales-list');
    htmlOutput.setTitle("営業メール");
    // insertImage();
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log(e);
  }
}

function checkSheetPosition() {
  Logger.log(SpreadsheetApp.getActiveSpreadsheet().getName());
  Logger.log(SpreadsheetApp.getActiveSheet().getName());

}

function openCustomerSidebar() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('customer-list');
    htmlOutput.setTitle("顧客リスト");
    // insertImage();
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

function showCompleteDisplay() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('comp');
  htmlOutput.setTitle("送信完了しました。");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  SpreadsheetApp.flush();

}

function insertCorporateDataWithImage() {
  try {    
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト");
    activeSheet.setActiveSheet(spreadsheet);


    var activeFolders = DriveApp.getFolders();

    // 案件別のフォルダを格納
    var activeFoldersArr = [];
    while(activeFolders.hasNext()) {
      var folder = activeFolders.next();
      if(folder.getName().match(/^DEPO_ヒアリングシート_.+/) !== null) {
        activeFoldersArr.push(folder);
      }
    }
    for(var i = 0; i < activeFoldersArr.length; i++) {
      // 案件フォルダごとにpngもしくはjpgの画像を取得
      var photoWithJPG = activeFoldersArr[i].getFilesByType(MimeType.JPEG);
      var photoWithPNG = activeFoldersArr[i].getFilesByType(MimeType.PNG);
      // 画像挿入処理
      if (photoWithJPG.hasNext()) {
        var imgjpg = photoWithJPG.next().getThumbnail();
        activeSheet.insertImage(imgjpg, 4, 2+i).setHeight(100).setWidth(100);
      } else if(photoWithPNG.hasNext()) {
        var imgpng = photoWithPNG.next().getThumbnail();
        activeSheet.insertImage(imgpng, 4, 2+i).setHeight(100).setWidth(100);
      }
      // 該当のスプレッドシートの1番上の行に企業名
      var cellPositionA = 'A'+(2+i).toString();
      var targetCellForName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange(cellPositionA);

      if (targetCellForName.getValue() === "") {
        var folderName = activeFoldersArr[i].getName();
        var corporateName = folderName.match(/DEPO_(.+)/)[1];
        activeSheet.setCurrentCell(targetCellForName).setValue(corporateName);  
      }
      
      // 該当のスプレッドシートの1番上の行にステータス
      var cellPositionB = 'B'+(2+i).toString();
      var targetCellForName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange(cellPositionB);

      if (targetCellForName.getValue() === "") {
        var folderName = activeFoldersArr[i].getName();
        var status = '進行中';
        activeSheet.setCurrentCell(targetCellForName).setValue(status);  
      }

      // 該当のスプレッドシートの1番上の行に日付
      var cellPositionC = 'C'+(2+i).toString();
      var targetCellForName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange(cellPositionC);

      if (targetCellForName.getValue() === "") {
        var nowDate = new Date(Date.now());
        var _nowDate = new Date(Date.now());
        var oneMonthAfterDate = new Date(_nowDate.setMonth(_nowDate.getMonth() + 1));
        activeSheet.setCurrentCell(targetCellForName).setValue(nowDate.getFullYear()+'/'+nowDate.getMonth()+'/'+nowDate.getDate()+'〜'+ oneMonthAfterDate.getFullYear()+'/'+oneMonthAfterDate.getMonth()+'/'+oneMonthAfterDate.getDate());  
      }
      
      // 該当のスプレッドシートの1番上の行に画像挿入
      var cellPositionD = 'D'+(2+i).toString();
      var targetCellForImage = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange(cellPositionD);

      Logger.log(targetCellForName.getValue());
      var formulaStr = '=HYPERLINK("https://drive.google.com/drive/u/0/folders/'+activeFoldersArr[i].getId()+'","Driveのフォルダ")';
      activeSheet.setCurrentCell(targetCellForImage).setFormula(formulaStr);  

      // 該当のスプレッドシートの1番上の行に初期値（案件名）挿入
      var cellPositionE = 'E'+(2+i).toString();
      var targetCellForImage = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange(cellPositionE);

      Logger.log('test '+targetCellForName.getValue());

      if (targetCellForName.getValue() === "") {
        var projectTitle = '案件名';
        activeSheet.setCurrentCell(targetCellForImage).setValue(projectTitle);  
      }

      // 該当のスプレッドシートの1番上の行に初期値（Slack URL）挿入
      var cellPositionF = 'F'+(2+i).toString();
      var targetCellForImage = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange(cellPositionF);

      if (targetCellForName.getValue() === "") {
        var urlStr = 'deposuito.slack.com';
        activeSheet.setCurrentCell(targetCellForImage).setValue(urlStr);  
      }

      // 該当のスプレッドシートの1番上の行に初期値（Slack ID/PASS）挿入
      var cellPositionG = 'G'+(2+i).toString();
      var targetCellForImage = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト").getRange(cellPositionG);

      if (targetCellForName.getValue() === "") {
        var slackIdPass = 'SlackのID/PASSを入力してください';
        activeSheet.setCurrentCell(targetCellForImage).setValue(slackIdPass);  
      }
    }
  } catch(e) {
    Logger.log(e);
  }
}

function setStatusColor() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEPO顧客リスト");
  var datas = spreadsheet.getRange(2, 2, spreadsheet.getLastRow() - 1);
  // Logger.log(spreadsheet.getRange(2, 2).getValue());

  for(var i = 2; i < 3; i++) {
    for(var j = 2; j <= datas.getNumRows(); j++) {
      var cellValue = spreadsheet.getRange(j, i).getValue();
      // Logger.log(spreadsheet.getRange(j, i));
      // Logger.log(cellValue);
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
      var splittedDate_1 = date.split('~')[1];
      var splittedDate_0 = date.split('~')[0];
      var limitMonthToDate_1 = new Date(Date.parse(splittedDate_1));
      var limitMonthToDate_0 = new Date(Date.parse(splittedDate_0));
      var nowDate = new Date(Date.now());
      // Logger.log(limitMonthToDate.getMonth() + 1 - nowDate.getMonth() + 1 <= 1);
      // Logger.log(splittedDate);
      // Logger.log(limitMonthToDate);
      // Logger.log(nowDate.getMonth() + 1);

      try {
        // Logger.log(limitMonthToDate.getMonth() + 1 - nowDate.getMonth() + 1);
        // Logger.log(limitMonthToDate.getMonth() + 1);
        // Logger.log(nowDate.getMonth() + 1);  
        Logger.log(limitMonthToDate_0);
        Logger.log('1 '+limitMonthToDate_0.getMonth());
        Logger.log('2 '+limitMonthToDate_0.getMonth() + 1 - limitMonthToDate_1.getMonth() + 1 <= 1);
        Logger.log('3 '+limitMonthToDate_1.getMonth() + 1 == 12);
        if(limitMonthToDate_1.getMonth() + 1 - nowDate.getMonth() + 1 <= 1) {
          spreadsheet.getRange(j, i).setBackground("red");
        } else if(limitMonthToDate_1.getMonth() + 1 == 12) {
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

function insertSalesFormData(form) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const values = [
      [form.company_name,form.url,form.mail,form.category,form.charge_person]
    ];
    const numRows = values.length;
    const numColumns = values[0].length;
    sheet.insertRows(2,numRows);
    sheet.getRange(2, 1, numRows, numColumns).setValues(values);

  } catch(e) {
    Logger.log(e);
  }
}

function insertCustomerFormData(form) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const values = [
      [form.name,form.introduction_source,form.drive_link,form.slack_id,form.slack_pass,form.contract_period,form.category,form.company_charge_person,form.charge_person,form.project_title,form.printing_requirement_and_start_date,form.printing_specification]
    ];
    const numRows = values.length;
    const numColumns = values[0].length;
    sheet.insertRows(2,numRows);
    sheet.getRange(2, 1, numRows, numColumns).setValues(values);

  } catch(e) {
    Logger.log(e);
  }
}

function insertReactedCustomerFormData(form) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const values = [
      [form.company_name,form.introduction_source,form.contact_date,form.mtg_day,form.charge_person]
    ];
    const numRows = values.length;
    const numColumns = values[0].length;
    sheet.insertRows(2,numRows);
    sheet.getRange(2, 1, numRows, numColumns).setValues(values);

  } catch(e) {
    Logger.log(e);
  }
}

function doPost() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('comp');
  return htmlOutput;
}

function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile("comp");

  } catch(ee) {
    Logger.log(ee);
  }
}