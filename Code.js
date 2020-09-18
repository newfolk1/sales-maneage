function onOpen(e) {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('sales-list');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch(e) {
    Logger.log(e);
  }
}

function openSidebar() {
  SpreadsheetApp.getUi().createMenu('My Menu')
  .addItem('My menu item', 'myFunction')
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('My sub-menu')
      .addItem('One sub-menu item', 'mySecondFunction')
      .addItem('Another sub-menu item', 'myThirdFunction'))
  .addToUi();
}