function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('sales-list')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);  
  } catch(e) {
    Logger.log(e);
  }
}