function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Power Query")
    .addItem("Get Data", "getData")
    .addItem("Data Cleaning", "headerConfirmation")
    .addToUi();
}

function headerConfirmation(){
  const ui = SpreadsheetApp.getUi();
  const userResponse = ui.alert("Header Confirmation", "Does your data have headers", ui.ButtonSet.YES_NO);
  if(userResponse === ui.Button.NO){
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.insertRowBefore(1);
    var userHeaders = [];
    for(let i = 1; i <= sheet.getLastColumn(); i++){
      userHeaders.push(`Column${i}`);
    }
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setValues([userHeaders]);
  }
  let dataTable = new table(SpreadsheetApp);
  dataTable.makeFirstRowHeaders();
  dataCleaningSideBar();
}

function dataCleaningSideBar(){
  var sideBar = HtmlService.createHtmlOutputFromFile("index.html");
  sideBar.setTitle("Data Cleaning")
  SpreadsheetApp.getUi().showSidebar(sideBar);
}

function showSubFeature(featureName){
  var html = HtmlService.createHtmlOutputFromFile(`${featureName}.html`);
  html.setTitle(`Manage ${featureName}`)
  SpreadsheetApp.getUi().showSidebar(html);
}
