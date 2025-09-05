function onOpen(){
  createPowerQueryMenu();
}

function createPowerQueryMenu(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Power Query");
  menu.addItem("Get Data", "getData");
  menu.addItem("Data Cleaning", "headerConfirmation");
  menu.addseparator();
  const flows = FLOW_MANAGER.getAllFlows();
  for(const flow of flows){
    if(FLOW_MANAGER.flow[flow] && FLOW_MANAGER.flow[flow].isSaved){
      menu.addItem(`Execution: ${flow}`, `executeFlow(${flow})`);
    }
  }
  menu.addToUi();
}

function executeFlow(flowName) {
  try {
    FLOW_MANAGER.executeFlow(flowName);
  } catch (error) {
    Logger.log(`Error executing flow ${flowName}: ${error}`);
  }
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
  var sideBar = HtmlService.createHtmlOutputFromFile("index.html")
    .setTitle("Data Cleaning")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(sideBar);
}

// This function acts as a bridge for the pattern replacement functionality
function replaceWithPattern(pattern, replaceWith) {
  try {
    return replacePatternInSheet(pattern, replaceWith);
  } catch (error) {
    Logger.log(`Error in replaceWithPattern: ${error}`);
    throw error;
  }
}