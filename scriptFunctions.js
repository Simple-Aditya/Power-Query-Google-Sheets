function insertNewSheet(sheetName = "New Sheet") {
  let name = sheetName;
  try{
    const existingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!existingSheet) {
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
      return name;
    }
  }
  catch (err){
    Logger.log(`Error checking for sheet: ${err}`);
  }
  let counter = 1;
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const sheetNames = allSheets.map(sheet => sheet.getName());
  while (sheetNames.includes(`${name}${counter}`)) {
    counter++;
  }
  SpreadsheetApp.getActiveSpreadsheet().insertSheet(`${name}${counter}`);
  return `${name}${counter}`;
}

function takeNumricInput(title, text){
  const ui = SpreadsheetApp.getUi();
  let userInput = "";
  while(true){
    let response = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.CANCEL){
      return null;
    }
    let input = response.getResponseText();
    if (isNaN(input) || input === "" || input === null || input <= 0){
      ui.alert("Invalid input. Please enter a valid number.");
    }
    else {
      userInput = input;
      break;
    }
  }
  return parseInt(userInput);
}

function keepTopNRows(n) {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.keepRows(n, "top");
  } catch (error) {
    Logger.log(`Error in keepTopNRows: ${error}`);
    throw error;
  }
}

function keepBottomNRows(n) {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.keepRows(n, "last");
  } catch (error) {
    Logger.log(`Error in keepBottomNRows: ${error}`);
    throw error;
  }
}

function removeTopNRows(n) {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.removeRows(n, "top");
  } catch (error) {
    Logger.log(`Error in removeTopNRows: ${error}`);
    throw error;
  }
}

function removeBottomNRows(n) {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.removeRows(n, "last");
  } catch (error) {
    Logger.log(`Error in removeBottomNRows: ${error}`);
    throw error;
  }
}

function removeBlankRows() {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.removeBlankRows();
  } catch (error) {
    Logger.log(`Error in removeBlankRows: ${error}`);
    throw error;
  }
}

function removeDuplicateRows() {
  try {
    // This function is missing implementation
    const ui = SpreadsheetApp.getUi();
    ui.alert("Remove duplicate rows functionality is not implemented yet.");
  } catch (error) {
    Logger.log(`Error in removeDuplicateRows: ${error}`);
    throw error;
  }
}

function downFill(){
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.fillDown();
  } catch (error) {
    Logger.log(`Error in downFill: ${error}`);
    throw error;
  }
}
function upFill(){
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.fillUp();
  } catch (error) {
    Logger.log(`Error in upFill: ${error}`);
    throw error;
  }
}

function setColumnProfiles(){
  let dataTable = new table(SpreadsheetApp);
  dataTable.columnProfile();
}

function reverseTableRows() {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.reverseRows();
  } catch (error) {
    Logger.log(`Error in reverseTableRows: ${error}`);
    throw error;
  }
}

function countRows() {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.countRows();
  } catch (error) {
    Logger.log(`Error in countRows: ${error}`);
    throw error;
  }
}

function lowerCase() {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.lowerCase();
  } catch (error) {
    Logger.log(`Error in lowerCase: ${error}`);
    throw error;
  }
}

function upperCase() {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.upperCase();
  } catch (error) {
    Logger.log(`Error in upperCase: ${error}`);
    throw error;
  }
}

function capitalize() {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.capitalize();
  } catch (error) {
    Logger.log(`Error in capitalize: ${error}`);
    throw error;
  }
}

function trim() {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.trim();
  } catch (error) {
    Logger.log(`Error in trim: ${error}`);
    throw error;
  }
}

function addSuffix() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Add Suffix", "Enter the suffix to add:", ui.ButtonSet.OK_CANCEL);
    if(response.getSelectedButton() === ui.Button.OK){
      let dataTable = new table(SpreadsheetApp);
      dataTable.addCharacters("", response.getResponseText());
    }
  } catch (error) {
    Logger.log(`Error in addSuffix: ${error}`);
    throw error;
  }
}

function addPrefix() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Add Prefix", "Enter the prefix to add:", ui.ButtonSet.OK_CANCEL);
    Logger.log(`received response: ${response.getResponseText()}`);
    if(response.getSelectedButton() === ui.Button.OK){
      let dataTable = new table(SpreadsheetApp);
      dataTable.addCharacters(response.getResponseText(), "");
    }
  } catch (error) {
    Logger.log(`Error in addPrefix: ${error}`);
    throw error;
  }
}

function extractLength(){
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.length();
  } catch (error) {
    Logger.log(`Error in extractLength: ${error}`);
    throw error;
  }
}

function extractFirstCharacters(){
  try {
    let title = "Extract First Characters";
    let text = "Enter the number of characters to extract:";
    let userInput = takeNumricInput(title, text);
    if (userInput === null) return;
    
    let dataTable = new table(SpreadsheetApp);
    dataTable.extractText("first", userInput);
  } catch (error) {
    Logger.log(`Error in extractFirstCharacters: ${error}`);
    throw error;
  }
}

function extractLastCharacters(){
  try {
    let title = "Extract Last Characters";
    let text = "Enter the number of characters to extract:";
    let userInput = takeNumricInput(title, text);
    if (userInput === null) return;
    
    let dataTable = new table(SpreadsheetApp);
    dataTable.extractText("last", 0, userInput);
  } catch (error) {
    Logger.log(`Error in extractLastCharacters: ${error}`);
    throw error;
  }
}

function extractTextBeforeDelimiter() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Extract Text Before Delimiter", "Enter the delimiter:", ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      let delimiter = response.getResponseText();
      if (delimiter === "") {
        ui.alert("Delimiter cannot be empty.");
        return;
      }
      let dataTable = new table(SpreadsheetApp);
      dataTable.extractTextDelimiter("before", delimiter);
    }
  } catch (error) {
    Logger.log(`Error in extractTextBeforeDelimiter: ${error}`);
    throw error;
  }
}

function extractTextAfterDelimiter() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Extract Text After Delimiter", "Enter the delimiter:", ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      let delimiter = response.getResponseText();
      if (delimiter === "") {
        ui.alert("Delimiter cannot be empty.");
        return;
      }
      let dataTable = new table(SpreadsheetApp);
      dataTable.extractTextDelimiter("after", delimiter);
    }
  } catch (error) {
    Logger.log(`Error in extractTextAfterDelimiter: ${error}`);
    throw error;
  }
}

function splitColumn(useCase) {
  try {
    let dataTable = new table(SpreadsheetApp);
    const ui = SpreadsheetApp.getUi();

    if(useCase === "delim"){
      const response = ui.prompt("Split Column by Delimiter", "Enter the delimiter:", ui.ButtonSet.OK_CANCEL);
      if (response.getSelectedButton() === ui.Button.OK) {
        let delimiter = response.getResponseText();
        if (delimiter === "") {
          ui.alert("Delimiter cannot be empty.");
          return;
        }
        dataTable.splitColByDelimiter(delimiter);
      }
    }
    else if(useCase === "TexttoNum"){
      dataTable.splitColTexttoNum();
    }
    else if(useCase === "NumtoText"){
      dataTable.splitColNumtoText();
    } 
    else if(useCase === "UptoLow"){
      dataTable.splitColUppertoLower();
    }
    else if(useCase === "LowtoUp"){
      dataTable.splitColLowertoUpper();
    }
    else {
      ui.alert("Invalid use case provided for splitting column.");
      return;
    }
  } catch (error) {
    Logger.log(`Error in splitColumn: ${error}`);
    throw error;
  }
}

function splitColumnByCharacters() {
  try {
    let title = "Split Column by Characters";
    let text = "Enter the number of characters to split by:";
    let userInput = takeNumricInput(title, text);
    if (userInput === null) return;
    
    let dataTable = new table(SpreadsheetApp);
    dataTable.splitColByCharacters(userInput);
  } catch (error) {
    Logger.log(`Error in splitColumnByCharacters: ${error}`);
    throw error;
  }
}

function mergeColumns() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Merge Columns", "Enter the delimiter to use for merging:", ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      let delimiter = response.getResponseText();
      let dataTable = new table(SpreadsheetApp);
      dataTable.mergeCols(delimiter);
    }
  } catch (error) {
    Logger.log(`Error in mergeColumns: ${error}`);
    throw error;
  }
}

function replaceValuesInSheet(oldValue, newValue, firstOccurrence) {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.replaceValue(oldValue, newValue, firstOccurrence);
  } catch (error) {
    Logger.log(`Error in replaceValuesInSheet: ${error}`);
    throw error;
  }
}

function replacePatternInSheet(pattern, newValue, firstOccurrence) {
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.replacePattern(pattern, newValue, firstOccurrence);
  } catch (error) {
    Logger.log(`Error in replacePatternInSheet: ${error}`);
    throw error;
  }
}

function transposeData(){
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.transpose();
  } catch (error) {
    Logger.log(`Error in transposeData: ${error}`);
    throw error;
  }
}

function index(){
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.insertIndex();
  } catch (error) {
    Logger.log(`Error in index: ${error}`);
    throw error;
  }
}

function makeFirstRowHeaders(){
  try {
    let dataTable = new table(SpreadsheetApp);
    dataTable.makeFirstRowHeaders();
  } catch (error) {
    Logger.log(`Error in makeFirstRowHeaders: ${error}`);
    throw error;
  }
}

function recordExecution(){
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Recording Execution", "Please enter a name for the flow:", ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      CURRENT_FLOW = FLOW_MANAGER.createFlow(response.getResponseText());
    }
  }
  catch (error) {
    Logger.log(`Error in recordExecution: ${error}`);
    throw error;
  }
}

function undoLastStep(){
  try {
    if(!CURRENT_FLOW || !FLOW_MANAGER.flow[CURRENT_FLOW] || FLOW_MANAGER.flow[CURRENT_FLOW].stepsCount === 0){
      Logger.log("No steps to undo.");
      return;
    }
    FLOW_MANAGER.deleteStep(CURRENT_FLOW);
  }
  catch (error) {
    Logger.log(`Error in undoLastStep: ${error}`);
    throw error;
  }
}

function saveExecution(){
  try {
    if(!CURRENT_FLOW || !FLOW_MANAGER.flow[CURRENT_FLOW] || FLOW_MANAGER.flow[CURRENT_FLOW].stepsCount === 0){
      Logger.log("No steps to save.");
      return;
    }
    FLOW_MANAGER.saveFlow(CURRENT_FLOW);
  }
  catch (error) {
    Logger.log(`Error in saveExecution: ${error}`);
    throw error;
  }
}

function showSteps(){
  return FLOW_MANAGER.showFlowSteps(CURRENT_FLOW);
}