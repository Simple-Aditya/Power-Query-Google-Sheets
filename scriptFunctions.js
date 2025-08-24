function insertNewSheet(sheetName = "New Sheet") {
  let name = sheetName;
  try{
    const existingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!existingSheet) {
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
      return name;
    }
  }
  catch (err){
    Logger.log(`Error checking for sheet: ${err}`);
  }
  let counter = 1;
  while (true) {
    try{
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${name}${counter}`);
      if (!sheet) {
        break;
      }
      counter++;
    }
    catch (err){
      Logger.log(`Error checking for sheet: ${err}`);
      break; 
    }
  }
  const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`${name}${counter}`);
  return `${name}${counter}`;
}

function takeNumricInput(title, text){
  const ui = SpreadsheetApp.getUi();
  let userInput = "";
  while(true){
    let response = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
    let input = response.getResponseText();
    if (isNaN(input) || input === "" || input === null || input <= 0){
      const ui = SpreadsheetApp.getUi();
      ui.alert("Invalid input. Please enter a valid number.");
      return;
    }
    else if (response.getSelectedButton() === ui.Button.CANCEL){
      return;
    }
    else {
      userInput = input;
      break;
    }
  }
  return parseInt(userInput);
}

function keepTopNRows(n) {
  let dataTable = new table(SpreadsheetApp);
  dataTable.keepRows(n, "top");
}

function keepBottomNRows(n) {
  let dataTable = new table(SpreadsheetApp);
  dataTable.keepRows(n, "last");
}

function removeTopNRows(n) {
  let dataTable = new table(SpreadsheetApp);
  dataTable.removeRows(n, "top");
}

function removeBottomNRows(n) {
  let dataTable = new table(SpreadsheetApp);
  dataTable.removeRows(n, "last");
}

function removeBlankRows() {
  let dataTable = new table(SpreadsheetApp);
  dataTable.removeBlankRows();
}

function downFill(){
  let dataTable = new table(SpreadsheetApp);
  dataTable.fillDown();
}
function upFill(){
  let dataTable = new table(SpreadsheetApp);
  dataTable.fillUp();
}

function reverseRows() {
  let dataTable = new table(SpreadsheetApp);
  dataTable.reverseRows();
}

function countRows() {
  let dataTable = new table(SpreadsheetApp);
  dataTable.countRows();
}

function lowerCase() {
  let dataTable = new table(SpreadsheetApp);
  dataTable.lowerCase();
}

function upperCase() {
  let dataTable = new table(SpreadsheetApp);
  dataTable.upperCase();
}

function capitalize() {
  let dataTable = new table(SpreadsheetApp);
  dataTable.capitalize();
}

function trim() {
  let dataTable = new table(SpreadsheetApp);
  dataTable.trim();
}

function addSuffix() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Add Suffix", "Enter the suffix to add:", ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton() === ui.Button.OK){
    let dataTable = new table(SpreadsheetApp);
    dataTable.addCharacters("", response.getResponseText());
  }
  return;
} 

function addPrefix() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Add Prefix", "Enter the prefix to add:", ui.ButtonSet.OK_CANCEL);
  Logger.log(`received response: ${response.getResponseText()}`);
  if(response.getSelectedButton() === ui.Button.OK){
    let dataTable = new table(SpreadsheetApp);
    dataTable.addCharacters(response.getResponseText(), "");
  }
  return;
} 

function extractLength(){
  let dataTable = new table(SpreadsheetApp);
  dataTable.length();
}

function extractFirstCharacters(){
  let title = "Extract First Characters";
  let text = "Enter the number of characters to extract:";
  let userInput = takeNumricInput(title, text);
  let dataTable = new table(SpreadsheetApp);
  dataTable.extractText("first", userInput);
}

function extractLastCharacters(){
  let title = "Extract Last Characters";
  let text = "Enter the number of characters to extract:";
  let userInput = takeNumricInput(title, text);
  let dataTable = new table(SpreadsheetApp);
  dataTable.extractText("last", 0, userInput);
}

function extractTextBeforeDelimiter() {
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
}

function extractTextAfterDelimiter() {
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
}

function splitColumn(useCase) {
  let dataTable = new table(SpreadsheetApp);

  if(useCase === "delim"){
    const ui = SpreadsheetApp.getUi();
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
      ui.alert("Invalid use case provided for splitting column by delimiter.");
      return;
    }
}

function splitColumnByCharacters() {
  let title = "Split Column by Characters";
  let text = "Enter the number of characters to split by:";
  let userInput = takeNumricInput(title, text);
  let dataTable = new table(SpreadsheetApp);
  dataTable.splitColByCharacters(userInput);
}

function mergeColumns() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Merge Columns", "Enter the delimiter to use for merging:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    let delimiter = response.getResponseText();
    let dataTable = new table(SpreadsheetApp);
    dataTable.mergeCols(delimiter);
  }
}

// function used to show the input box
function replaceValuesDialog(){
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile("multiInput.html")
    .setWidth(300)
    .setHeight(200);
  ui.showModalDialog(html, "Replace Values");
}

//function to handle the replacement of values
function replaceValuesInSheet(oldValue, newValue) {
  let dataTable = new table(SpreadsheetApp);
  dataTable.newValue(oldValue, newValue);
}

function transposeData(){
  let dataTable = new table(SpreadsheetApp);
  dataTable.transpose();
}

function index(){
  let dataTable = new table(SpreadsheetApp);
  dataTable.insertIndex();
}

function makeFirstRowHeaders(){
  let dataTable = new table(SpreadsheetApp);
  dataTable.makeFirstRowHeaders();
}