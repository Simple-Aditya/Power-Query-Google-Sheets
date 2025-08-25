function getData(){
  var sideBar = HtmlService.createHtmlOutputFromFile("getDataDialog.html");
  sideBar.setTitle("Get Data")
  SpreadsheetApp.getUi().showSidebar(sideBar);
}

function processLinks(links, delimiter, location) {
  try{
    if(!links){
        SpreadsheetApp.getUi().alert("Please provide links.");
        return;
    }
    const urlArray = links.split(delimiter).map(link => link.trim()).filter(link => link);
    if(urlArray.length === 0){
        SpreadsheetApp.getUi().alert("No valid links found.");
        return;
    }
    let targetSheet;
    if(location === "current"){
      targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }
    else if(location === "specify"){
      targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(location);
    }
    let inValidUrl = [];
    let spreadSheets = 0;
    let sheets = 0;
    for(let i = 0; i < urlArray.length; i++){
        try {
            const url = urlArray[i];
            const fileId = url.match(/[-\w]{25,}/);
            if(!fileId){
                inValidUrl.push(url);
                continue;
            }
            const ss = SpreadsheetApp.openById(fileId[0]);
            const allSheets = ss.getSheets();
            for(let j = 0; j < allSheets.length; j++){
                if(allSheets[j].getLastRow() === 0) continue;
                const sheetName = allSheets[j].getName();
                const sourceSheet = ss.getSheetByName(sheetName);
                if(location === "separate"){
                  let newSheetName = insertNewSheet(sheetName);
                  targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
                }
                const data = sourceSheet.getDataRange().getValues();
                targetSheet.getRange(targetSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
                sheets++;
            }
            spreadSheets++;
        } 
        catch (error) {
            Logger.log(`Error processing URL: ${urlArray[i]}\n${error.message}`);
        }
    }
    SpreadsheetApp.getUi().alert(`Data import completed!\n\nTotal Spreadsheets Processed: ${spreadSheets}\nTotal Sheets Imported: ${sheets}\n\nInvalid URLs:\n${inValidUrl.length ? inValidUrl.join("\n") : "None"}`);
  }
  catch (err){
    Logger.log(`Error in retriving data from sheets: ${err}`);
    this._ui.alert("An error occurred while retriving data. Please check the logs for details.");
  }
}