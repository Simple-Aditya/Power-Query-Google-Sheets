function getData(){
  var sideBar = HtmlService.createHtmlOutputFromFile("getDataDialog.html");
  sideBar.setTitle("Get Data")
  SpreadsheetApp.getUi().showSidebar(sideBar);
}

function processLinks(links, delimiter, location) {
  try {
    if(!links){
        SpreadsheetApp.getUi().alert("Please provide drive or sheet links.");
        return;
    }
    const urlArray = links.split(delimiter).map(link => link.trim()).filter(link => link);
    if(urlArray.length === 0){
        SpreadsheetApp.getUi().alert("No valid links found.");
        return;
    }
    let inValidUrl = [];
    let spreadSheets = 0;
    let sheets = 0;

    for(let i = 0; i < urlArray.length; i++){
      const url = urlArray[i];
      
      if (url.includes("drive.google.com/drive/folders/")) {
        const folderId = url.match(/[-\w]{25,}/);
        if (!folderId) {
          inValidUrl.push(url);
          continue;
        }
        
        const folder = DriveApp.getFolderById(folderId[0]);
        const sheetFiles = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
        while (sheetFiles.hasNext()) {
          const file = sheetFiles.next();
          const fileID = file.getId();
          const result = getDataFromLink(fileID, location);
          if(result.inValidUrl) inValidUrl.push(result.inValidUrl);
          spreadSheets += result.spreadSheets;
          sheets += result.sheets;
        }
        
        let allFiles = [];
        const excelFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
        while (excelFiles.hasNext()) {
          allFiles.push(excelFiles.next());
        }

        const excelXFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL_OPENXML);
        while (excelXFiles.hasNext()) {
          allFiles.push(excelXFiles.next());
        }

        const csvFiles = folder.getFilesByType(MimeType.CSV);
        while (csvFiles.hasNext()) {
          allFiles.push(csvFiles.next());
        }

        for (let i = 0; i < allFiles.length; i++) {
          const file = allFiles[i];
          const fileId = file.getId();
          
          let result;
          try {
            const blob = file.getBlob();
            const resource = {
              title: file.getName(),
              mimeType: MimeType.GOOGLE_SHEETS
            };
            const convertedFile = Drive.Files.insert(resource, blob);
            result = getDataFromLink(convertedFile.id, location);
            DriveApp.getFileById(convertedFile.id).setTrashed(true);
          } 
          catch (convError) {
            Logger.log(`Error converting file ${file.getName()}: ${convError.message}`);
            inValidUrl.push(file.getUrl());
            continue;
          }
          
          if (result.inValidUrl) inValidUrl.push(result.inValidUrl);
          spreadSheets += result.spreadSheets;
          sheets += result.sheets;
        }
      } 
      else {
        const result = getDataFromLink(url, location);
        if(result.inValidUrl) inValidUrl.push(result.inValidUrl);
        spreadSheets += result.spreadSheets;
        sheets += result.sheets;
      }
    }

    SpreadsheetApp.getUi().alert(`Data import completed!\n\nTotal Spreadsheets Processed: ${spreadSheets}\nTotal Sheets Imported: ${sheets}\n\nInvalid URLs:\n${inValidUrl.length ? inValidUrl.join("\n") : "None"}`);
  }
  catch (err){
    Logger.log(`Error in retrieving data: ${err}`);
    SpreadsheetApp.getUi().alert("An error occurred while retrieving data. Please check the logs for details.");
  }
}

function getDataFromLink(link, location){
  try{
    let targetSheet;
    if(location === "current"){
      targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }
    else if(typeof location === "string" && location !== "separate"){
      try{
        targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(location);
      }
      catch (error) {
        SpreadsheetApp.getUi().alert(`Sheet named "${location}" not found in the current spreadsheet.`);
        Logger.log(`Error retrieving sheet: ${error}`);
        return {
          inValidUrl: link,
          spreadSheets: 0,
          sheets: 0
        };
      }
    }
    
    let sheets = 0;
    let fileId;
    
    if (link.match(/^[-\w]{25,}$/)) {
      fileId = link;
    } 
    else {
      const idMatch = link.match(/[-\w]{25,}/);
      if(!idMatch){
        Logger.log(`Invalid URL: ${link}`);
        return {
          inValidUrl: link,
          spreadSheets: 0,
          sheets: 0
        };
      }
      fileId = idMatch[0];
    }
    
    const ss = SpreadsheetApp.openById(fileId);
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

    return {
      inValidUrl: "",
      spreadSheets: 1,
      sheets: sheets
    };
  }
  catch (error) {
    Logger.log(`Error processing URL: ${link}\n${error.message}`);
    return {
      inValidUrl: link,
      spreadSheets: 0,
      sheets: 0
    };
  }
}