class table {
  constructor(spreadsheet) {
    this._sheet = spreadsheet.getActiveSheet();
    this._range = this._sheet.getRange(1, 1, this._sheet.getLastRow(), this._sheet.getLastColumn());
    this._numRows = this._sheet.getLastRow() || 1;
    this._numCols = this._sheet.getLastColumn();
    this._table = this._range.getValues();
    this._data = this._numRows > 1 ? this._table.slice(1) : [];
    this._headers = this._numRows > 0 ? this._table[0] : [];
    this._ui = spreadsheet.getUi();
    this._currentCell = this._sheet.getActiveCell();
    this._currentCol = this._currentCell.getColumn();
    this._currentColRange = this._sheet.getActiveRange().getNumColumns();
  }

  columnProfile(){
    if (this._numRows < 2) {
      Logger.log("Not enough data to analyze");
      return;
    }

    const profile = this._headers.map((header, colIndex) => {
      const colData = this._data.map(row => row[colIndex]);
      const uniqueValues = [...new Set(colData)].filter(v => v !== null && v !== "").length;
      const nullValues = [...new Set(colData)].filter(v => v === null || v === "").length;
      return {
        header,
        uniqueValues,
        nullValues
      };
    });

    const formattedProfiles = profile.map(col => {
      return `Column: ${col.header}\nUnique Values: ${col.uniqueValues}\nNull Values: ${col.nullValues}`;
    });
  
    const profileRange = this._sheet.getRange(2, 1, 1, profile.length);
    profileRange.setValues([formattedProfiles]);
    profileRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    this._sheet.setRowHeight(2, 120); 
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "columnProfile");   
    Logger.log("Set Column Profile");
  }

  reverseRows() {
    if (this._numRows <= 1) {
      Logger.log("Not enough data to reverse");
      return;
    }

    this._data = this._data.reverse();
    const bodyRange = this._sheet.getRange(2, 1, this._numRows - 1, this._numCols);
    bodyRange.clearContent();
    bodyRange.setValues(this._data);

    Logger.log("Reversed the rows");
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "reverseRows");
  }

  makeFirstRowHeaders() {
    if (this._numRows < 1) {
      Logger.log("No data to set headers");
      return;
    }
    let skipRows = 0;
    for(let i = 0; i < this._table.length; i++){
      if(this._table[i].every(cell => cell == null || cell.toString().trim() === "")){
        skipRows++;
      } else break;
    }
    if(skipRows > 0){
      this._sheet.deleteRows(1, skipRows);
      Logger.log(`Removed ${skipRows} empty rows from the top`);
      this._numRows = this._sheet.getLastRow() || 1;
      this._table = this._sheet.getRange(1, 1, this._numRows, this._sheet.getLastColumn()).getValues();
      this._data = this._table.slice(1);
      this._headers = this._table[0];
    }
    this._sheet.setFrozenRows(1);
    this._sheet.getRange(1, 1, 1, this._sheet.getLastColumn())
    .setFontWeight("bold")
    .setBackground("#f2f2f2");

    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "makeFirstRowHeaders");
  }

  countRows(){
    let newSheet = insertNewSheet("Row Count");
    newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheet);
    newSheet.getRange(1, 1, 1, 2).setValues([[`Total Rows in ${this._sheet.getName()}`, this._numRows]]);

    newSheet.setColumnWidth(1, 200);
    newSheet.setColumnWidth(2, 100);    
    newSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f2f2f2");
    newSheet.activate();

    Logger.log(`Total Rows: ${this._numRows}`);
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "countRows");
  }

  replaceValue(oldValue, newValue = "", firstOccurrence = true, column = this._currentCol, colRange = this._currentColRange) {
  try {
    if (this._numRows < 1) {
      Logger.log("No data to replace values");
      return;
    }
    
    if (!oldValue || oldValue.toString().trim() === "") {
      Logger.log("Cannot replace empty or whitespace-only values");
      this._ui.alert("Cannot replace empty or whitespace-only values");
      return;
    }
    
    let newData = [];
    let replacementCount = 0;
    
    for (let i = 0; i < this._data.length; i++) {
      let newRow = [];
      for (let j = 0; j < colRange; j++) {
        let cell = this._data[i][column + j - 1];

        if (cell != null) {
          let cellStr = cell.toString();
          if (cellStr.includes(oldValue.toString())) {
            let newCellValue = "";
            if (firstOccurrence) {
              newCellValue = cellStr.replace(oldValue.toString(), newValue.toString());
            } else {
              newCellValue = cellStr.replaceAll(oldValue.toString(), newValue.toString());
            }
            newRow.push(newCellValue);
            replacementCount++;
          } else {
            newRow.push(cell);
          }
        } else {
          newRow.push("");
        }
      }
      newData.push(newRow);
    }
    
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "replaceValue", [oldValue, newValue, firstOccurrence, this._currentCol, this._currentColRange]);

    if (replacementCount > 0) {
      this._sheet.getRange(2, column, newData.length, colRange).setValues(newData);
      this._data = newData;
      Logger.log(`Replaced "${oldValue}" with "${newValue}" in ${replacementCount} cells`);
    } 
    else {
      Logger.log(`No instances of "${oldValue}" found to replace`);
      this._ui.alert(`No instances of "${oldValue}" found in the selected range`);
    }
    
  } catch (err) {
    Logger.log(`Error replacing values: ${err}`);
    this._ui.alert("An error occurred while replacing values. Please check the logs for details.");
  }
}

replacePattern(pattern, newValue = "", firstOccurrence = true, column = this._currentCol, colRange = this._currentColRange) {
  try {
    if (this._numRows < 1) {
      Logger.log("No data to replace values");
      return;
    }
    if(!pattern || pattern.toString().trim() === ""){
      Logger.log("Cannot replace empty or whitespace-only patterns");
      this._ui.alert("Cannot replace empty or whitespace-only patterns");
      return;
    }

    let newData = [];
    let replacementCount = 0;
    for(let i = 0; i < this._data.length; i++){
      let newRow = [];
      for(let j = 0; j < colRange; j++){
        let cell = this._data[i][column + j - 1];
        if(cell != null){
          let cellStr = cell.toString();
          try {
            let regex = new RegExp(pattern);
            if(!firstOccurrence){
              regex = new RegExp(pattern, 'g');
            }
            let newCellValue = cellStr.replace(regex, newValue.toString());
            if(cellStr !== newCellValue) {
              replacementCount++;
            }
            newRow.push(newCellValue);
          } 
          catch (regexErr) {
            Logger.log(`Invalid regex pattern: ${regexErr}`);
            this._ui.alert("The provided pattern is invalid. Please correct it and try again.");
            return;
          }
        }
        else {
          newRow.push("");
        }
      }
      newData.push(newRow);
    }
    
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "replacePattern", [pattern, newValue, firstOccurrence, this._currentCol, this._currentColRange]);
    if(replacementCount > 0){
      this._sheet.getRange(2, column, newData.length, colRange).setValues(newData);
      Logger.log(`Replaced values matching regex "${pattern}" with "${newValue}"`);
      this._executionSteps.push(`Replaced values with pattern ${pattern}"`);
    }
    else {
      Logger.log(`No instances matching regex "${pattern}" found to replace`);
      this._ui.alert(`No instances matching pattern found in the selected range`);
    }
  }
  catch (err) {
    Logger.log(`Error replacing values with regex: ${err}`);
    this._ui.alert("An error occurred while replacing values with regex. Please check the logs for details.");
  }
}
  
  transpose(){
    try{
      if(this._numRows < 1 || this._numCols < 1){
        Logger.log("Insufficient data to transpose");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "transpose");
        return;
      }
      
      let newData = [];
      for(let j = 0; j < this._numCols; j++){
        let newRow = [];
        for(let i = 0; i < this._numRows; i++){
          newRow.push(this._table[i][j] || "");
        }
        newData.push(newRow);
      }
      
      this._sheet.clearContents();
      this._sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      this.makeFirstRowHeaders();
      
      Logger.log("Transposed the data");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "transpose");
    }

    catch(err){
      Logger.log(`Unable to transpose the data because ${err}`);
      this._ui.alert("An error occurred while transposing the data. Please check the logs for details")
      throw err;
    }
  }

  keepRows(rows, condition = "top") {
    if (this._numRows <= 1) {
      Logger.log(`Kept ${condition} ${rows} rows`);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "keepRows", [rows, condition]);
      return;
    }
    let newData = [];
    if (condition === "last") {
      newData = this._data.slice(this._numRows - rows, this._numRows);
    } else {
      newData = this._data.slice(0, rows);
    }

    this._sheet.getRange(2, 1, this._numRows - 1, this._sheet.getLastColumn()).clearContent();
    if (newData.length > 0) {
      this._sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
    }
    this._data = newData;
    Logger.log(`Kept ${condition} ${rows} rows`);
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "keepRows", [rows, condition]);
  }

  removeRows(rows, condition = "top") {
    if (rows <= 0) {
      this._ui.alert("Please enter a valid input.");
      return;
    }

    if (this._numRows < 1) {
      Logger.log(`Unable to remove ${condition} ${rows} rows due to insufficiency.`);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "removeRows", [rows, condition]);
      return;
    }

    let newData = [];
    if (condition === "last") {
      newData = this._data.slice(0, this._numRows - rows);
    } 
    else {
      newData = this._data.slice(rows, this._numRows);
    }
    this._sheet.getRange(2, 1, this._numRows - 1, this._sheet.getLastColumn()).clearContent();
    if (newData.length > 0) {
      this._sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
    }
    this._data = newData;
    Logger.log(`Removed ${condition} ${rows} rows.`);
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "removeRows", [rows, condition]);
  }

  removeBlankRows() {
    if (this._numRows < 1) {
      Logger.log("No Data to remove blanks from");
      return;
    }

    let cleanData = this._data.filter(row => row.some(cell => cell != null && cell.toString().trim() !== ""));
    this._sheet.getRange(2, 1, this._numRows - 1, this._sheet.getLastColumn()).clearContent();
    if (cleanData.length > 0) {
      this._sheet.getRange(2, 1, cleanData.length, cleanData[0].length).setValues(cleanData);
    }
    this._data = cleanData;
    
    Logger.log("Removed empty rows");
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "removeBlankRows");
  }

  // Column operations

  fillDown(column = this._currentCol) {
    if(this._numRows <= 1) {
      Logger.log("Not enough data to fill the data");
      return;
    }
    let newData = [];
    let lastPushed = "";
    for(let i = 0; i < this._data.length; i++){
      let cell = this._data[i][column - 1];
      if (cell != null && cell.toString().trim() !== ""){
        newData.push([cell]);
        lastPushed = cell;
      }
      else {
        newData.push([lastPushed]);
      }
    }
    this._sheet.getRange(2, column, newData.length, 1).setValues(newData);

    Logger.log("Performed down fill on the data");
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "fillDown", [this._currentCol]);
  }
  
  fillUp(column = this._currentCol) {
    if(this._numRows <= 1) {
      Logger.log("Not enough data to fill");
      return;
    }
    let newData = [];
    let lastPushed = "";
    for(let i = this._data.length - 1; i >= 0 ; i--){
      let cell = this._data[i][column - 1];
      if (cell != null && cell.toString().trim() !== ""){
        newData.push([cell]);
        lastPushed = cell;
      }
      else {
        newData.push([lastPushed]);
      }
    }
    this._sheet.getRange(2, column, newData.length, 1).setValues(newData);
    Logger.log("Performed up fill on the data");
    FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "fillUp", [this._currentCol]);
  }

  mergeCols(delimiter = "", column = this._currentCol, colRange = this._currentColRange) {
    try{
      if (this._numCols < 2) {
        Logger.log("Not enough columns to merge");
        return;
      }
      let newData = [];
      for (let i = 0; i < this._data.length; i++) {
        let newRow = [];
        let mergedValue = "";
        for (let j = 0; j < colRange; j++) {
          let cell = this._data[i][column + j - 1];
          if (cell != null) {
            mergedValue += cell.toString() + delimiter;
          }
        }
        if (mergedValue.length > 0 && delimiter) {
          mergedValue = mergedValue.slice(0, -delimiter.length);
        }
        newRow.push(mergedValue);
        newData.push(newRow);
      }

      this._sheet.getRange(2, column, newData.length, 1).setValues(newData);
      for(let i = column + 1; i <= column + colRange - 1; i++) {
        this._sheet.deleteColumn(i);
      }
      this._sheet.getRange(1, column).setValue(`Merged Columns`);
      Logger.log("Merged columns");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "mergeCols", [delimiter, this._currentCol, this._currentColRange]);
    }
    catch (err) {
      Logger.log(`Error merging columns: ${err}`);
      this._ui.alert("An error occurred while merging columns. Please check the logs for details.");
    }
  }

  insertIndex(start = 0){
    try{
      if(this._numRows < 1) {
        Logger.log("Not enough data to fill");
        return;
      }

      let index = start;
      let indexCol = this._data.map(row => [index++]);
      this._sheet.insertColumnAfter(this._numCols);
      this._sheet.getRange(2, this._numCols + 1, indexCol.length, 1).setValues(indexCol);
      this._sheet.getRange(1, this._numCols + 1).setValue("Index");

      Logger.log("Inserted Index Column");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "insertIndex", [start]);
    }
    catch (err) {
      Logger.log(`Error inserting index col: ${err}`);
      this._ui.alert("An error occurred while inserting index col. Please check the logs for details.");
    }
  }

  // split methods
  splitColByDelimiter(delimiter, column = this._currentCol, repeat = true){
    try{
      if(this._numRows <= 1 || this._numCols < 1){
        Logger.log("Not enough data to split");
        this._ui.alert("Not enough data to split");
        return;
      }
      let splitData = this._data.map(row => {
        let cell = row[column - 1];
        if (cell != null && cell.toString().trim() !== "") {
          return cell.toString().split(delimiter);
        }
        return [""];
      });

      let maxSplit = Math.max(...splitData.map(row => row.length));
      
      if(maxSplit < 2) {
        Logger.log("No split occurred, delimiter not found in any cell.");
        this._ui.alert("No split occurred, delimiter not found in any cell.");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "splitColByDelimiter", [delimiter, column, repeat]);
        return;
      }

      if(maxSplit > 1){
        this._sheet.insertColumnsAfter(column, maxSplit - 1);
      }
      let headerName = this._headers[column - 1]|| "column";

      for (let i = 0; i < splitData.length; i++) {
        while (splitData[i].length < maxSplit) splitData[i].push(""); 
      }

      this._sheet.getRange(2, column, splitData.length, maxSplit).setValues(splitData);
      let splitHeaders = [];
      for(let i = 1; i <= maxSplit; i++){
        splitHeaders.push(`${headerName}.${i}`);
      }

      this._sheet.getRange(1, column, 1, splitHeaders.length).setValues([splitHeaders]);
      Logger.log("Split column");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "splitColByDelimiter", [delimiter, column, repeat]);
    }
    catch(err){
      Logger.log(`Unable to split the columns ${err}`);
      this._ui.alert("An error occurred while spliting the columns. Please check the logs for details");
    }
  }

  splitColByCharacters(charNum, column = this._currentCol, repeat = true){
    try{
      if(this._numCols < 2){
        Logger.log("Not enough columns to split");
        return;
      }
      let splitData = [];
      let maxSplits = 0;

      for(let i = 0; i < this._data.length; i++){
        let cell = this._data[i][column - 1];
        if(cell != null && cell.toString().trim() !== ""){
          if(cell.toString().length < charNum){
            charNum = cell.toString().length;
          }
          let str = cell.toString();
          let splitRow = [];
          if(repeat){
            for(let j = 0; j < str.length; j += charNum){
              splitRow.push(str.substring(j, j + charNum));
            }
          }
          else{
            splitRow.push(str.substring(0, charNum));
            splitRow.push(str.substring(charNum));
          }
          splitData.push(splitRow);
          maxSplits = Math.max(maxSplits, splitRow.length);
        }
        else {
          splitData.push([""]);
        }
      }
      
      if(maxSplits > 1){
        this._sheet.insertColumnsAfter(column, maxSplits - 1);
        for (let i = 0; i < splitData.length; i++) {
          while (splitData[i].length < maxSplits) splitData[i].push("");
        }
        this._sheet.getRange(2, column, splitData.length, maxSplits).setValues(splitData);
        let headerName = this._sheet.getRange(1, column).getValue();
        let splitHeaders = [];
        for(let i = 1; i <= maxSplits; i++){
          splitHeaders.push(`${headerName}.${i}`);
        }
        this._sheet.getRange(1, column, 1, maxSplits).setValues([splitHeaders]);
      }

      Logger.log("Split columns by characters");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "splitColByCharacters", [charNum, column, repeat]);
    }
    catch(err){
      Logger.log(`Error splitting columns by characters: ${err}`);
      this._ui.alert("An error occurred while splitting columns. Please check the logs for details.");
      return;
    }
  }

  splitColTexttoNum(repeat = true){
    try{
      this.splitColByDelimiter(/(?<=[a-zA-Z])(?=\d)/);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "splitColTexttoNum", [repeat]);
    }
    catch(err){
      Logger.log(`Error splitting column text to number: ${err}`);
      this._ui.alert("An error occurred while splitting column text to number. Please check the logs for details.");
      return;
    }
  }

  splitColNumtoText(repeat = true){
    try{
      this.splitColByDelimiter(/(?<=\d)(?=[a-zA-Z])/);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "splitColNumtoText", [repeat]);
    }
    catch(err){
      Logger.log(`Error splitting column number to text: ${err}`);
      this._ui.alert("An error occurred while splitting column number to text. Please check the logs for details.");
      return;
    }
  }

  splitColUppertoLower(repeat = true){
    try{
      this.splitColByDelimiter(/(?<=[A-Z])(?=[a-z])/);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "splitColUppertoLower", [repeat]);
    }
    catch(err){
      Logger.log(`Error splitting column upper to lower: ${err}`);
      this._ui.alert("An error occurred while splitting column upper to lower. Please check the logs for details.");
      return;
    }
  }

  splitColLowertoUpper(repeat = true){
    try{
      this.splitColByDelimiter(/(?<=[a-z])(?=[A-Z])/);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "splitColLowertoUpper", [repeat]);
    }
    catch(err){
      Logger.log(`Error splitting column lower to upper: ${err}`);
      this._ui.alert("An error occurred while splitting column lower to upper. Please check the logs for details.");
      return;
    }
  }

  // formatting methods
  lowerCase(column = this._currentCol, colRange = this._currentColRange){
    try{
      if (this._numRows < 1) {
        Logger.log("Not enough data to convert to lower case");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "lowerCase", [column, colRange]);
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < colRange; j++) {
          if(this._data[i][column + j - 1] != null) {
            newRow.push(this._data[i][column + j - 1].toString().toLowerCase());
          }
          else {
            newRow.push("");
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, column, newData.length, colRange).setValues(newData);
      Logger.log("Converted data to lower case");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "lowerCase", [this._currentCol, this._currentColRange]);
      return;
    }

    catch (err){
      Logger.log(`Error converting to lower case: ${err}`);
      this._ui.alert("An error occurred while converting to lower case. Please check the logs for details.");
      return;
    }
  }

  upperCase(column = this._currentCol, colRange = this._currentColRange){
    try{
      if (this._numRows < 1) {
        Logger.log("Not enough data to convert to upper case");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "upperCase", [column, colRange]);
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < colRange; j++) {
          let cell = this._data[i][column + j - 1];
          if(cell != null) {
            newRow.push(cell.toString().toUpperCase());
          } else {
            newRow.push("");
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, column, newData.length, colRange).setValues(newData);
      Logger.log("Converted data to upper case");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "upperCase", [column, colRange]);
    }
    catch (err){
      Logger.log(`Error converting to upper case: ${err}`);
      this._ui.alert("An error occurred while converting to upper case. Please check the logs for details.");
    }
  }

  capitalize(column = this._currentCol, colRange = this._currentColRange){
    try{
      if(this._numRows < 1) {
        Logger.log("Not enough data to capitalize");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "capitalize", [column, colRange]);
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < this._currentColRange; j++) {
          let cell = this._data[i][this._currentCol + j - 1];
          if(cell != null && cell.toString().length > 0) {
            let str = cell.toString();
            str = str.split(" ").map(word => word.trim().slice(0, 1).toUpperCase() + word.trim().slice(1).toLowerCase()).join(" ");
            newRow.push(str);
          } else {
            newRow.push("");
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, this._currentCol, newData.length, this._currentColRange).setValues(newData);
      Logger.log("Capitalized the data");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "capitalize", [this._currentCol, this._currentColRange]);
    }
    catch (err){
      Logger.log(`Error capitalizing data: ${err}`);
      this._ui.alert("An error occurred while capitalizing. Please check the logs for details.");
    }
  }

  trim(column = this._currentCol, colRange = this._currentColRange){
    try{
      if(this._numRows < 1){
        Logger.log("Not enough data to trim");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "trim", [this._currentCol, this._currentColRange]);
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < colRange; j++) {
          let cell = this._data[i][column + j - 1];
          if(cell != null) {
            newRow.push(cell.toString().trim());
          } else {
            newRow.push("");
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, column, newData.length, colRange).setValues(newData);
      Logger.log("Trimmed the data");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "trim", [column, colRange]);
    }
    catch (err){
      Logger.log(`Error trimming data: ${err}`);
      this._ui.alert("An error occurred while trimming. Please check the logs for details.");
    }
  }

  addCharacters(prefix, suffix, column = this._currentCol, colRange = this._currentColRange){
    try{
      if(this._numRows < 1){
        Logger.log("Not enough data to add prefix");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < colRange; j++) {
          let cell = this._data[i][column + j - 1];
          if(cell != null) {
            newRow.push(prefix + cell.toString() + suffix);
          } else {
            newRow.push(prefix + suffix);
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, column, newData.length, colRange).setValues(newData);
      this._data = newData;
      if(prefix === ""){
        Logger.log("Added suffix");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "addCharacters", [prefix, suffix, column, colRange]);
      }
      else{
        Logger.log("Added prefix");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "addCharacters", [prefix, suffix, column, colRange]);
      }
    }
    catch (err){
      Logger.log(`Error adding characters: ${err}`);
      this._ui.alert("An error occurred while adding characters. Please check the logs for details.");
    }
  }

  //Extract methods
  length(column = this._currentCol) {
    try{
      if(this._numRows < 1){
        Logger.log("Not enough data to find length");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "length", [this._currentCol]);
        return;
      }
      let newLengthData = [];
      for(let i = 0; i < this._data.length; i++) {
        let cell = this._data[i][column - 1];
        if(cell != null) {
          newLengthData.push([cell.toString().length]);
        }
        else {
          newLengthData.push([0]);
        }
      }
      this._sheet.getRange(2, this._numCols + 1, newLengthData.length, 1).setValues(newLengthData);
      let currentColName = this._sheet.getRange(1, column).getValue();
      this._sheet.getRange(1, this._numCols + 1).setValue(`Length ${currentColName}`);
      Logger.log("Calculated length of the column");
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "length", [this._currentCol]);
    }
    catch (err){
      Logger.log(`Error calculating length: ${err}`);
      this._ui.alert("An error occurred while calculating length. Please check the logs for details.");
    }
  }

  extractText(condition, start, end = 1, column = this._currentCol) {
    try {
      if (this._numRows < 1) {
        Logger.log("Not enough data to extract text");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let cell = this._data[i][column - 1];
        if(cell != null) {
          let str = cell.toString();
          if(condition === "first") {
            if(start > str.length) {
              start = str.length;
            }
            newData.push([str.substring(0, start)]);
          } 
          else if(condition === "last") {
            if(end > str.length) {
              end = str.length;
            }
            newData.push([str.substring(str.length - end, str.length)]);
          } 
          else if(condition === "range") {
            if(end > str.length) {
              end = str.length;
            }
            newData.push([str.substring(start, str.length - end)]);
          } 
          else {
            Logger.log("Invalid condition for text extraction");
            return;
          }
        } 
        else {
          newData.push([""]);
        }
      }
      this._sheet.insertColumnAfter(this._currentCol);
      this._sheet.getRange(2, this._currentCol + 1, newData.length, 1).setValues(newData);
      this._sheet.getRange(1, this._currentCol + 1).setValue(`Extracted Text`);
      Logger.log(`Extracted ${condition} characters`);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "extractText", [condition, start, end, this._currentCol]);
    }
    catch (err) {
      Logger.log(`Error extracting text: ${err}`);
      this._ui.alert("An error occurred while extracting text. Please check the logs for details.");
    }
  }

  extractTextDelimiter(condition, delimiter, skip = 0, column = this._currentCol) {
    try{
      if(this._numRows < 1) {
        Logger.log("Not enough data to extract text by delimiter");
        FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "extractTextDelimiter", [condition, delimiter, skip, column]);
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let cell = this._data[i][column - 1];
        if(cell != null) {
          let str = cell.toString();
          if(condition === "before") {
            let parts = str.split(delimiter);
            if(skip > 0){
              if(parts.length > skip) {
                newData.push([parts.slice(0, skip).join(delimiter)]);
              } else {
                newData.push([str]);
              }
            }
            else{
              newData.push([parts[0]]);
            }
          } 
          else if(condition === "after") {
            let parts = str.split(delimiter);
            if(skip > 0){
              if(parts.length > skip) {
                newData.push([parts.slice(skip).join(delimiter)]);
              } 
              else {
                newData.push([str]);
              }
            }
            else{
              newData.push([parts[parts.length - 1]]);
            }
          } 
        } 
        else {
          newData.push([""]);
        }
      }
      this._sheet.insertColumnAfter(this._currentCol);
      this._sheet.getRange(2, this._currentCol + 1, newData.length, 1).setValues(newData);
      this._sheet.getRange(1, this._currentCol + 1).setValue(`Extracted Text`);
      Logger.log(`Extracted text by ${condition} delimiter`);
      FLOW_MANAGER.flow.addStep(CURRENT_FLOW, "extractTextDelimiter", [condition, delimiter, skip, column]);
    }
    catch (err) {
      Logger.log(`Error extracting text by delimiter: ${err}`);
      this._ui.alert("An error occurred while extracting text by delimiter. Please check the logs for details.");
    }
  }
}
