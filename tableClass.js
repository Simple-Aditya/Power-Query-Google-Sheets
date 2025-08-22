class table {
  constructor(spreadsheet) {
    this._sheet = spreadsheet.getActiveSheet();
    this._range = this._sheet.getRange(1, 1, this._sheet.getLastRow(), this._sheet.getLastColumn());
    this._executionSteps = [];
    this._numRows = this._sheet.getLastRow() || 1;
    this._numCols = this._sheet.getLastColumn();
    this._table = this._range.getValues();
    this._data = this._table.slice(1);
    this._headers = this._table[0];
    this._ui = spreadsheet.getUi();
    this._currentCell = this._sheet.getActiveCell();
    this._currentCol = this._currentCell.getColumn();
    this._currentColRange = this._sheet.getActiveRange().getNumColumns();
  }

  // table methods
  reverseRows(){
    if (this._numRows <= 1) {
      Logger.log("Not enough data to reverse");
      return;
    }
    this._data.reverse();
    this._sheet.getRange(2, 1, this._numRows - 1, this._sheet.getLastColumn()).clearContent();
    if (this._data.length > 0) {
      this._sheet.getRange(2, 1, this._data.length, this._data[0].length).setValues(this._data);
    }
    Logger.log("Reversed the rows");
    this._executionSteps.push("Reversed the rows");
  }

  // This method sets the first row as headers and freezes it
  makeFirstRowHeaders() {
    this._sheet.setFrozenRows(1);
    this._sheet.getRange(1, 1, 1, this._sheet.getLastColumn())
    .setFontWeight("bold")
    .setBackground("#f2f2f2");
  }

  // This method inserts a new sheet with a Row Count name and gives the total number of rows
  countRows(){
    const newSheet = insertNewSheet("Row Count");
    newSheet.getRange(1, 1, 1, 2).setValues([[`Total Rows in ${this._sheet.getName()}`, this._numRows]]);
    Logger.log(`Total Rows: ${this._numRows}`);
    this._executionSteps.push(`Total Rows: ${this._numRows}`);
  }

  replaceValue(oldValue, newValue = "") {
  try {
    if (this._numRows < 1) {
      Logger.log("No data to replace values");
      return;
    }
    
    // Validate oldValue - don't allow empty string searches
    if (!oldValue || oldValue.toString().trim() === "") {
      Logger.log("Cannot replace empty or whitespace-only values");
      this._ui.alert("Cannot replace empty or whitespace-only values");
      return;
    }
    
    let newData = [];
    let replacementCount = 0;
    
    for (let i = 0; i < this._data.length; i++) {
      let newRow = [];
      for (let j = 0; j < this._currentColRange; j++) {
        let cell = this._data[i][this._currentCol + j - 1];
        
        if (cell != null) {
          let cellStr = cell.toString();
          if (cellStr.includes(oldValue.toString())) {
            // Use replaceAll to replace ALL occurrences, not just the first
            let newCellValue = cellStr.replaceAll(oldValue.toString(), newValue.toString());
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
    
    if (replacementCount > 0) {
      this._sheet.getRange(2, this._currentCol, newData.length, this._currentColRange).setValues(newData);
      this._data = newData;
      Logger.log(`Replaced "${oldValue}" with "${newValue}" in ${replacementCount} cells`);
      this._executionSteps.push(`Replaced "${oldValue}" with "${newValue}" in ${replacementCount} cells`);
    } else {
      Logger.log(`No instances of "${oldValue}" found to replace`);
      this._ui.alert(`No instances of "${oldValue}" found in the selected range`);
    }
    
  } catch (err) {
    Logger.log(`Error replacing values: ${err}`);
    this._ui.alert("An error occurred while replacing values. Please check the logs for details.");
  }
}
  
  transpose(){
    try{
      if(this._numRows < 1){
        Logger.log("Insufficient data to transpose");
        return;
      }
      let newData = [];
      for(let j = 0; j < this._numCols; i++){
        let newRow = [];
        for(let i = 0; i < this._numRows; i++){
          newRow.push(this._data[i][j] ? this._data[i][j] : "");
        }
        newData.push(newRow);
      }
      this._sheet.clearContents();
      this._sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      this.makeFirstRowHeaders();
      Logger.log("Transposed the data");
      this._executionSteps.push("Transposed Data");
    }
    catch(err){
      Logger.log(`Unable to transpose the data because ${err}`);
      this._ui.alert("An error occurred while transposing the data. Please check the logs for details")
    }
  }

  // row methods
  // This method returns the first row as headers
  // params: rows: number of rows to keep or remove, condition: "top" or "last"
  keepRows(rows, condition = "top") {
    if (this._numRows <= 1) {
      Logger.log(`Kept ${condition} ${rows} rows`);
      this._executionSteps.push(`Kept ${condition} ${rows} rows`);
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
    this._executionSteps.push(`Kept ${condition} ${rows} rows`);
  }

  // This method removes rows from the top or bottom based on the condition
  // params: rows: number of rows to remove, condition: "top" or "last"
  removeRows(rows, condition = "top") {
    if (rows <= 0) {
      this._ui.alert("Please enter a valid input.");
      return;
    }
    if (this._numRows < 1) {
      Logger.log(`Unable to remove ${condition} ${rows} rows due to insufficiency.`);
      this._executionSteps.push(`Unable to remove ${condition} ${rows} rows due to insufficiency.`);
      return;
    }
    let newData = [];
    if (condition === "last") {
      newData = this._data.slice(0, this._numRows - rows);
    } else {
      newData = this._data.slice(rows, this._numRows);
    }
    this._sheet.getRange(2, 1, this._numRows - 1, this._sheet.getLastColumn()).clearContent();
    if (newData.length > 0) {
      this._sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
    }
    this._data = newData;
    Logger.log(`Removed ${condition} ${rows} rows.`);
    this._executionSteps.push(`Removed ${condition} ${rows} rows.`);
  }

  // This method removes blank rows from the data
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
    this._executionSteps.push("Removed empty rows");
  }

  // column methods
  fillDown(){
    if(this._numRows <= 1) {
      Logger.log("Not enough data to fill the data");
      return;
    }
    let newData = [];
    let lastPushed = "";
    for(let i = 0; i < this._data.length; i++){
      let cell = this._data[i][this._currentCol - 1];
      if (cell != null && cell.toString().trim() !== ""){
        newData.push([cell]);
        lastPushed = cell;
      }
      else {
        newData.push([lastPushed]);
      }
    }
    this._sheet.getRange(2, this._currentCol, newData.length, 1).setValues(newData);
    Logger.log("Performed down fill on the data");
    this._executionSteps.push("Data filled down");
  }
  
  fillUp(){
    if(this._numRows <= 1) {
      Logger.log("Not enough data to fill");
      return;
    }
    let newData = [];
    let lastPushed = "";
    for(let i = this._data.length - 1; i >= 0 ; i--){
      let cell = this._data[i][this._currentCol - 1];
      if (cell != null && cell.toString().trim() !== ""){
        newData.push([cell]);
        lastPushed = cell;
      }
      else {
        newData.push([lastPushed]);
      }
    }
    this._sheet.getRange(2, this._currentCol, newData.length, 1).setValues(newData);
    Logger.log("Performed up fill on the data");
    this._executionSteps.push("Data filled Up");
  }

  mergeCols(delimiter = "") {
    try{
      if (this._numCols < 2) {
        Logger.log("Not enough columns to merge");
        return;
      }
      let newData = [];
      for (let i = 0; i < this._data.length; i++) {
        let newRow = [];
        let mergedValue = "";
        for (let j = 0; j < this._currentColRange; j++) {
          let cell = this._data[i][this._currentCol + j - 1];
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
      this._sheet.getRange(2, this._currentCol, newData.length, 1).setValues(newData);
      for(let i = this._currentCol + 1; i <= this._currentCol + this._currentColRange - 1; i++) {
        this._sheet.deleteColumn(i);
      }
      this._sheet.getRange(1, this._currentCol).setValue(`Merged Columns`);
      Logger.log("Merged columns");
      this._executionSteps.push("Merged columns");
    }
    catch (err) {
      Logger.log(`Error merging columns: ${err}`);
      this._ui.alert("An error occurred while merging columns. Please check the logs for details.");
    }
  }

  // can do only when take column name input
  pivot(){
    if(this._numRows <= 1) {
      Logger.log("Not enough data to fill");
      return;
    }
    let pivotHeaders = new Set();
    for(let i = 0; i < this._numRows; i++){
      pivotHeaders.add(this._data[i][this._currentCol - 1]);
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
      this._executionSteps.push("Inserted Index Col");
    }
    catch (err) {
      Logger.log(`Error inserting index col: ${err}`);
      this._ui.alert("An error occurred while inserting index col. Please check the logs for details.");
    }
  }

  // split methods
  splitColByDelimiter(delimiter){
    try{
      if(this._numRows <= 1 || this._numCols < 1){
        Logger.log("Not enough data to split");
        this._ui.alert("Not enough data to split");
        return;
      }
      let splitData = this._data.map(row => {
        let cell = row[this._currentCol - 1];
        if (cell != null && cell.toString().trim() !== "") {
          return cell.toString().split(delimiter);
        }
        return [""];
      });

      let maxSplit = Math.max(...splitData.map(row => row.length));
      
      if(maxSplit < 2) {
        Logger.log("No split occurred, delimiter not found in any cell.");
        this._ui.alert("No split occurred, delimiter not found in any cell.");
        return;
      }

      if(maxSplit > 1){
        this._sheet.insertColumnsAfter(this._currentCol, maxSplit - 1);
      }
      let headerName = this._headers[this._currentCol - 1]|| "column";

      for (let i = 0; i < splitData.length; i++) {
        while (splitData[i].length < maxSplit) splitData[i].push(""); 
      }

      this._sheet.getRange(2, this._currentCol, splitData.length, maxSplit).setValues(splitData);
      let splitHeaders = [];
      for(let i = 1; i <= maxSplit; i++){
        splitHeaders.push(`${headerName}.${i}`);
      }
      this._sheet.getRange(1, this._currentCol, 1, splitHeaders.length).setValues([splitHeaders]);
      Logger.log("Split column");
      this._executionSteps.push("Split column");
    }
    catch(err){
      Logger.log(`Unable to split the columns ${err}`);
      this._ui.alert("An error occurred while spliting the columns. Please check the logs for details");
    }
  }

  splitColByCharacters(charNum, repeat = true){
    try{
      if(this._numCols < 2){
        Logger.log("Not enough columns to split");
        return;
      }
      let splitData = [];
      let maxSplits = 0;
      for(let i = 0; i < this._data.length; i++){
        let cell = this._data[i][this._currentCol - 1];
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
        this._sheet.insertColumnsAfter(this._currentCol, maxSplits - 1);
      }
      for (let i = 0; i < splitData.length; i++) {
        while (splitData[i].length < maxSplits) splitData[i].push("");
      }
      this._sheet.getRange(2, this._currentCol, splitData.length, maxSplits).setValues(splitData);
      let headerName = this._sheet.getRange(1, this._currentCol).getValue();
      let splitHeaders = [];
      for(let i = 1; i <= maxSplits; i++){
        splitHeaders.push(`${headerName}.${i}`);
      }
      this._sheet.getRange(1, this._currentCol, 1, maxSplits).setValues([splitHeaders]);
      Logger.log("Split columns by characters");
      this._executionSteps.push("Split columns by characters");
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
    }
    catch(err){
      Logger.log(`Error splitting column lower to upper: ${err}`);
      this._ui.alert("An error occurred while splitting column lower to upper. Please check the logs for details.");
      return;
    }
  }

  // formatting methods
  lowerCase(){
    try{
      if (this._numRows < 1) {
        Logger.log("Not enough data to convert to lower case");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < this._currentColRange; j++) {
          if(this._data[i][this._currentCol + j - 1] != null) {
            newRow.push(this._data[i][this._currentCol + j - 1].toString().toLowerCase());
          }
          else {
            newRow.push("");
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, this._currentCol, newData.length, this._currentColRange).setValues(newData);
      Logger.log("Converted data to lower case");
      this._executionSteps.push("Converted data to lower case");
      return;
    }
    catch (err){
      Logger.log(`Error converting to lower case: ${err}`);
      this._ui.alert("An error occurred while converting to lower case. Please check the logs for details.");
      return;
    }
  }

  upperCase(){
    try{
      if (this._numRows < 1) {
        Logger.log("Not enough data to convert to upper case");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < this._currentColRange; j++) {
          let cell = this._data[i][this._currentCol + j - 1];
          if(cell != null) {
            newRow.push(cell.toString().toUpperCase());
          } else {
            newRow.push("");
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, this._currentCol, newData.length, this._currentColRange).setValues(newData);
      Logger.log("Converted data to upper case");
      this._executionSteps.push("Converted data to upper case");
    }
    catch (err){
      Logger.log(`Error converting to upper case: ${err}`);
      this._ui.alert("An error occurred while converting to upper case. Please check the logs for details.");
    }
  }

  capitalize(){
    try{
      if(this._numRows < 1) {
        Logger.log("Not enough data to capitalize");
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
      this._executionSteps.push("Capitalized the data");
    }
    catch (err){
      Logger.log(`Error capitalizing data: ${err}`);
      this._ui.alert("An error occurred while capitalizing. Please check the logs for details.");
    }
  }

  trim(){
    try{
      if(this._numRows < 1){
        Logger.log("Not enough data to trim");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < this._currentColRange; j++) {
          let cell = this._data[i][this._currentCol + j - 1];
          if(cell != null) {
            newRow.push(cell.toString().trim());
          } else {
            newRow.push("");
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, this._currentCol, newData.length, this._currentColRange).setValues(newData);
      Logger.log("Trimmed the data");
      this._executionSteps.push("Trimmed the data");
    }
    catch (err){
      Logger.log(`Error trimming data: ${err}`);
      this._ui.alert("An error occurred while trimming. Please check the logs for details.");
    }
  }

  addCharacters(p, s){
    try{
      if(this._numRows < 1){
        Logger.log("Not enough data to add prefix");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let newRow = [];
        for(let j = 0; j < this._currentColRange; j++) {
          let cell = this._data[i][this._currentCol + j - 1];
          if(cell != null) {
            newRow.push(p + cell.toString() + s);
          } else {
            newRow.push(p + s);
          }
        }
        newData.push(newRow);
      }
      this._sheet.getRange(2, this._currentCol, newData.length, this._currentColRange).setValues(newData);
      this._data = newData;
      if(p === ""){
        Logger.log("Added suffix");
        this._executionSteps.push("Added suffix");
      }
      else{
        Logger.log("Added prefix");
        this._executionSteps.push("Added Prefix");
      }
    }
    catch (err){
      Logger.log(`Error adding characters: ${err}`);
      this._ui.alert("An error occurred while adding characters. Please check the logs for details.");
    }
  }

  //Extract methods
  length(){
    try{
      if(this._numRows < 1){
        Logger.log("Not enough data to find length");
        return;
      }
      let newLengthData = [];
      for(let i = 0; i < this._data.length; i++) {
        let cell = this._data[i][this._currentCol - 1];
        if(cell != null) {
          newLengthData.push([cell.toString().length]);
        }
        else {
          newLengthData.push([0]);
        }
      }
      this._sheet.getRange(2, this._numCols + 1, newLengthData.length, 1).setValues(newLengthData);
      let currentColName = this._sheet.getRange(1, this._currentCol).getValue();
      this._sheet.getRange(1, this._numCols + 1).setValue(`Length ${currentColName}`);
      Logger.log("Calculated length of the column");
      this._executionSteps.push("Calculated length");
    }
    catch (err){
      Logger.log(`Error calculating length: ${err}`);
      this._ui.alert("An error occurred while calculating length. Please check the logs for details.");
    }
  }

  extractText(condition, start, end = 1) {
    try{

      if(this._numRows < 1) {
        Logger.log("Not enough data to extract text");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let cell = this._data[i][this._currentCol - 1];
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
      this._executionSteps.push(`Extracted ${condition} characters`)
    }
    catch (err) {
      Logger.log(`Error extracting text: ${err}`);
      this._ui.alert("An error occurred while extracting text. Please check the logs for details.");
    }
  }

  extractTextDelimiter(condition, delimiter, skip = 0){
    try{
      if(this._numRows < 1) {
        Logger.log("Not enough data to extract text by delimiter");
        return;
      }
      let newData = [];
      for(let i = 0; i < this._data.length; i++) {
        let cell = this._data[i][this._currentCol - 1];
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
      this._executionSteps.push(`Extracted text by ${condition} delimiter`);
    }
    catch (err) {
      Logger.log(`Error extracting text by delimiter: ${err}`);
      this._ui.alert("An error occurred while extracting text by delimiter. Please check the logs for details.");
    }
  }

}
