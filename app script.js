var spreadsheetID = "0AlaVan9pZtAzdEF5Wm9HQzFiTlpNQVF4a3hmWDJxSGc";

var logSheetName = "Log";
var customizeStatusColorSheetName = "Status Color";
var statusChangeColumnName = "Status";

var addTodayWhenEdit = [ // [whenEditCell, addTodayCell]
  ["Report By", "Report Date"],
  ["Completed By", "Completion Date"]
];

var backgroundColorPriority = [
  // status: Top priority <--------------------> Low priority
  ["tailor make", "hardcode", "holding", "follow up", "misreporting", "cancelled", "pending", "release", "done"],
  // stauts background color
  ["#d9d2e9", "#f4cccc", "#f4cccc", "#c9daf8", "#efefef", "#efefef", "#fff2cc", "#d9ead3", "#d9ead3"]
];

// Prepare a New Line
var isPrepareNewLines = false; // turn this on(ture) if you kown what you are doing
// ====================================================================
var autoGenKeyColumn = ["#.", ""]; // [columnHeaderName, autoGenFormat], autoGenFormat just a idea not done
var prepareNewLineWhenEdit = "Description";
var selectionList = [
  //[statusChangeColumnName, backgroundColorPriority[0]], // uncomment this if your status allow a single status only
  ["Issue Type", ["Bug", "Improvement", "Tuning", "What", "Check"]],
  ["Priority", ["P1", "P2", "P3", "P4"]]
];
var prepareHowManyRows = 5;

//not implemented
var requireValidation = [ // instance, validationMethods, [argumentType, argumentValue...]
  [addTodayWhenEdit[1][1], "requireDateOnOrAfter", ["date", "eval('sheet.getRange(prepareThisRow, findColumnIndexByHeader(sheet, addTodayWhenEdit[0][1]]])).getValue()')"]] // "Completion Date" must OnOrAfter "Report Date"
];

Logger.clear();

function changeBgColorByStatus() {
  Logger.log("changeBgColorByStatus() function execute");
  var ss = SpreadsheetApp.openById(spreadsheetID);
  var sheet = ss.getSheetByName(logSheetName);
  var customColorSheet = ss.getSheetByName(customizeStatusColorSheetName);
  
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var range = sheet.getDataRange();
  var numRowsHaveData = range.getNumRows();
  
  Logger.log("Current Spreadsheet max rows: "+numRowsHaveData);
  
  // white,undefined  	= nothing
  // Hardcode 			= #f4cccc
  // Holding 			= #f4cccc
  // Misreporting 		= #efefef
  // Cancelled 			= #efefef
  // Pending 			= #fff2cc
  // Release 			= #d9ead3
  // Done 				= #d9ead3
  
  for(var rowIndex = frozenRows+1; rowIndex<numRowsHaveData; rowIndex++){
    Logger.log("Try to change color at row:"+rowIndex);
    changeARowBgColorByStatus(customColorSheet, sheet, rowIndex);
  }
}

// http://stackoverflow.com/questions/3703676/google-spreadsheet-script-to-change-row-color-when-a-cell-changes-text
/*
function doGet(e)
{
  var app = UiApp.createApplication();
  var site = SitesApp.getActiveSite();
  var label  = app.createLabel("Hello Wrold");
  app.add(label);

  //changeBgColorByStatus()
  return app;
}

function onOpen() {
  changeBgColorByStatus();
}
*/

function changeARowBgColorByStatus(statusColorSheet, sheet, rowIndex){
  var row, currentBackgroudColor = "";
  var currentRowStatusCell, currentRowStatusCellValue;
  var statusColumnIndex;
  var colors = new Array();
  var isStatusCellWhiteBg=false, isValueEmpty=false;

  // Ignore the frozen column when fill background
  var frozenCols = sheet.getFrozenColumns();
  // How many column need to fill
  var maxColumns = sheet.getMaxColumns();
  
  // if status column not found, exit function
  statusColumnIndex = findColumnIndexByHeader(sheet, statusChangeColumnName);
  if(statusColumnIndex==-1)
    return;
  
  Logger.log("Ststus column found at:"+statusColumnIndex);
  
  currentRowStatusCell = sheet.getRange(rowIndex, statusColumnIndex);
  currentRowStatusCellValue = currentRowStatusCell.getValue().toLowerCase();
  currentBackgroudColor = currentRowStatusCell.getBackground().toLowerCase();
  isStatusCellWhiteBg = currentBackgroudColor=="white" || currentBackgroudColor=="#ffffff";
  isValueEmpty = currentRowStatusCellValue=="";
  
  if(statusColorSheet!=null) // get customize BGColor, if customize BGColor sheet exist
    backgroundColorPriority = customStatusColor(statusColorSheet);

  Logger.log("Value:"+currentRowStatusCellValue+" ,isValueEmpty:"+isValueEmpty);
  
  var range = sheet.getRange(rowIndex, frozenCols+1, 1, maxColumns-frozenCols);
  
  // if status cell empty, clear row background color
    if(isValueEmpty){
      var isBackgroundColor = backgroundColorPriority[1].indexOf(currentBackgroudColor);
	  fillRangeBackground(sheet, range, "white");
      return;
    }
  
  // cell not empty, check is status specified
      var isStautsFound = backgroundColorPriority[0].indexOf(currentRowStatusCellValue);
      var isFillColor = false;
      
      if(isStautsFound >= 0){
          isFillColor = true;
      }else{
        for(var m=0; m<backgroundColorPriority[0].length; m++){
          isStautsFound = currentRowStatusCellValue.indexOf(backgroundColorPriority[0][m]);
          Logger.log("a part of status found at "+m+", which is "+backgroundColorPriority[0][m]);
          if(isStautsFound >=0){
            isFillColor = true;
            isStautsFound = m;
            break;
          }
        }
      }
	  
	  // start to fill color
      if(isFillColor){
        if(currentBackgroudColor==backgroundColorPriority[1][isStautsFound]){
		  // if current status cell background same as the corresponding ststus color, need do not to change
          return;
        }
          var bgColorchangeToThis = backgroundColorPriority[1][isStautsFound];
          var rgbColors = new RGBColor(bgColorchangeToThis);
        if(rgbColors.ok){
          bgColorchangeToThis = rgbColors.toHex();
        }else{
		  //if the specify color is invalid color
		  bgColorchangeToThis = "white"
		}
        fillRangeBackground(sheet, range, bgColorchangeToThis);
      }else{
	  // clear background if status cannot specify
	  fillRangeBackground(sheet, range, "white");
      }
}

function fillRangeBackground(sheet, range, color){
  var rangeWidth, rangeHeight;
  var colorArray = new Array();
  rangeWidth = range.getNumColumns();
  rangeHeight = range.getNumRows();
  
  // prepear string array for setBackgrounds
  for(var row=0; row<rangeHeight; row++){
    colorArray[row] = new Array();
    for(var col=0; col<rangeWidth; col++){
	  colorArray[row][col] = color;
	}
  }
  Logger.log(range.setBackgrounds(colorArray).getA1Notation()+" background color change to "+color);
}

function printTodayAt(sheet, r){
  Logger.log("printTodayAt() function execute");
  var targetValue = r.getValue()
  var targetCell = r;
  if(targetValue==""){
    var now = new Date();
    // the time zone not same as where I'm local, so toLocaleDateString() before
    //now = now.toLocaleDateString();
    //Logger.log("Insert:"+now.getFullYear()+"-"+(now.getMonth()+1)+"-"+now.getDate()+" into "+targetCell.setNumberFormat("yyyy-MM-dd").setValue(now).getA1Notation());
    Logger.log("Insert:"+now+" into "+targetCell.setNumberFormat("yyyy-MM-dd").setValue(now).getA1Notation());
  }
}

function onEdit(event){
  Logger.log("onEdit() function execute");  
  Logger.log("Locale:"+event.source.getSpreadsheetLocale());
  Logger.log("TimeZone:"+event.source.getSpreadsheetTimeZone());
  
  // get sheet
  //var ss = SpreadsheetApp.openById("spreadsheetID");
  //var sheet = ss.getSheetByName(logSheetName);
  var sheet = event.source.getSheetByName(logSheetName);
  //var sheet = event.source.getActiveSheet();
  var customColorSheet = event.source.getSheetByName(customizeStatusColorSheetName);
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();
  
  // get active range
  var r = event.range;
  
  // get the top left Coordinate of range
  var rowIndex = r.getRowIndex();
  var columnIndex = r.getColumn();
  //var cellValue = r.getValue();
  
  Logger.log("Edit Row:"+rowIndex+" Col: "+columnIndex+" Value:"+r.getValue());
  Logger.log("First:"+frozenRows+" row(s), First:"+frozenCols+" col(s) are Igrone(Frozen area)");
  
  if(frozenRows<=0){
    Logger.log("No row(s) frozen, assume the Header at the first row.");
    frozenRows = 1;
  }else if(rowIndex<frozenRows || columnIndex<frozenCols){
    Logger.log("Edit in frozen area, igrone the onEdit Event & Trigger");
    return;
  }
  
  // is editing status column?
  if(columnIndex==findColumnIndexByHeader(sheet, statusChangeColumnName)){
    changeARowBgColorByStatus(customColorSheet, sheet, rowIndex);
  }
  
  // check all the first element of the addTodayWhenEdit sub array
  for(var j=0; j<addTodayWhenEdit.length; j++){
    // is trigger insert today date?
    if(columnIndex==findColumnIndexByHeader(sheet, addTodayWhenEdit[j][0])){
      if(!r.isBlank()){ // if empty after edit, don't printToday. e.g after press Del
        printTodayAt(sheet, sheet.getRange(rowIndex, findColumnIndexByHeader(sheet, addTodayWhenEdit[j][1])));
      }
    }
  }
  
  if(columnIndex == findColumnIndexByHeader(sheet, prepareNewLineWhenEdit)){
    var autoGenColIndex = findColumnIndexByHeader(sheet, autoGenKeyColumn[0]);
    
    Logger.log("isOnEditCellEmpty:"+ r.isBlank() +" isOnEditAutoGenCell:"+ autoGenColIndex == columnIndex);
    if(!r.isBlank() && autoGenColIndex != columnIndex){
      // prepareNewLines
      prepareNewLines(sheet, rowIndex);
    }
  }
}

function customStatusColor(sheet){
  Logger.log("customStatusColor() function execute");
  if(sheet==null)
    return backgroundColorPriority;
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();
  Logger.log("frozenRows:"+frozenRows+" frozenCols:"+frozenCols);
  
  // This represents ALL the data
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  Logger.log("Data range Total row(s):"+numRows+" Total col(s)"+numCols);
  
  var isValue, isPriority;
  var customColorPriority = new Array();
  customColorPriority[0] = new Array();
  customColorPriority[1] = new Array();
  
  var tempStatus="";
  var tempBackgroudColor="";
  var isTempColorEmpty = false;
  var tempColor="";
  var isColorVaildate = false;
  var isTempProioritynNumeric = false;
  var tempProiority = 0;
  var theProiority=0;
  
  range = sheet.getRange(frozenRows+1, frozenCols+1, numRows, numCols);
  values = range.getValues();
  
  for(var readThisRows = frozenRows+1; readThisRows<=numRows; readThisRows++){
    var firstCell = sheet.getRange(readThisRows, 1);
    var firstCellValue = firstCell.getValue();
    var isStatusEmpty = firstCell.isBlank();
    
    Logger.log("firstCellValue:"+firstCellValue+" isStatusEmpty:"+isStatusEmpty);
    if(isStatusEmpty){
      Logger.log("No status specify, skill to next row.");
      continue;
    }
    
    // each column validation
    var thisRowRange = sheet.getRange(readThisRows, frozenCols+1, 1, numCols);
    var thisRowValue = thisRowRange.getValues();
    Logger.log(thisRowValue[0]);
    tempStatus = thisRowValue[0][0].toLowerCase();
    tempColor = thisRowValue[0][1];
    tempProiority = thisRowValue[0][2];
    
    isTempColorEmpty = sheet.getRange(readThisRows, frozenCols+2, 1).isBlank();
    isTempProioritynNumeric = !isNaN(parseInt(tempProiority));
    Logger.log(" isTempColorEmpty:"+isTempColorEmpty+" isTempProioritynNumeric:"+isTempProioritynNumeric);
    isColorVaildate = new RGBColor(tempColor);
    
    if(isColorVaildate.ok){
      tempBackgroudColor = tempColor;
    }else{
      tempBackgroudColor = sheet.getRange(readThisRows, frozenCols+2, 1).getBackground();
      Logger.log("No color specify, use background color:"+tempBackgroudColor);
    }
    
    if(isTempProioritynNumeric){
      // if Proiority is numeric
      customColorPriority[0][tempProiority] = tempStatus;
      customColorPriority[1][tempProiority] = tempBackgroudColor;
    }else{
      // if Proiority is non-numeric, push to the end of array
      customColorPriority[0].push(tempStatus);
      customColorPriority[1].push(tempBackgroudColor);
    }
  }
  
  customColorPriority[0].clean(undefined);
  customColorPriority[1].clean(undefined);
  for(var i=0; i<customColorPriority[0].length; i++){
    Logger.log("status:"+i+" "+customColorPriority[0][i]+" color:"+i+" "+customColorPriority[1][i]);
  }
  return customColorPriority;
}

// http://stackoverflow.com/questions/281264/remove-empty-elements-from-an-array-in-javascript
Array.prototype.clean = function(deleteValue) {
  for (var i = 0; i < this.length; i++) {
    if (this[i] == deleteValue) {         
      this.splice(i, 1);
      i--;
    }
  }
  return this;
};

function prepareANewLine(sheet, prepareThisRow){
  Logger.log("prepareANewLine() function execute");
  var isStop;
  autoGenNum(sheet, prepareThisRow, findColumnIndexByHeader(sheet, autoGenKeyColumn[0]));
  copyRuleToNextRow(sheet, prepareThisRow-1);
	/*
  if(prepareThisRow>sheet.getFrozenRows()+1){
	copyRuleToNextRow(sheet, prepareThisRow-1);
  }else{
    for(var i=0; i<selectionList.length; i++){
	  selectionListInsertion(sheet, findColumnIndexByHeader(sheet, selectionList[i][0]), selectionList[i][1]);
    }
  }
  */
}

String.prototype.capitalize = function() {
    return this.replace(/(?:^|\s)\S/g, function(a) { return a.toUpperCase(); });
};

function prepareNewLines(sheet, prepareAfterThisRow){
  Logger.log("prepareNewLines() function execute, check After Row:"+prepareAfterThisRow+" prepareHowManyRows:"+prepareHowManyRows);
  var theNumberOfGeneratedRow = 0;
  var prepareRow = true;
  var halfOfPrepareHowManyRows = Math.round(prepareHowManyRows/2);
  // checking row in top-down flow
  for(var prepareThisRow = prepareAfterThisRow+1; prepareThisRow<=prepareAfterThisRow+prepareHowManyRows; prepareThisRow++){
    Logger.log("Checking Row no.:"+prepareThisRow);
    var prepareRowRange = sheet.getRange(prepareThisRow, findColumnIndexByHeader(sheet, autoGenKeyColumn[0])); //, 1, 1);
    var prepareRowValue = prepareRowRange.getValue();
    //prepareRowValue = prepareRowValue[0][0];
    var isPrepareRowEmpty = prepareRowRange.isBlank();
    var isKeyColumnNumeric = !isNaN(prepareRowValue) && !isPrepareRowEmpty; // is prepareThisRow key column value numeric
    Logger.log("Key Value:"+prepareRowValue+" isPrepareRowEmpty:"+isPrepareRowEmpty+" isKeyColumnNumeric:"+isKeyColumnNumeric);
	
	// if prepare Row Key Column is numberic, is the sequence number valid?
    if(isKeyColumnNumeric){
      var prepareRowUpperRowRange = sheet.getRange((prepareThisRow-1), findColumnIndexByHeader(sheet, autoGenKeyColumn[0])); //, 1, 1);
      var prepareRowUpperRowValue = prepareRowUpperRowRange.getValue();
      prepareRowValue = Number(prepareRowValue);
      // is the prepare row sequence number vaild, skip to next row if valid
      if(prepareRowUpperRowValue+1==prepareRowValue){ // 
        theNumberOfGeneratedRow+=1;
        Logger.log("The row:"+prepareThisRow+" already prepared before");
        if(theNumberOfGeneratedRow>=halfOfPrepareHowManyRows){
          Logger.log("The prepared rows >= the half of prepareHowManyRows, strike to generate auto numbers");
		  return; // if the sequence number valid (int)prepareHowManyRows/2 times, that means a half of prepareHowManyRows are prepared
        }
        continue; // skip to gen this row
      }
    }
  }
  
  if(theNumberOfGeneratedRow>0){
	prepareAfterThisRow += theNumberOfGeneratedRow;
  }
  
  // checking row in bottom-up
  var theNumberOfGeneratedRow = 0;
  if(prepareAfterThisRow-sheet.getFrozenRows()>halfOfPrepareHowManyRows)
  for(var preparedRow = prepareAfterThisRow-halfOfPrepareHowManyRows+1; preparedRow<prepareAfterThisRow; preparedRow++){
	if(sheet.getRange(preparedRow, findColumnIndexByHeader(sheet, prepareNewLineWhenEdit)).isBlank()){
		theNumberOfGeneratedRow +=1;
		if(theNumberOfGeneratedRow>=halfOfPrepareHowManyRows){
          Logger.log("The blank rows >= the half of prepareHowManyRows, strike to generate auto numbers");
		  return;
		}
	}
  }
  for(var prepareThisRow = prepareAfterThisRow+1; prepareThisRow<=prepareAfterThisRow+prepareHowManyRows; prepareThisRow++){
    prepareANewLine(sheet, prepareThisRow);
  }
}

function selectionListInsertion(sheet, prepareThisCell, selectionList){
  Logger.log("selectionListInsertion() function execute");
  var rules = prepareThisCell.getDataValidations();
  for (var i = 0; i < rules.length; i++) {
     var rule = rules[i];

     if (rule != null) {
       var rule = SpreadsheetApp.newDataValidation().requireValueInList(selectionList, true).build();
       prepareThisCell.setDataValidation(rule);
       return false;
     }
  }
  return true;
}

function autoGenNum(sheet, autoGenNumThisRow, autoGenColumnIndex){
  Logger.log("autoGenNum() function execute");
  // return bool(false/true) to continues/stop prepareNewLines;
  var theNextNumber = -1;
  
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();
  var maxRows = parseInt(sheet.getMaxRows())
  var maxColumns = parseInt(sheet.getMaxColumns());
  
  // This represents ALL the data
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  var numOfRows = values.length-frozenRows;
  range = sheet.getRange(frozenRows+1, autoGenColumnIndex, numOfRows, 1);
  values = range.getValues();
  //var values = SpreadsheetApp.getActiveSheet().getRange(2, 3, 6, 4).getValues();
  Logger.log("Total "+numOfRows+" row(s) read, start from row "+frozenRows);
  
  // Min and max in multidimensional array
  var xVals = values.map(function(obj) { return obj; });
  var max = Math.max.apply(null, xVals);
  var min = Math.min.apply(null, xVals);
  // var max = Math.max.apply(Math, a.map(function(obj){return obj;}));
  // if the array like this, change obj to obj.x
  // var a = new Array();
  // a[0] = {x: 10,y: 10};
  // a[1] = {x: 20,y: 50};
  // http://stackoverflow.com/questions/15042887/min-and-max-in-multidimensional-array
  // http://stackoverflow.com/questions/4020796/finding-the-max-value-of-an-attribute-in-an-array-of-objects
  
  var currentRowAutoGenCell = sheet.getRange(autoGenNumThisRow-1, autoGenColumnIndex);
  
  Logger.log("max:"+max+", min:"+min);
  Logger.log("Current row auto gen num cell:"+currentRowAutoGenCell.getValue() );
  
  // if i copy row to sheet and tpying a new row, new row autoGenCell are forever empty
  // checking handling in description checking in prepareNewLines
  /* is activate row autoGenNumCell is empty
  if(min==0 || currentRowAutoGenCell.getValue()==""){
    Logger.log("The insertion is invalid. Because some of the auto gen cell is empty or 0");
    Logger.log("Please insert into the row which auto gen num cell value is vaild");
    Logger.log("Stop auto generate number");
    return true;
  }*/
  
  var autoGenNumCell = sheet.getRange(autoGenNumThisRow, autoGenColumnIndex);
  var newAutoGenValue = max+1;
  Logger.log("Row:"+(autoGenNumThisRow)+" Col:"+autoGenColumnIndex+" setValueue:"+newAutoGenValue);
  if(autoGenNumCell.isBlank()){
    autoGenNumCell.setValue(newAutoGenValue);
  }else{
    Logger.log("The cell is not empty, don't replace the current value "+autoGenNumCell.getValue());
  }
}

function findColumnIndexByHeader(sheet, header){
  Logger.log("findColumnIndexByHeader() function execute");
  var frozenHeaderRange = new Array();
  var frozenHeaderValues = new Array();
  var headerFoundAtColumn = -1;
  var frozenRows = sheet.getFrozenRows();
  var maxColumns = parseInt(sheet.getMaxColumns());
  
  // get the Header(FrozenRows) cell value
  for(var rowPointer = 1; rowPointer<=frozenRows;rowPointer++){
    frozenRowRange = sheet.getRange(rowPointer, 1, 1, maxColumns);
    frozenRowValues = frozenRowRange.getValues();
    headerFoundAtColumn = frozenRowValues[0].indexOf(header);
    if(headerFoundAtColumn>=0){
      Logger.log("header:"+header+" FoundAtColumn:"+headerFoundAtColumn+" return:"+(headerFoundAtColumn+1));
      return headerFoundAtColumn+1;
      break;
    }
  }
  return headerFoundAtColumn;
}

function findHeaderByColumnIndex(sheet, columnIndex){
  Logger.log("findHeaderByColumnIndex() function execute");
  var frozenHeaderRange = new Array();
  var frozenHeaderValues = new Array();
  var headerValue = "";
  var frozenRows = sheet.getFrozenRows();
  
  // get the Header value
  frozenRowValues = sheet.getRange(1, columnIndex, frozenRows, 1).getValues();
  return frozenRowValues;
}

function onOpen(){
  //var ss = SpreadsheetApp.openById("spreadsheetID");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "Refresh Colors", functionName: "changeBgColorByStatus"});
  menuEntries.push({name: "Clear All List", functionName: "clearAllSelectionList"});
  ss.addMenu("KeithBox3.2 Log", menuEntries);
}

function testing(){
  var ss = SpreadsheetApp.openById(spreadsheetID);
  var sheet = ss.getSheetByName(logSheetName);
  var range = sheet.getRange(3, 1, 1, sheet.getMaxColumns());
  var rules = range.getDataValidations();
  var nextRange = sheet.getRange(4, 1, 1, sheet.getMaxColumns());
  nextRange.setDataValidations(rules);
}

function copyRuleToNextRow(sheet, copyThisRow){
  Logger.log("copyRuleToNextRow() function execute");
  var range = sheet.getRange(copyThisRow, 1, 1, sheet.getMaxColumns());
  var rules = range.getDataValidations();
  var nextRange = sheet.getRange(++copyThisRow, 1, 1, sheet.getMaxColumns());
  Logger.log("Row:"+copyThisRow+" rules copy to next row");
  nextRange.setDataValidations(rules);
}

function clearAllSelectionList(){
  var ss = SpreadsheetApp.openById(spreadsheetID);
  var sheet = ss.getSheetByName(logSheetName);
}