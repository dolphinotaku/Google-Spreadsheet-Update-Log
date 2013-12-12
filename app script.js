var spreadsheetID = "0AlaVan9pZtAzdEF5Wm9HQzFiTlpNQVF4a3hmWDJxSGc";
var logSheetName = "Log";
var customizeStatusColorSheetName = "Status Color";
var statusChangeColumnName = "Status";

var addTodayWhenEdit = [ // [whenEditCell, addTodayCell]
  ["Report By", "Report Date"],
  ["Completed By", "Completed Date"]
];

var backgroundColorPriority = [
  // status: Top priority <--------------------> Low priority
  ["tailor make", "hardcode", "holding", "follow up", "misreporting", "cancelled", "pending", "release", "done"],
  // stauts background color
  ["#d9d2e9", "#f4cccc", "#f4cccc", "#c9daf8", "#efefef", "#efefef", "#fff2cc", "#d9ead3", "#d9ead3"]
];

// Prepare a New Line
var autoGenKeyColumn = ["#.", ""]; // [columnHeaderName, autoGenFormat]
var dateValidation = [ // instance, validationMethods, argument [, argument]
  [addTodayWhenEdit[1][1], "DateOnOrAfter", addTodayWhenEdit[0][1]] // "Completed Date" must OnOrAfter "Report Date"
];
var prepareHowManyRows = "20";

Logger.clear();

function changeBgColorByStatus() {
  var ss = SpreadsheetApp.openById("spreadsheetID");
  var sheet = ss.getSheetByName(logSheetName);
  var customColorSheet = ss.getSheetByName(customizeStatusColorSheetName);
  //var sheet = ss.get
  //sheet = ss.getSheets()[0];
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();  
  var maxRows = parseInt(sheet.getMaxRows())
  var maxColumns = parseInt(sheet.getMaxColumns());
  Logger.log("Current Spreadsheet max rows: "+maxRows);
  Logger.log("Current Spreadsheet max columns: "+maxColumns);
  var cell, row, cellValue, currentBackgroudColor = "";
  
  // white,undefined  	= nothing
  // Hardcode 			= #f4cccc
  // Holding 			= #f4cccc
  // Misreporting 		= #efefef
  // Cancelled 			= #efefef
  // Pending 			= #fff2cc
  // Release 			= #d9ead3
  // Done 				= #d9ead3 
  
  var colors = new Array();
  var isCurCellBgColor=false, isCurCellValueEmpty=false;
  
  for(var rowIndex = frozenRows+1; rowIndex<maxRows; rowIndex++){
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
  var cell, row, cellValue, currentBackgroudColor = "";
  var colors = new Array();
  var isCurCellBgColor=false, isCurCellValueEmpty=false;

  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();  
  var maxRows = parseInt(sheet.getMaxRows())
  var maxColumns = parseInt(sheet.getMaxColumns());
  
  var i= rowIndex;
  
  //find the statusChangeColumnName column index
  for(var fRows = 1; fRows<=frozenRows;fRows++){
    frozenHeaderRange = sheet.getRange(fRows, 1, 1, maxColumns);
    frozenHeaderValues = frozenHeaderRange.getValues();
    var statusColumnIndex = frozenHeaderValues[0].indexOf(statusChangeColumnName);
    if(statusColumnIndex>=0)
      break;
  }
  if(statusColumnIndex==-1)
    return;
  else{
    statusColumnIndex+=1;
  }
  
  Logger.log("Row:"+rowIndex+" column:"+statusColumnIndex);
  
  cell = sheet.getRange(rowIndex, statusColumnIndex);
  cellValue = cell.getValue().toLowerCase();
  currentBackgroudColor = cell.getBackground().toLowerCase();
  isCurCellBgColor = currentBackgroudColor=="white";
  isCurCellValueEmpty = cellValue=="";
  
  
  //var ss = SpreadsheetApp.openById("spreadsheetID");
  //var statusColorSheet = ss.getSheetByName(customizeStatusColorSheetName);
  if(statusColorSheet!=null)
    backgroundColorPriority = customStatusColor(statusColorSheet);
  
  
  Logger.log("isCurCellValueEmpty:"+isCurCellValueEmpty+" ,cellValue:"+cellValue);
    if(!isCurCellValueEmpty){
      var isStautsFound = backgroundColorPriority[0].indexOf(cellValue);
      var isFillColor = false;
      
      if(isStautsFound >= 0){
          isFillColor = true;
      }else{
        Logger.log("Rows: "+i)
        for(var m=0; m<backgroundColorPriority[0].length; m++){
          isStautsFound = cellValue.indexOf(backgroundColorPriority[0][m]);
          Logger.log("Cell Value: "+cellValue+" = "+backgroundColorPriority[0][m]+" : "+m);
          if(isStautsFound >=0){
            isFillColor = true;
            isStautsFound = m;
            break;
          }
        }
      }
      if(isFillColor){
        if(currentBackgroudColor==backgroundColorPriority[1][isStautsFound]){
          return;
        }
          colors[0]  = new Array(maxColumns);
          var bgColorchangeToThis = backgroundColorPriority[1][isStautsFound];
          var rgbColors = new RGBColor(bgColorchangeToThis);
        if(rgbColors.ok){
          bgColorchangeToThis = rgbColors.toHex();
        }
        
          for(var k=0; k<maxColumns; k++){
            colors[0][k] = bgColorchangeToThis;
          }
          row = sheet.getRange(i, 1, 1, maxColumns);
          row.setBackgrounds(colors);
        
          Logger.log("Change row "+i+" background color: "+currentBackgroudColor+" change to "+bgColorchangeToThis);
        /*if(colors.length<=7)
          row.setBackgrounds(colors);
        else{
          var rgbColors = new RGBColor(colors);
          row.setBackgroundRGB(rgbColors.r, rgbColors.g, rgbColors.b);
        }*/
      }else{
        row = sheet.getRange(i, 1, 1, maxColumns);
        colors[0]  = new Array(maxColumns);
        for(var k=0; k<maxColumns; k++){
          colors[0][k] = "white";
        }
        row.setBackgrounds(colors);
      }
    }else{
      var isBackgroundColor = backgroundColorPriority[1].indexOf(currentBackgroudColor);
      if(isBackgroundColor<0)
        Logger.log("Row E"+i+", isBgColorWhite: "+isCurCellBgColor+" && isCellNotEmpty: "+!isCurCellValueEmpty);{
        Logger.log("i: "+currentBackgroudColor+", cellValue: "+cellValue);
      }
      row = sheet.getRange(i, 1, 1, maxColumns);
        colors[0]  = new Array(maxColumns);
      for(var k=0; k<maxColumns; k++){
        colors[0][k] = "white";
      }
      row.setBackgrounds(colors);
    }
}

function printTodayAt(sheet, r){
  Logger.log("printTodayAt() function execute");
  var targetValue = r.getValue()
  var targetCell = r;
  if(targetValue==""){
    var now = new Date();
    Logger.log("Insert "+now.toLocaleDateString()+" into X:"+targetCell.getColumn()+" Y:"+targetCell.getRow());
    now = now.toLocaleDateString();
    targetCell.setNumberFormat("yyyy-MM-dd");
    targetCell.setValue(now);
  }
}

function onEdit(event){
  Logger.log("onEdit() function execute");
  var frozenHeaderRange = new Array();
  var frozenHeaderValues = new Array();
  
  // get sheet
  //var ss = SpreadsheetApp.openById("spreadsheetID");
  //var sheet = ss.getSheetByName(logSheetName);
  var sheet = event.source.getSheetByName(logSheetName);
  //var sheet = event.source.getActiveSheet();
  var customColorSheet = event.source.getSheetByName(customizeStatusColorSheetName);
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();  
  var maxRows = parseInt(sheet.getMaxRows())
  var maxColumns = parseInt(sheet.getMaxColumns());
  
  // get active range
  var r = event.range;
  //var r = event.source.getActiveRange();
  
  // get the top left Coordinate of range
  var rowIndex = r.getRowIndex();
  var columnIndex = r.getColumn();
  
  Logger.log("Edit Range start at X:"+columnIndex+" Y: "+rowIndex);
  Logger.log("First "+frozenRows+" row(s) are Igrone(Frozen)");
  
  if(frozenRows<=0){
    Logger.log("No row(s) frozen, assume the Header at the first row.");
    frozenRows = 1;
  }else if(rowIndex<frozenRows){
    Logger.log("Edit in frozen area.");
    return;
  }
  
  var editCellHeader = findHeaderByColumnIndex(sheet, columnIndex, rowIndex);
  

  for(var i=0; i<editCellHeader.length; i++){
    // is editing status column?{
    Logger.log("editCellHeader[i]:"+editCellHeader[i]+" statusChangeColumnName:"+statusChangeColumnName);
    
    if(editCellHeader[i].indexOf(statusChangeColumnName)>=0){
      changeARowBgColorByStatus(customColorSheet, sheet);
      break;
    }
    // is trigger insert today date? 
    for(var j=0; i<addTodayWhenEdit.length; j++){
      if(editCellHeader[i].indexOf(addTodayWhenEdit[j][0])>=0){
          var r = sheet.getRange(rowIndex, findColumnIndexByHeader(sheet, addTodayWhenEdit[j][0]));
        
          printTodayAt(sheet, r);
      }
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
  var maxRows = parseInt(sheet.getMaxRows())
  var maxColumns = parseInt(sheet.getMaxColumns());
  
  // This represents ALL the data
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  var readRows = values.length;
  Logger.log("Read rows Total:"+readRows);
  
  var isValue, isPriority;
  var customColorPriority = new Array();
  customColorPriority[0] = new Array();
  customColorPriority[1] = new Array();
  
  var tempStatus="";
  var tempBackgroudColor="";
  var tempColor="";
  var theProiority="";
  //backgroundColorPriority
  Logger.log("frozenRows:"+frozenRows+" frozenCols:"+frozenCols+" maxRows:"+maxRows+" maxColumns:"+maxColumns);
  for(var startRow = frozenRows; startRow<readRows; startRow++){
    for(var j=0; j<maxColumns; j++){
      isValue = false;
      isProiority = true;
      if(values[startRow][j]){
        Logger.log("row:"+(startRow+frozenRows)+" col:"+(j+1)+" value:"+values[startRow][j]);
        isValue = true;
      }else{
        Logger.log("row:"+(startRow+frozenRows)+" col:"+(j+1)+" value was undefined");
        if(j==0){
          Logger.log("No status specify, skill to next row.");
          break;
        }else if(j==2){
          Logger.log("Priority have no set, routed to the tail.");
          isProiority = false;
        }
      }
      switch(j){
        case 0: // Status column
          tempStatus = values[startRow][j].toLowerCase();
          break;
        case 1: // Color in Hex/RGB column
          tempColor = values[startRow][j];
          var isColorValidate = new RGBColor(tempColor);
          Logger.log("The text '"+tempColor+"' is a valid color: "+isColorValidate.ok);
          if(isColorValidate.ok)
              tempBackgroudColor = tempColor;
          else{
            tempBackgroudColor = sheet.getRange(startRow+1, j+1).getBackground();
          }
          Logger.log("Read color row "+(startRow+1)+": "+tempBackgroudColor);
          /*
          if(!tempColor){
            // no value get the backgroup color
            tempBackgroudColor = sheet.getRange(startRow+1, j+1).getBackground();
            Logger.log(tempBackgroudColor);
          }else{
            //var isColorValidate = colorInTextValidator(tempColor);
            var isColorValidate = new RGBColor(tempColor);
            if(isColorValidate.ok)
              tempBackgroudColor = tempColor;
          }
          */
          break;
        case 2:
          if(isValue){
            if(!isNaN(parseInt(values[startRow][j]))){
              if(parseInt(values[startRow][j])>0){
              theProiority = parseInt(values[startRow][j])-1;
              // if Proiority is numeric
              customColorPriority[0][theProiority] = tempStatus;
              customColorPriority[1][theProiority] = tempBackgroudColor;
              }
            }else{
              // if Proiority is non-numeric, push to the end of array
              customColorPriority[0].push(tempStatus);
              customColorPriority[1].push(tempBackgroudColor);
            }
          }
          break;
      }
    }
  }
  
  for(var i=0; i<backgroundColorPriority[0].length; i++){
    Logger.log("status:"+i+" "+backgroundColorPriority[0][i]+" color:"+i+" "+backgroundColorPriority[1][i]);
  }
  
  return customColorPriority;
}

// replace by rgbcolor.js
function colorInTextValidator(color){
  var isColorValid = /^#[0-9A-F]{6}$/i.test(color);
  Logger.log( color +" is a color = "+isColorValid );
  return isColorValid;
  // /(^#[0-9A-F]{6}$)|(^#[0-9A-F]{3}$)/i.test('#ac3') // for #f00 (Thanks Smamatti)
}

function prepareANewLine(){
}
function prepareNewLines(){
  //
  
  autoGenNum();
  
}

function findColumnIndexByHeader(sheet, header, rowIndex){
  Logger.log("findColumnIndexByHeader() function execute");
  var frozenHeaderRange = new Array();
  var frozenHeaderValues = new Array();
  var isHeaderFound = false;
  var headerFoundAtColumn = -1;
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();
  var maxRows = parseInt(sheet.getMaxRows())
  var maxColumns = parseInt(sheet.getMaxColumns());
  
  // get the Header(FrozenRows) cell value
  for(var rowPointer = 1; rowPointer<=frozenRows;rowPointer++){
    frozenRowRange = sheet.getRange(rowPointer, 1, 1, maxColumns);
    frozenRowValues = frozenRowRange.getValues();
    headerFoundAtColumn = frozenRowValues[0].indexOf(header);
    //Logger.log("frozenRowValues:"+frozenRowValues[0]+" headerFoundAtColumn:"+headerFoundAtColumn);
    if(headerFoundAtColumn>=0){
      isHeaderFound = true;
      break
    }
  }
  return headerFoundAtColumn;
}

function findHeaderByColumnIndex(sheet, columnIndex, rowIndex){
  Logger.log("findHeaderByColumnIndex() function execute");
  var frozenHeaderRange = new Array();
  var frozenHeaderValues = new Array();
  var headerValue = "";
  // get sheet Properties
  var frozenRows = sheet.getFrozenRows();
  var frozenCols = sheet.getFrozenColumns();
  var maxRows = parseInt(sheet.getMaxRows())
  var maxColumns = parseInt(sheet.getMaxColumns());
  
  //Logger.log("frozenRows:"+frozenRows+" ,frozenCols:"+frozenCols);
  
  // get the Header value
  frozenRowValues = sheet.getRange(1, columnIndex, frozenRows, 1).getValues();

  return frozenRowValues;
}

function onOpen(){
  //var ss = SpreadsheetApp.openById("spreadsheetID");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user selects "addMenuExample" menu, and clicks "Menu Entry 1", the function function1 is executed.
  menuEntries.push({name: "Refresh Colors", functionName: "changeBgColorByStatus"});
  //menuEntries.push({name: "Menu Entry 2", functionName: "function2"});
  ss.addMenu("KeithBox3.2 Log", menuEntries);
  
  // Use customize status color
  //var s2 = SpreadsheetApp.getActiveSheet();
  /*
  var statusColorSheet = ss.getSheetByName(customizeStatusColorSheetName);
  customStatusColor(statusColorSheet);
  */
}