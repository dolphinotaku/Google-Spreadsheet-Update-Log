Logger.clear();
var backgroundColorPriority = [
  ["tailor make", "hardcode", "holding", "follow up", "misreporting", "cancelled", "pending", "release", "done"],
  ["#d9d2e9", "#f4cccc", "#f4cccc", "#c9daf8", "#efefef", "#efefef", "#fff2cc", "#d9ead3", "#d9ead3"]
];
var addTodayWhenEdit = [
  ["Report By", "Report Date"],
  ["Completed By", "Completed Date"]
];

var statusChangeColumnName = "Status";
var logSheetName = "Log";
var customizeStatusColorSheetName = "Status Color";

function changeBgColorByStatus() {
  var ss = SpreadsheetApp.openById("0AlaVan9pZtAzdEF5Wm9HQzFiTlpNQVF4a3hmWDJxSGc");
  var sheet = ss.getSheetByName(logSheetName);


  var customColorSheet = ss.getSheetByName(customizeStatusColorSheetName);
  
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
  //Logger.log("Cell color: "+sheet.getRange("E79").getBackground());  
  
  var colors = new Array();
  var isCurCellBgColor=false, isCurCellValueEmpty=false;
  //var r = sheet.getRange("E3");
  
  for(var rowIndex = frozenRows+1; rowIndex<maxRows; rowIndex++){
    //r = sheet.getRange("E"+rowIndex);
    changeARowBgColorByStatus(customColorSheet, sheet, rowIndex);
  }
  
  // other ways to get cell data
  // var values = SpreadsheetApp.getActiveSheet().getDataRange().getValues()
  // values[0][2]
}

// http://stackoverflow.com/questions/3703676/google-spreadsheet-script-to-change-row-color-when-a-cell-changes-text

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
          for(var k=0; k<maxColumns; k++){
            colors[0][k] = backgroundColorPriority[1][isStautsFound];
          }
          row = sheet.getRange(i, 1, 1, maxColumns);
        
          Logger.log("Change "+i+" row background color: "+currentBackgroudColor+" change to "+backgroundColorPriority[1][isStautsFound]);
          row.setBackgrounds(colors);
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
  Logger.log("start printTodayAt");
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
  Logger.log("onEdit triggered");
  var frozenHeaderRange = new Array();
  var frozenHeaderValues = new Array();
  
  // get sheet
  var sheet = event.source.getActiveSheet();
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
  
  if(rowIndex<frozenRows)
    return;
  
  // is status change
  
  for(var fRows = 1; fRows<=frozenRows;fRows++){
    frozenHeaderRange = sheet.getRange(fRows, 1, 1, maxColumns);
    frozenHeaderValues = frozenHeaderRange.getValues();
    
    var editColumnHeader = frozenHeaderValues[0][columnIndex-1].toLowerCase();
    // is trriger to change color?
    if(statusChangeColumnName.toLowerCase() == editColumnHeader){
      changeARowBgColorByStatus(customColorSheet, sheet, rowIndex);
      break;
    }
    
    // is trriger to add to today?
    for(var i=0; i<addTodayWhenEdit.length; i++){
      var triggerHeader = addTodayWhenEdit[i][0].toLowerCase();
      //Logger.log("triggerHeader: "+triggerHeader+", editColumnHeader: "+editColumnHeader);
      
      if(triggerHeader==editColumnHeader){
        var targetColumnIndex = frozenHeaderValues[0].indexOf(addTodayWhenEdit[i][1]);
        if(targetColumnIndex>=0){
          var r = sheet.getRange(rowIndex, targetColumnIndex+1);
           //Logger.log("r: "+r.getValue())
          printTodayAt(sheet, r);
        }
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
  
  var tempValue="";
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
        case 0:
          tempValue = values[startRow][j].toLowerCase();
          break;
        case 1:
          if(!values[startRow][j]){
            // no value get the backgroup color
            tempColor = sheet.getRange(startRow+1, j+1).getBackground();
            Logger.log(tempColor);
          }else{
          }
          break;
        case 2:
          if(isValue){
            if(!isNaN(parseInt(values[startRow][j]))){
              if(parseInt(values[startRow][j])>0){
              theProiority = parseInt(values[startRow][j])-1;
              // if value is numeric
              customColorPriority[0][theProiority] = tempValue;
              customColorPriority[1][theProiority] = tempColor;
              }
            }else{
              // if value is non-numeric, push to the end of array
              customColorPriority[0].push(tempValue);
              customColorPriority[1].push(tempColor);
            }
          }
          break;
      }
    }
  }
  /*
  for(var i=0; i<backgroundColorPriority[0].length; i++){
    Logger.log("status:"+i+" "+backgroundColorPriority[0][i]+" color:"+i+" "+backgroundColorPriority[1][i]);
  }
  */
  return customColorPriority;
}

function onOpen(){
  //var ss = SpreadsheetApp.openById("0AlaVan9pZtAzdEF5Wm9HQzFiTlpNQVF4a3hmWDJxSGc");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user selects "addMenuExample" menu, and clicks "Menu Entry 1", the function function1 is executed.
  menuEntries.push({name: "Refresh Colors", functionName: "changeBgColorByStatus"});
  //menuEntries.push({name: "Menu Entry 2", functionName: "function2"});
  ss.addMenu("KeithBox3.2 Log", menuEntries);
  
  // Use customize status color
  var statusColorSheet = ss.getSheetByName(customizeStatusColorSheetName);
  customStatusColor(statusColorSheet);
}
