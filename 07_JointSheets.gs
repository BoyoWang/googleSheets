function S07_ReadSheets_MakeSheetList(){
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  
  //Make Sheet List
  Make_list_For_all_sheets(mainSheet, address_firstCell_A1_Style.sheetList);
}

  
function S07_ReadSheets_MakeSheetList(iShtStart, iShtEnd, XorY){

  //Create New Sheet
  spreadsheet.insertSheet(SpreadsheetApp.getActive().getNumSheets());
  var newSheet = spreadsheet.getActiveSheet();
  newSheet.setName("Combined_"+ Date.now());
    
  var rangeSheetListFirstCell = mainSheet.getRange(address_firstCell_A1_Style.sheetList);
  var rangeSheetListArray = returnListRangeExcludeTopRows(mainSheet, address_firstCell_A1_Style.sheetList, 2).getValues();
  
  Logger.log(rangeSheetListArray);
  
  var sheetNameIndex = address_firstCell_A1_Style.sheetListTitleColIndex.name[0];
  
  //Original looping all sheets code:
  for(var i = iShtStart; i <= iShtEnd; i++){
    var shtCopyFrom = spreadsheet.getSheetByName(rangeSheetListArray[i][sheetNameIndex]);
    var shtCopyTo = newSheet;
    Logger.log(shtCopyFrom.getSheetName());
    
    var rangeCopyFrom = shtCopyFrom.getDataRange();
    var rangeCopyFromArry = rangeCopyFrom.getValues();
    
    var rangeCopyFromColSize = rangeCopyFromArry[0].length;
    var rangeCopyFromRowSize = rangeCopyFromArry.length;
    
    var rangeCopyToFirstCell = newSheet.getRange(newSheet.getLastRow() + 1, 1);
    var rangeCopyTo = newSheet.getRange(rangeCopyToFirstCell.getRow(), 
                                        rangeCopyToFirstCell.getColumn(), 
                                        rangeCopyFromRowSize, 
                                        rangeCopyFromColSize);
    rangeCopyTo.setValues(rangeCopyFromArry);
  };
};
