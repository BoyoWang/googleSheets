function S06_readSheetsReturnSheetNameArray() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  var rangeSheetListFirstCell = mainSheet.getRange(
    address_firstCell_A1_Style.sheetList
  );
  var rangeSheetListArray = returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.sheetList,
    2
  ).getValues();

  var sheetNameIndex =
    address_firstCell_A1_Style.sheetListTitleColIndex.name[0];

  var sheetNameArray = [];
  for (var i = 0; i < rangeSheetListArray.length; i++) {
    sheetNameArray.push(rangeSheetListArray[i][sheetNameIndex]);
  }

  //  Logger.log(sheetNameArray);

  return sheetNameArray;
}

function S06_readColsReturnObject(/*sheet*/ sheet) {
  var spreadsheet = SpreadsheetApp.getActive();
  var targetSheet = spreadsheet.getActiveSheet();

  //  targetSheet = sheet;

  var dataRegionArray = targetSheet.getDataRange().getValues();

  var dataStartRow;

  for (var i = 0; i < dataRegionArray.length; i++) {
    if (!isNaN(dataRegionArray[i][0])) {
      dataStartRow = i + 1;
      break;
    }
  }

  var dataEndRow = dataRegionArray.length;

  for (var j = 0; j < dataRegionArray[0].length; j++) {}

  var sheetName = targetSheet.getName();
  var xAxisDataSerieName = "";
  var dataSeriesName = "";
  var dataSeriesRangeA1Style = "";
}
