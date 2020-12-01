function updateSheetList() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  var rangeCurrentSheetList = mainSheet
    .getRange(address_firstCell_A1_Style.sheetList)
    .getDataRegion();

  rangeCurrentSheetList.clear();

  Make_list_For_all_sheets(mainSheet, address_firstCell_A1_Style.sheetList);

  FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.sheetList,
    1
  ).applyRowBanding(); //Banding rows for sheetList
}

function changeSheetsName() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  var rangeSheetListFirstCell = mainSheet.getRange(
    address_firstCell_A1_Style.sheetList
  );
  var rangeSheetListArray = FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.sheetList,
    2
  ).getValues();

  var sheetNameIndex =
    address_firstCell_A1_Style.sheetListTitleColIndex.name[0];
  var sheetNewNameIndex =
    address_firstCell_A1_Style.sheetListTitleColIndex.newName[0];

  var sheetCurrentNameArray = [];
  var sheetNewNameArray = [];
  for (var i = 0; i < rangeSheetListArray.length; i++) {
    if (
      rangeSheetListArray[i][sheetNewNameIndex] == "-" ||
      rangeSheetListArray[i][sheetNewNameIndex] == ""
    ) {
      var newNameFinal = rangeSheetListArray[i][sheetNameIndex];
    } else {
      var newNameFinal = rangeSheetListArray[i][sheetNewNameIndex];
    }
    sheetCurrentNameArray.push(rangeSheetListArray[i][sheetNameIndex]);
    sheetNewNameArray.push(newNameFinal);
  }

  for (var i = 0; i < rangeSheetListArray.length; i++) {
    spreadsheet
      .getSheetByName(sheetCurrentNameArray[i])
      .setName(sheetNewNameArray[i]);
  }

  updateSheetList();
}
