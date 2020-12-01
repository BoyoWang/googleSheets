function updateFileList() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  var rangeCurrentFileList = mainSheet
    .getRange(address_firstCell_A1_Style.csvFileList)
    .getDataRegion();

  rangeCurrentFileList.clear();

  list_all_files_inside_folder(
    folderGoogleDriveID,
    mainSheet,
    address_firstCell_A1_Style.csvFileList
  );
  FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.csvFileList,
    1
  ).applyRowBanding(); //Banding rows for csvFileList
  Sort_CSV_FileList();
}

function changeFileName() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  var rangeFileListFirstCell = mainSheet.getRange(
    address_firstCell_A1_Style.csvFileList
  );
  var rangeFileListArray = FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.csvFileList,
    2
  ).getValues();

  var fileNameIndex =
    address_firstCell_A1_Style.csvFileListTitleColIndex.name[0];
  var fileIDIndex =
    address_firstCell_A1_Style.csvFileListTitleColIndex.googleDriveID[0];
  var fileNewNameIndex =
    address_firstCell_A1_Style.csvFileListTitleColIndex.newName[0];

  var fileIDArray = [];
  var fileNewNameArray = [];
  for (var i = 0; i < rangeFileListArray.length; i++) {
    if (
      rangeFileListArray[i][fileNewNameIndex] == "-" ||
      rangeFileListArray[i][fileNewNameIndex] == ""
    ) {
      var newNameFinal = rangeFileListArray[i][fileNameIndex];
    } else {
      var newNameFinal = rangeFileListArray[i][fileNewNameIndex];
    }
    fileIDArray.push(rangeFileListArray[i][fileIDIndex]);
    fileNewNameArray.push(newNameFinal);
  }

  for (var i = 0; i < rangeFileListArray.length; i++) {
    var file = DriveApp.getFileById(fileIDArray[i]);
    file.setName(fileNewNameArray[i]);
  }

  updateFileList();
}
