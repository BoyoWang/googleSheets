//ZZ_For_CSV
var folderGoogleDriveID = "1DamhKC87tu8yx9iGbwMXnYP8FcjMPi_g";

//20191210_Burst_P
//var folderGoogleDriveID = "1Qeb5fVewYN3IHn5x_H2Ie-9Z4_8aA0My";

var address_firstCell_A1_Style = {
  csvFileList: "A1",
  csvFileList_MainTitleText: "CSV_Data_List",
  csvFileListTitleColIndex: {
    name: [0, "Name"],
    size: [2, "Size"],
    googleDriveID: [1, "ID"],
    newName: [3, "NewName"],
  },
  sheetList: "F1",
  sheetList_MainTitleText: "Sheet_List",
  sheetListTitleColIndex: {
    index: [0, "Index"],
    name: [1, "Name"],
    newName: [2, "NewName"],
  },
  dataSeriesInfosList: "J1",
  dataSeriesInfosList_MainTitleText: "Data_Series_Infos",
  dataSeriesInfosListTitleColIndex: {
    index: [0, "Index"],
    sheetName: [1, "Index"],
    serieName: [2, "SerieName"],
    X_AxisA1Address: [3, "X_AxisA1Address"],
    serieA1Address: [4, "SerieA1Address"],
  },
};

var name_importantSheets = {
  mainSheet: "Main",
};
