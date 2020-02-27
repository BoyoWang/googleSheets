function G021_Import_All_CSV(){
  importCSVExcuteAll();
};

function G022_Reset_File(){
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  // if mainSheet doesn't exist create it
  if (!mainSheet) {
    spreadsheet.insertSheet(name_importantSheets.mainSheet);
    mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
    Logger.log("mainSheet created.");
  };
  resetFile();
};

function G03_Apply_Actions_To_Sheets(){
  ApplyActionToAllSheets();
};

function G041_Update_FileList(){
  updateFileList();
};

function G042_Change_FileName(){
  changeFileName();
};

function G051_Update_SheetList(){
  updateSheetList();
};

function G052_Change_SheetName(){
  changeSheetsName();
};

function GZZ1_MonotarouExcuteAll(){
  var spreadsheet = SpreadsheetApp.getActive();
  var targetSheet = spreadsheet.getActiveSheet();
  Monotarou_DoSthAfterPaste1();
  Monotarou_DoSthAfterPaste2();
};

function GZZ2_ResetFile(){
  
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  
  // if mainSheet doesn't exist create it
  if (!mainSheet) {
    spreadsheet.insertSheet(name_importantSheets.mainSheet);
    mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
    Logger.log("mainSheet created.");
  };
  
  resetFile();
  
};

function GZZ3_JointShts(){
  Monotaro_JointShts();
};

