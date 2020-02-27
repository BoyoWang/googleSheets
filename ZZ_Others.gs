function Monotaro_JointShts(){
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  
  //Make Sheet List
  Make_list_For_all_sheets(mainSheet, address_firstCell_A1_Style.sheetList);

  //Banding rows for sheetList
//  returnListRangeExcludeTopRows(mainSheet, 
//                                address_firstCell_A1_Style.sheetList, 
//                                1).applyRowBanding();
  
  //Create New Sheet
  spreadsheet.insertSheet(SpreadsheetApp.getActive().getNumSheets());
  var newSheet = spreadsheet.getActiveSheet();
  newSheet.setName("Combined_"+ Date.now());
  
  
  var rangeSheetListFirstCell = mainSheet.getRange(address_firstCell_A1_Style.sheetList);
  var rangeSheetListArray = returnListRangeExcludeTopRows(mainSheet, address_firstCell_A1_Style.sheetList, 2).getValues();
  
  Logger.log(rangeSheetListArray);
  
  var sheetNameIndex = address_firstCell_A1_Style.sheetListTitleColIndex.name[0];
  
  //Original looping all sheets code:
  //for(var i = 0; i < rangeSheetListArray.length; i++){
  for(var i = 2; i <= 8; i++){
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


function Monotarou_DoSthAfterPaste1(){
  var spreadsheet = SpreadsheetApp.getActive();
  var targetSheet = spreadsheet.getActiveSheet();
  
  var newSheetNameInput = Browser.inputBox("Enter the purchase date:");
  if (newSheetNameInput == ""){newSheetNameInput = "Unknown Date"};
  
  //Delete usless rows
  var DataStart = findCellByText_ReturnRange(targetSheet, "商品名");
  var rowsToDelete = DataStart.getRow() - 1;
  if (rowsToDelete >= 1) {
    targetSheet.deleteRows(1, rowsToDelete);
  };
  //Create new Sheet
  spreadsheet.insertSheet(SpreadsheetApp.getActive().getNumSheets());
  var newSheet = spreadsheet.getActiveSheet();
  var currentSheetName = targetSheet.getName();
  var targetSheetTempName = currentSheetName + '_t';
  var newSheetName = newSheetNameInput;
  
  targetSheet.setName(targetSheetTempName);
  newSheet.setName(newSheetName);
  
  //copy needed columns
  var columnsToKeepText = [["商品名", 'General'], 
                           ["単価", 'General'],
                           ["数量", 'General'],
                           ["金額(税抜)", 'General']
                          ];
  for (j = 0;j < columnsToKeepText.length;j++){
    var rangeCopyFrom = get_ColRange_In_TitleRow(columnsToKeepText[j][0], targetSheet);
    rangeCopyFrom.setNumberFormat(columnsToKeepText[j][1]);
    //Set format before copy
    rangeCopyFrom.copyTo(newSheet.getRange(1, j+1));
    
  };
  
  
  //Delete old and rename new sheet
  spreadsheet.deleteSheet(targetSheet);
};

function Monotarou_DoSthAfterPaste2(){
  var spreadsheet = SpreadsheetApp.getActive();
  var targetSheet = spreadsheet.getActiveSheet();
  
  
  var dataFirstRow = findCellByText_ReturnRange(targetSheet, "商品名").getRow() + 1;
  var dataLastRow = findCellByText_ReturnRange(targetSheet, "小計").getRow() - 1;
  
  var rowAmt = dataLastRow - dataFirstRow + 1;
  
  var allData = [];
  
  for(var i = 0; i < (rowAmt); i++){
    var orderDate = targetSheet.getName();
    var orderCode = "'" + transformOrderCode(targetSheet.getRange(dataFirstRow + i, 1).getValue());
    var orderPerson = "Poyao Wang";
    var productName = transformProductName(targetSheet.getRange(dataFirstRow + i, 1).getValue());
    var singlePrice = removeYenAndComma(targetSheet.getRange(dataFirstRow + i, 2).getValue());
    var orderAmt = targetSheet.getRange(dataFirstRow + i, 3).getValue();
    var totalPrice = removeYenAndComma(targetSheet.getRange(dataFirstRow + i, 4).getValue());
    
    var element = [orderDate, orderCode, orderPerson, productName, singlePrice, orderAmt, totalPrice];
    allData.push(element);
  };
  
  for(var i = 0; i < allData.length; i++){
    Logger.log(allData[i]);
  };
  
  targetSheet.getDataRange().clearContent();
  var rangeToPasteFirstCell = targetSheet.getRange("A1");
  var rangeToPaste = targetSheet.getRange(
    rangeToPasteFirstCell.getRow(),    //row 
    rangeToPasteFirstCell.getColumn(), //column 
    allData.length,                    //numRows 
    allData[0].length);                //numColumns)
  
  rangeToPaste.setValues(allData);

  function transformProductName(input){
    var firstRow = input.split("\n")[0];
    var productName = firstRow.split("注文コード")[0];
    return productName;
  };
  
  function transformOrderCode(input){
    var firstRow = input.split("\n")[0];
    var orderCode = firstRow.split("注文コード")[1];
    return orderCode;
  };
  
  function removeYenAndComma(input){
    input = input.replace(/￥/g, "");
    input = input.replace(/,/gi, "");
    return input;
  };
  
};