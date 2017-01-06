function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Generate Loan Overview", functionName: "copyToMasterSheet"},
                       {name: "Show Filters", functionName: "showDialog"}];
  ss.addMenu("Tasks", csvMenuEntries);
}

function copyToMasterSheet() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var loanOverviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loan overview");
  var lastRow = loanOverviewSheet.getLastRow();
  if(lastRow != 1) {
    loanOverviewSheet.deleteRows(2, lastRow)
  }
  var total = 0;
  for(var sheetNumber = 1; sheetNumber < sheets.length; sheetNumber++) {
    var sheet = sheets[sheetNumber];
    if(!sheet.isSheetHidden() && sheet.getName() != "Pending") {
      var loanDetails = sheet.getRange("B7:B9").getDisplayValues();
      loanOverviewSheet.appendRow([loanDetails[1][0], loanDetails[0][0], loanDetails[2][0]]);
      total = total + Number(sheet.getRange("B9").getValue());
    }
  }
  lastRow = loanOverviewSheet.getLastRow();
  loanOverviewSheet.showRows(1, lastRow);
  loanOverviewSheet.appendRow(["", "Total", "$" + total])
  lastRow = loanOverviewSheet.getLastRow();
  var totalCell = loanOverviewSheet.getRange(lastRow, 3);
  totalCell.setNumberFormat("$#,##0.00");
}

function showDialog() {
   var html = HtmlService.createTemplateFromFile('Page').evaluate().setTitle("Filter - Company Names");
   SpreadsheetApp.getUi().showSidebar(html);
}

function hideRowsAndCalculateTotal(formData) {
  var selectedValues = [];
  for(var data in formData){
    if(data.substr(0, 2) == 'ch')
      selectedValues.push(formData[data]);
  }
  if(selectedValues.length == 0) {
    clearFilters();
    return;
  }
  var loanOverviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loan overview");
  var lastRow = loanOverviewSheet.getLastRow();
  var total = 0;
  var exists;
  for (var rowNumber = 2; rowNumber < lastRow; rowNumber++) {
    var companyName = loanOverviewSheet.getRange(rowNumber, 2).getDisplayValue();
    exists = false;
    for(var index = 0; index < selectedValues.length; index++) {
      if(selectedValues[index] == companyName) {
        exists = true;
        loanOverviewSheet.showRows(rowNumber)
        total = total + Number(loanOverviewSheet.getRange(rowNumber, 3).getValue());
        break;
      }
    }
    if (exists == false)
      loanOverviewSheet.hideRows(rowNumber);
  }
  loanOverviewSheet.getRange(lastRow, 3).setValue(total);
  SpreadsheetApp.getActiveSpreadsheet().toast("Completed", "Status", 2);
}

function getUniqueValues() {
  var loanOverviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loan overview");
  var lastRow = loanOverviewSheet.getLastRow();
  var values = loanOverviewSheet.getRange(2, 2, lastRow - 2, 1).getDisplayValues();
  var companyNames = [];
  var exists;
  for(var x = 0; x < values.length; x++){
    exists = false;
    for(var y = 0; y < companyNames.length; y++){
      if(values[x].toString() === companyNames[y].toString()){
         exists = true;
         break;
      }
    }
    if(!exists)companyNames.push(values[x]);
  }
  return companyNames;
}

function clearFilters() {
  var loanOverviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loan overview");
  var fullSheet = loanOverviewSheet.getRange(1, 1, loanOverviewSheet.getMaxRows(), loanOverviewSheet.getMaxRows());
  loanOverviewSheet.unhideRow(fullSheet);

  var lastRow = loanOverviewSheet.getLastRow();
  var total = 0;
  for(var rowNumber = 2; rowNumber < lastRow; rowNumber++)
      total = total + Number(loanOverviewSheet.getRange(rowNumber, 3).getValue());
  Logger.log(total)
  loanOverviewSheet.getRange(lastRow, 3).setValue(total);
  SpreadsheetApp.getActiveSpreadsheet().toast("Completed", "Status", 2);
}
