var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function go() {
  var sheets = spreadsheet.getSheets();
  var transactionsSheet = spreadsheet.getSheetByName('Transactions');
  var transactionsRange = transactionsSheet.getRange(2, 2, transactionsSheet.getLastRow(), transactionsSheet.getLastColumn());
  var transactionValues = transactionsRange.getValues();

  var clients = {};

  for ( var i = 1; i < transactionValues.length; i++) {
    var transactionRow = transactionValues[i];
    if (transactionRow[0] == '') {
      continue;
    }

    if (!clients[transactionRow[0]]) {
      clients[transactionRow[0]] = [];
    }
    clients[transactionRow[0]].push(transactionRow);
  }

 var codesSheets = getCodesSheets(); //Array exists tabs. Access by their codes.
 var companyNames = getCompanyNames(); //Array names of tabs. From "Settings" tab. Access by their codes.

 for(var key in clients) {
   Logger.log("Key Value ---" + parseInt(key))
   Logger.log(codesSheets)
   if (codesSheets[parseInt(key)] != null && codesSheets[parseInt(key)] == companyNames[key]) {
     var sheetName = codesSheets[Number(key)];
     var sheet = spreadsheet.getSheetByName(sheetName);
   } else {
     var sheetName = companyNames[key];
     sheetName = sheetName?sheetName:key;
     Logger.log('key = ' + key + '; sheetName = ' + sheetName);
     var sheet = spreadsheet.getSheetByName(sheetName);;
     if(sheet != null)
       spreadsheet.deleteSheet(sheet);
     sheet = spreadsheet.insertSheet(sheetName);
     sheet.deleteColumns(8, 19);                    //Updated code
   }
   fillSheet(sheet, clients[key]);
   //var lr = FindLastRowofSpecificColumn(sheet,0)+2; //Updated code
   //sheet.deleteRows(lr,(1000-lr+1));                //Updated code
 }
}


function getCompanyNames() {
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  var settingsRange = settingsSheet.getDataRange();
  var settingsValues = settingsRange.getValues();
  var companyNames = {};

  for (var i = 1; i < settingsValues.length; i++) {
    companyNames[settingsValues[i][1]] = settingsValues[i][0];
  }
  Logger.log(companyNames);
  return companyNames;
}


function fillSheet(sheet, values) {
  sheet.clearContents();
  var rowCount = values.length;
  if (rowCount == 0) {
    exit;
  }
  var colCount = values[0].length;
  Logger.log("Filling Sheet" + sheet.getName())
  var range = sheet.getRange(1, 1, rowCount, colCount);
  Logger.log('\n\rowCount = ' + rowCount + '; colCount = ' + colCount);
  range.setValues(values);
}


function getCodesSheets() {
  var sheets = spreadsheet.getSheets();
  var codesSheets = {};
  for (var i = 0; i < sheets.length; i++){
    var range = sheets[i].getRange(1, 1);
    var value = range.getValue();
    Logger.log(parseInt(value))
    if (value != '') {
      codesSheets[parseInt(value)] = sheets[i].getName();
    }
  }
  return codesSheets;
}


function delTabs() {
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if ((sheetName == 'Client 1') || (sheetName == 'Client 2') || (sheetName == 'Transactions') || (sheetName == 'Settings')) {
      continue;
    } else {
      spreadsheet.deleteSheet(sheets[i]);
    }
  }
}


function startTimer() {
  var allTriggers = ScriptApp.getProjectTriggers();
  if (allTriggers.length < 1) {
    ScriptApp.newTrigger('go').timeBased().everyMinutes(10).create();
  }
    spreadsheet.toast('The timer is running.', '' , 10);
}


function stopTimer() {
  var allTriggers = ScriptApp.getProjectTriggers();
  if (allTriggers.length > 0) {
    ScriptApp.deleteTrigger(allTriggers[0]);
  }
    spreadsheet.toast('The timer is stopped.', '' , 10);
}


function onOpen() {
  var menuEntries = [];
  menuEntries.push({name: "Go", functionName: 'go'},
                   {name: "Del Tabs", functionName: 'delTabs'},
                   {name: "Start Timer", functionName: "startTimer"},
                   {name: "Stop Timer", functionName: "stopTimer"}
  );
  spreadsheet.addMenu("Script", menuEntries);
}


function parseOb(ob) {
  Logger.log('\n\n     Start function parseOb()');
  var i = 0;
  for(var key in ob) {
    Logger.log('i = ' + i + ', ' + key + ' = ' + ob[key]);
    i++;
  }
}


function getSettings() {
  var id ='1DD85jAFf0syOBw_jm7UzGUaIaLfIBS5Ve9MhNfrLAJs';
  var sourceSpreadsheet = SpreadsheetApp.openById(id);
  var sourceSheets = sourceSpreadsheet.getSheets();

  var settingsValues = [];
  var j = 0;
  for (var i = 0; i < sourceSheets.length; i++){
    var sourceSheetName = sourceSheets[i].getName()
    if ((sourceSheetName == 'Cards') || (sourceSheetName == 'Transactions')) {
      continue;
    }
    var sourceRange = sourceSheets[i].getRange(1, 5);
    var sourceValue = sourceRange.getValue();

    settingsValues[j] = [sourceSheetName, sourceValue];
    j++;

  }
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  var rowCount = settingsValues.length;
  var colCount = settingsValues[0].length;
  var settingsRange = settingsSheet.getRange(2, 1, rowCount, colCount);
  settingsRange.setValues(settingsValues);
//  Logger.log(settingsValues)
}


function FindLastRowofSpecificColumn(sheet,col){   //A=0,B=1
    var data = sheet.getDataRange().getValues();
    for(var i = data.length-1 ; i >=0 ; i--){
    if (data[i][col] != null && data[i][col] != ''){
        return i+1 ;
    }
  }
}
