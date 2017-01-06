function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Generate PDF for first 15 days", functionName: "generatePDFFirst15Days"},
                        {name: "Generate PDF for last 15 days", functionName: "generatePDFLast15Days"},
                        {name: "Delete employee record with zero payment", functionName: "deleteEmployeeWithZeroPayment"}];
  ss.addMenu("Tasks", csvMenuEntries);
}

function deleteEmployeeWithZeroPayment() {
  var ss = SpreadsheetApp.getActive();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(var sheetNumber = 3; sheetNumber < sheets.length; sheetNumber++) {
    var sheet = sheets[sheetNumber];
    var J29Cell = Number(sheet.getRange(29, 10).getValue());
    var J61Cell = Number(sheet.getRange(61, 10).getValue());
    if(J29Cell + J61Cell <= 0) {
      var sheetToDeleted = ss.getSheetByName(sheet.getName());
      ss.deleteSheet(sheetToDeleted);
    }
  }
}

function generatePDFFirst15Days() {
  var documentName = "NYC_" + Utilities.formatDate(new Date(), "GMT+9:00", "yyyy_dd_MMMM") + "_UC_Payslip";
  var doc = DocumentApp.create(documentName);
  var body = doc.getBody();
  var pointsInInch = 72;
  body.setPageHeight(5.83 * pointsInInch);
  body.setPageWidth(8.27 * pointsInInch);
  body.setMarginTop(15);
  body.setMarginBottom(0);
  body.setMarginLeft(0);
  body.setMarginRight(0);

  var tableStyle = {};
  tableStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
  tableStyle[DocumentApp.Attribute.BOLD] = false;
  tableStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  tableStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;

  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BOLD] = true;

  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.FONT_SIZE] = 6;
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  cellStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  cellStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var flag = 0;
  for(var sheetNumber = 3; sheetNumber < sheets.length; sheetNumber++){
    var sheet = sheets[sheetNumber];
    var J29Cell = Number(sheet.getRange(29, 10).getValue());
    if(J29Cell > 0) {
      if(flag == 1)
      {
        body.appendParagraph("");
      }
      flag = 1;
      var table = body.appendTable();
      table.setAttributes(tableStyle);
      table.setBorderColor("#FFFFFF");

      for(var row = 1; row < 32; row++) {
        if(row == 8|| row == 16 || row == 23 || row == 27)
          continue;
        var tr = table.appendTableRow();
        if(row == 7 || row == 15 || row == 22 || row == 26)
          tr.setMinimumHeight(17);
        else
          tr.setMinimumHeight(10);
        for(var col = 1; col < 12; col++) {
          var range = sheet.getRange(row, col);
          var cellValue = range.getDisplayValue();
          var cell = tr.appendTableCell(cellValue);
          if(col == 1)
            cell.setWidth(65);
          else
            cell.setWidth(55);

          cell.setPaddingBottom(0);
          cell.setPaddingTop(0);
          cell.setPaddingLeft(0);
          cell.setPaddingRight(0);
          if(row == 5 && col == 2)
            cell.setAttributes(cellStyle);
          else
            cell.setAttributes(tableStyle);
          cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          if((row == 20 || row == 21 || row == 28 || row == 29) && col == 10) {
            cell.setAttributes(headerStyle);
          }
        }
      }
      body.appendPageBreak();
    }
  }
  var docId = doc.getId();
  doc.saveAndClose();
  var docBlob = doc.getAs('application/pdf');
  docBlob.setName(doc.getName() + ".pdf");
  DriveApp.getFileById(docId).setTrashed(true);
  if (flag == 1) {
    var parentFolder = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
    parentFolder.next().createFile(docBlob);
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('No employee with more than zero payment');
  }
}

function generatePDFLast15Days() {
  var documentName = "NYC_" + Utilities.formatDate(new Date(), "GMT+9:00", "yyyy_dd_MMMM") + "_UC_Payslip-2";
  var doc = DocumentApp.create(documentName);
  var body = doc.getBody();
  var pointsInInch = 72;
  body.setPageHeight(5.83 * pointsInInch);
  body.setPageWidth(8.27 * pointsInInch);
  body.setMarginTop(15);
  body.setMarginBottom(0);
  body.setMarginLeft(0);
  body.setMarginRight(0);

  var tableStyle = {};
  tableStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
  tableStyle[DocumentApp.Attribute.BOLD] = false;
  tableStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  tableStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;

  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.FONT_SIZE] = 6;
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  cellStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  cellStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;


  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BOLD] = true;

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var flag = 0;
  for(var sheetNumber = 3; sheetNumber < sheets.length; sheetNumber++){
    var sheet = sheets[sheetNumber];
    var J61Cell = Number(sheet.getRange(61, 10).getValue());
    if(J61Cell > 0) {
      if(flag == 1)
      {
        body.appendParagraph("");
      }
      flag = 1;
      var table = body.appendTable();
      table.setAttributes(tableStyle);
      table.setBorderColor("#FFFFFF");

      for(var row = 33; row < 64; row++) {
        if(row == 40|| row == 48 || row == 55 || row == 59)
          continue;
        var tr = table.appendTableRow();
        if(row == 39 || row == 47 || row == 54 || row == 58)
          tr.setMinimumHeight(17);
        else
          tr.setMinimumHeight(10);

        for(var col = 1; col < 12; col++) {
          var range = sheet.getRange(row, col);
          var cellValue = range.getDisplayValue();
          var cell = tr.appendTableCell(cellValue);
          if(col == 1)
            cell.setWidth(65);
          else
            cell.setWidth(55);
          cell.setPaddingBottom(0);
          cell.setPaddingTop(0);
          cell.setPaddingLeft(0);
          cell.setPaddingRight(0);
          if(row == 37 && col == 2)
            cell.setAttributes(cellStyle);
          else
            cell.setAttributes(tableStyle);
          cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          if((row == 52 || row == 53 || row == 60 || row == 61) && col == 10) {
            cell.setAttributes(headerStyle);
          }
        }
      }
      body.appendPageBreak();
    }
  }
  var docId = doc.getId();
  doc.saveAndClose();
  var docBlob = doc.getAs('application/pdf');
  docBlob.setName(doc.getName() + ".pdf");
  DriveApp.getFileById(docId).setTrashed(true);
  if (flag == 1) {
    var parentFolder = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
    parentFolder.next().createFile(docBlob)
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('No employee with more than zero payment');
  }
}
