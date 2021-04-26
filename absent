function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 1; j <= data.length+1; j++){
      Logger.log(data[i][j]);
    }
  }
}

function newSheets(sheetname) {
  //sheetname = 'test1'
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //var data = activeSpreadsheet.getDataRange().getValues();
  var testSheets = activeSpreadsheet.getSheetByName(sheetname);

  if (testSheets != null) {
    activeSpreadsheet.deleteSheet(testSheets);
  }

  testSheets = activeSpreadsheet.insertSheet();
  testSheets.setName(sheetname);
  /*testSheets = activeSpreadsheet.insertSheet();
  testSheets.setName(sheetname);*/
  
  /*var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(['test', '111111']);*/
}

function insertosheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    sheetname = data[i][2];
    newSheets(sheetname);
    SpreadsheetApp.getActiveSheet().appendRow(data[0]);
  }

  for (var i = 1; i < data.length; i++) {
    sheetname = data[i][2];
    var testSheets = sheet.getSheetByName(sheetname);
    testSheets.activate();

    SpreadsheetApp.getActiveSheet().appendRow(data[i]);
    Logger.log(testSheets.getUrl());
  }
  
}
function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 2; // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 3);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i in data) {
    var row = data[i];
    var emailAddress = row[2]; // First column
    var message = row[1]; // Second column
    var subject = 'Sending emails from a Spreadsheet';
    MailApp.sendEmail(emailAddress, subject, message);
  }
}
//var ss = SpreadsheetApp.getActiveSpreadsheet();
//Logger.log(ss.getUrl());
//var ssNew = SpreadsheetApp.create("Finances", 50, 5);
//Logger.log(ssNew.getUrl());

/*function newChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.BAR)
     .addRange(sheet.getRange('A2:A6'))
     .setPosition(5, 5, 0, 0)
     .build();

  sheet.insertChart(chart);
} */
