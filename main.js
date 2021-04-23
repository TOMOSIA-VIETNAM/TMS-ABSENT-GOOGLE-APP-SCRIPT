function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 1; j <= data.length + 1; j++) {
      Logger.log(data[i][j]);
    }
  }
}

function newSheets(sheetname) {
  /*sheetname = 'test1'*/
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = activeSpreadsheet.getDataRange().getValues();
  var testSheets = activeSpreadsheet.getSheetByName(sheetname);

  if (testSheets != null) {
    activeSpreadsheet.deleteSheet(testSheets);
  }

  testSheets = activeSpreadsheet.insertSheet();
  testSheets.setName(sheetname);
  SpreadsheetApp.getActiveSheet().appendRow(data[1]);
  /*var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(['test', '111111']);*/
}

function insertosheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] == "ho2@gmail.com") {
      sheetname = "ho2"
      newSheets(sheetname);
      SpreadsheetApp.getActiveSheet().appendRow(data[i]);
    }
  }
}

/*function newChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.BAR)
     .addRange(sheet.getRange('A2:A6'))
     .setPosition(5, 5, 0, 0)
     .build();

  sheet.insertChart(chart);
}*/
