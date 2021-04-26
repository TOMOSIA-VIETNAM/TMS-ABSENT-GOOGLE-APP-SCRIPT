/*function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 1; j <= data.length+1; j++){
      Logger.log(data[i][j]);
    }
  }
}*/

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
  for (var i = 1; i < data.length; i++) {
    sheetname = data[i][2];
    newSheets(sheetname);
    SpreadsheetApp.getActiveSheet().appendRow(data[0]);
  }

  for (var i = 1; i < data.length; i++) {
    sheetname = data[i][2];
    var testSheets = sheet.getSheetByName(sheetname);
    testSheets.activate();
    SpreadsheetApp.getActiveSheet().appendRow(data[i]);
  }
}

// Send message to Slack

var SLACK_URL = 'https://hooks.slack.com/services/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';

function sendSlackMessage() {

  var slackMessage = {
    channel: '@tt.bao.ngo',
    username: "Absent report",
    text: 'absent report'
  };

  var options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(slackMessage)
  };
  UrlFetchApp.fetch(SLACK_URL, options);
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
