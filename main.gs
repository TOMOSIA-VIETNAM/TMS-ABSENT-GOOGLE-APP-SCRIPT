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

function insertToSheet() {
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
  }
}

// Send message to Slack

var SLACK_URL = 'https://hooks.slack.com/services/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';

function sendSlackMessage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getDataRange().getValues();
  sheetname = 'DSNhanvien';
  var sheetDsnhanvien = sheet.getSheetByName(sheetname);
  sheetDsnhanvien.activate();
  var data1 = sheet.getDataRange().getValues();
  var tmp = [];
  for (var i = 1; i < data1.length; i++){
    for (var j=1; j < data.length; j++){
      if (data1[i][2] == data[j][2]){
        tmp.push([data[0][3] + ": "
        + data[j][3] + "\t" + data[0][4] + ": " + data[j][4] + "\t" +  data[0][5] + ": " + data[j][5] + '\n']);
      }
    }
    var position = data1[i][2].indexOf('@');
    var slicestring = data1[i][2].slice(0,position);
    var slackMessage = {
    channel: '@'+ slicestring,
    username: data1[i][1],
    text: 'Absent\n' + tmp 
    };
    var options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(slackMessage)
    };
    UrlFetchApp.fetch(SLACK_URL, options);
    tmp = [];
  }
}

/*function createNewSendSlackMessagetrigger(){
  ScriptApp.newTrigger('sendSlackMessage')
      .timeBased()
      .everyMinutes(1)
      .create();
}*/

function duplicate(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getDataRange().getValues();
  var tmp = [];
  for(var j=2; j < data.length; j++){
    tmp.push(data[j][2])
  }
  Logger.log([...new Set(tmp)])
}

function newChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.COLUMN)
     .addRange(sheet.getRange('A2:A6'))
     .setPosition(5, 10, 0, 0)
     .build();

  sheet.insertChart(chart);
}
