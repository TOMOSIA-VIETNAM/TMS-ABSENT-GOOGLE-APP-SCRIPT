// Create new sheets for person
function newSheets(sheetname) {
  //sheetname = 'test1'
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //var data = activeSpreadsheet.getDataRange().getValues();
  var newSheets = activeSpreadsheet.getSheetByName(sheetname);

  if (newSheets != null) {
    activeSpreadsheet.deleteSheet(newSheets);
  }

  newSheets = activeSpreadsheet.insertSheet();
  newSheets.setName(sheetname);
}

//Insert data into sheet of person
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

function dataInit(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  data = sheet.getDataRange().getValues();
  sheetname = 'DSNhanvien';
  var sheetDsnhanvien = sheet.getSheetByName(sheetname);
  sheetDsnhanvien.activate();
  data1 = sheet.getDataRange().getValues();
  return data, data1;
}

//Send message report summary absent to Slack for person 
var SLACK_URL = 'https://hooks.slack.com/services/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';

function sendSlackMessageToPerson() {
  dataInit();
  var tmp = [];
  var today = new Date();
  for (var i = 1; i < data1.length; i++){
    for (var j=1; j < data.length; j++){
      var stringMonth = data[j][4].slice(-7,-5);
      var stringYear = data[j][4].slice(-4);
      if (data1[i][2] == data[j][2] && (today.getMonth()+1) == stringMonth && today.getFullYear() == stringYear){
        tmp.push([data[0][3] + ": "
        + data[j][3] + "\t" + data[0][4] + ": " + data[j][4] + "\t" +  data[0][5] + ": " + data[j][5] + '\n']);
      }
    }
    if(tmp.length === 0){
      var slackMessage = {
      channel: '@'+ slicestring,
      username: data1[i][1],
      text: 'Absent summary month '+ (today.getMonth()+1) + '\n' + 'No day off' 
      };
      var options = {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(slackMessage)
      };
      UrlFetchApp.fetch(SLACK_URL, options);
      tmp = [];
    }
    else{
      var position = data1[i][2].indexOf('@');
      var slicestring = data1[i][2].slice(0,position);
      var slackMessage = {
      channel: '@'+ slicestring,
      username: data1[i][1],
      text: 'Absent summary month '+ (today.getMonth()+1) + '\n' + tmp 
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
}


function test(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getDataRange().getValues();
  sheetname = 'DSNhanvien';
  var sheetDsnhanvien = sheet.getSheetByName(sheetname);
  sheetDsnhanvien.activate();
  var data1 = sheet.getDataRange().getValues();
}

//Send message report summary absent to Slack for HR 
function sendSlackMessageToHR() {
  dataInit();
  var tmp = [];
  var today = new Date();
  for (var i = 1; i < data1.length; i++){
    for (var j=1; j < data.length; j++){
      var slicesstring = data[j][4].slice(-7,-5);
      var stringYear = data[j][4].slice(-4);
      if (data1[i][2] == data[j][2] && (today.getMonth()+1) == slicesstring && today.getFullYear() == stringYear){
        tmp.push([data[0][3] + ": "
        + data[j][3] + "\t" + data[0][4] + ": " + data[j][4] + "\t" +  data[0][5] + ": " + data[j][5] + '\n']);
      }
    }
    if(tmp.length === 0){
      var slackMessage = {
      channel: '#testabsentreport',
      username: data1[i][1],
      text: 'Absent summary month '+ (today.getMonth()+1) + '\n' + 'No day off' 
      };
      var options = {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(slackMessage)
      };
      UrlFetchApp.fetch(SLACK_URL, options);
      tmp = [];
    }
    else{
      var slackMessage = {
      channel: '#testabsentreport',
      username: data1[i][1],
      text: 'Absent summary month '+ (today.getMonth()+1) + '\n' + tmp 
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
}

function callSheetConfig(){
  var sheet1 = SpreadsheetApp.openById('1jl8YgBIwHILgdezpvVSjspI_tJZ7RZm682W_ALx7j5w');
  datainit = sheet1.getDataRange().getValues();
  var sheetname = 'Template';
  testsheet = sheet1.getSheetByName(sheetname);
  testsheet.activate();
  dataconfig = testsheet.getDataRange().getValues();
}

//Send email report summary absent to Email for person
function sendEmailToPerson() {
  dataInit();
  callSheetConfig();
  var tmp = []
  var today = new Date();
  for (var i = 1; i < data1.length; i++){
    for (var j=1; j < data.length; j++){
      var slicesstring = data[j][4].slice(-7,-5);
      var stringYear = data[j][4].slice(-4);
      if(data1[i][2] == data[j][2] && (today.getMonth()+1) == slicesstring && today.getFullYear() == stringYear){
        tmp.push([data[0][3] + ": "
        + data[j][3] + "\t" + data[0][4] + ": " + data[j][4] + "\t" +  data[0][5] + ": " + data[j][5] + '\n']);   
      }
    }
    if (tmp.length === 0){
      var emailAddress = data1[i][2]
      var message = dataconfig[1][1] + '\n' + 'No day off'
      var subject = dataconfig[1][0];
      MailApp.sendEmail(emailAddress, subject, message)
      tmp =[];
    }
    else{
      var emailAddress = data1[i][2]
      var message = dataconfig[1][1] + '\n' + tmp
      var subject = dataconfig[1][0];
      MailApp.sendEmail(emailAddress, subject, message)
      tmp =[];
    }
  }
}

//Send email to HR
function sendEmailToHR(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  callSheetConfig();
  var emailAddress = datainit[1][2];
  var message = dataconfig[1][1] + "\n" + sheet.getUrl() 
  var subject = dataconfig[1][0];
  MailApp.sendEmail(emailAddress, subject, message);
}
//Trigger test
/*function createNewSendSlackMessagetrigger(){
  ScriptApp.newTrigger('sendSlackMessageToPerson')
      .timeBased()
      .everyMinutes(1)
      .create();
}*/

// Create trigger for send Slack message to Person
/*function createNewSendSlackMessagetrigger(){
  ScriptApp.newTrigger('sendSlackMessageToPerson')
      .timeBased()
      .onMonthDay(1)
      .atHour(9)
      .create();
}*/

// Create trigger for send Email to Person
/*function createNewSendEmailrigger(){
  ScriptApp.newTrigger('sendEmailToPerson')
      .timeBased()
      .onMonthDay(1)
      .atHour(9)
      .create();
}*/

// Create trigger for send Slack message to HR
/*function createNewSendSlackMessagetrigger(){
  ScriptApp.newTrigger('sendSlackMessageToHR')
      .timeBased()
      .onMonthDay(1)
      .atHour(9)
      .create();
}*/

// Create trigger for send Email to HR
/*function createNewSendEmailrigger(){
  ScriptApp.newTrigger('sendEmailToHR')
      .timeBased()
      .onMonthDay(1)
      .atHour(9)
      .create();
}*/

// Create trigger to Insert data report
/*function createInsertDataReport(){
  ScriptApp.newTrigger('insertDataReport')
      .timeBased()
      .onMonthDay(1)
      .atHour(7)
      .create();
}*/

//Create trigger for edit Sheet
function createTriggerForChangeSheet(){
  var sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('insertToSheet')
    .forSpreadsheet(sheet)
    .onEdit()
    .create();
}

function insertDataReport(){
  dataInit();
  var today = new Date();
  var sheetReport = 'Report Absent tháng ' + (today.getMonth()+1);
  newSheets(sheetReport);
  SpreadsheetApp.getActiveSheet().appendRow(['Email', 'Số ngày nghỉ']);
  for (var i = 1; i < data1.length; i++){
    var number = 0;
    for(var j = 1 ; j <data.length; j++){
      if(data1[i][2] == data[j][2]){
        if(data[j][5] != 'Several days'){
          switch (data[j][5]){
            case 'Morning':
              number += 0.5;
              break;
            case 'Afternoon':
              number += 0.5;
              break;
            case 'Cả ngày':
              number += 1;
              break;
          }
        }
        else number += parseInt(data[j][4].slice(6,8)) - parseInt(data[j][3].slice(6,8));
      }
    }
    SpreadsheetApp.getActiveSpreadsheet().appendRow([data1[i][2], number]);
  }
}

function reportChart(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  insertDataReport();
  var today = new Date();
  var sheetReport = 'Report Absent tháng ' + (today.getMonth()+1);
  var sheetReport = ss.getSheetByName(sheetReport);
  sheetReport.activate();
  var data = ss.getDataRange().getValues();
  var range = sheetReport.getRange(2,1,(data.length-1),2);
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(range)
    .setPosition(2, 8, 0, 0)
    .build();

  sheetReport.insertChart(chart);
}
