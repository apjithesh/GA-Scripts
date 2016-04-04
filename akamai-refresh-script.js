

function sendEmails(emailAddress, message, subject) {
  /* var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 2;   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 2)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0];  // First column
    var message = row[1];       // Second column
    var subject = "Sending emails from a Spreadsheet";
    
  } */
  
  MailApp.sendEmail(emailAddress, subject, message);
}

function URLToRefresh () {
  /* Add Moment library by going to Resources -> Libraries in script editor and using this key - MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48 */
  var moment = Moment.load();
  var fileID = "1rF5SuyQPGssrcbTjcAmtnHO8659Mh-2JqYcmJJjiN00";
  var DateCellinfo = "N1";
  var URLDataRangeinfoStart = "B";
  var URLDataRangeinfoend = "B";
  var emailDataRangeinfoStart = "C";
  var emailDataRangeinfoend = "C";
  var SheetwithData = "Refresh-data-sheet";
  var emailLastemptyrow = "N4";
  var URLlastemptyrow = "N3";
  var akausernamecell = "N5";
  var akaPasscell = "N6";
  var akapurgeURLCell = "N7";
  var resultSheet = "Execution Results";
  //Get the sheet and update the cell to current date
  var getResultSheet = SpreadsheetApp.openById(fileID);// getActiveSpreadsheet();
  var allSheets = getResultSheet.getSheets();
  var responseResultSheet = getResultSheet.getSheetByName(resultSheet);
  if (!responseResultSheet) {
    responseResultSheet = getResultSheet.insertSheet(resultSheet, allSheets.length + 1);
    responseResultSheet.getRange(1,1).setValue("Execution Time");
    responseResultSheet.getRange(1,2).setValue("URLs for Refresh");
    responseResultSheet.getRange(1,3).setValue("Requestors");
    responseResultSheet.getRange(1,4).setValue("Akamai Response");
    responseResultSheet.getRange(1,5).setValue("Send Email Response");
  }
  
  var sheet = SpreadsheetApp.openById(fileID); //openById(fileId)
  sheet = sheet.getSheetByName(SheetwithData);
  var nowDatecell = sheet.getRange(DateCellinfo);
  nowDatecell.setValue(moment().format('YYYY-MM-DD H:mm'));
  var dataURLRange = sheet.getRange(URLDataRangeinfoStart + ":" + URLDataRangeinfoend);
  
  Utilities.sleep(10 * 1000)
  //setTimeout(function(){ dataRange.getValues(); }, 10000);
  var URLs = dataURLRange.getValues();
  /*var ct =0;
   while ( URLs[ct][0] != "" ) {
    ct++;
  } */
  dataURLRange = sheet.getRange(URLDataRangeinfoStart + "1:" + URLDataRangeinfoend + sheet.getRange(URLlastemptyrow).getValue());
  URLs = dataURLRange.getValues();
  //Browser.msgBox(URLs.toString());
  var URLForPayload = "";
  var URLSList = new Array();
      URLSList = URLs.toString().split(",");
  //Browser.msgBox("URLslist is " + URLSList + " and lengths is " + URLSList.length + " URLs is " + URLs + " and length is " + URLs.length);
    for (i=0; i< URLSList.length; i++){
    if (URLForPayload.length > 0) {
      if (URLSList[i] != "" && i != (URLSList.length - 1)) { URLForPayload = URLForPayload + ',' + URLSList[i].trim(); }
    } else {
       // Browser.msgBox(URLSList[i]);

      if (URLSList[i] != "") { URLForPayload = URLSList[i].trim(); }
    }
    
    //URLObject = URLObject + ',\n"' + URLS[i] + '"';
  }
  //Browser.msgBox(URLs);
  var dataEmailRange = sheet.getRange(emailDataRangeinfoStart + ":" + emailDataRangeinfoend);
  var emails = dataEmailRange.getValues();
  /* var ctEmail =0;
   while ( emails[ctEmail][0] != "" ) {
    ctEmail++;
  } */
  dataEmailRange = sheet.getRange(emailDataRangeinfoStart + "1:" + emailDataRangeinfoend + sheet.getRange(emailLastemptyrow).getValue());
  emails = dataEmailRange.getValues();
  //sendEmails("jithesh@sumologic.com",URLs.replace(/,/g, '\n'), "URLS for Akamai Refresh");
  var message = "Hello \n\nRequest for your cache refresh has been kicked off, please check the web pages after 10 minutes, they should show refreshed content. \n\nIf issue persists, please reach out to jithesh@sumologic.com for help. \n\nURL's being refreshed are:- \n" + URLForPayload + "\n\nRegards\nWeb Team@Sumo Logic";
  var akaUsername = sheet.getRange(akausernamecell).getValue();
  var akaPass = sheet.getRange(akaPasscell).getValue();
  var akaPurgeURL = sheet.getRange(akapurgeURLCell).getValue();
  var payloadInfo = '{ "action" : "remove" , "domain" : "production", "objects": ['+ URLForPayload + ']}';
  //sheet.getRange("N8").setValue(payloadInfo);
  //payloadInfo = JSON.parse(payloadInfo);
  //payloadInfo = sheet.getRange("N8").getValue();
  //sheet.getRange("N15").setValue(payloadInfo);
  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(akaUsername + ':' + akaPass ),
      "Content-Type" : "application/json"
    },
    "payload":  payloadInfo
  };
  //sheet.getRange("N9").setValue(URLs);
  var lastRow = responseResultSheet.getLastRow();
  lastRow = lastRow + 1;
  if (URLForPayload.length > 0) {
 
    try {
      if (URLForPayload.length > 0) {
        var response = UrlFetchApp.fetch(akaPurgeURL, options);
        responseResultSheet.getRange(lastRow, 4).setValue(response);
      } else {
        responseResultSheet.getRange(lastRow, 4).setValue("No URLs to refresh");
      }
      responseResultSheet.getRange(lastRow, 1).setValue(moment().format('YYYY-MM-DD H:mm'));
      if (URLForPayload.indexOf(",") >0 ){ //this if is not neeeded, but will keep it in there for now
        //responseResultSheet.getRange(lastRow, 2).setValue(URLForPayload.replace(/,/g, ',\n')); 
        responseResultSheet.getRange(lastRow, 2).setValue(URLForPayload); 
      } else {
        responseResultSheet.getRange(lastRow, 2).setValue(URLForPayload); 
      }
      
      
      
      //sheet.getRange("N10").setValue(response);
    } catch(e){
      responseResultSheet.getRange(lastRow, 1).setValue(moment().format('YYYY-MM-DD H:mm'));
      if (URLForPayload.indexOf(",") >0 ){ //this if is not neeeded, but will keep it in there for now
        //responseResultSheet.getRange(lastRow, 2).setValue(URLForPayload.replace(/,/g, ',\n')); 
        responseResultSheet.getRange(lastRow, 2).setValue(URLForPayload); 
      } else {
        responseResultSheet.getRange(lastRow, 2).setValue(URLForPayload); 
      }
      responseResultSheet.getRange(lastRow, 4).setValue(e);
      //sheet.getRange("N10").setValue(e);
    }
    try {
      MailApp.sendEmail(emails, "Your Cache refresh request kicked off", message);
      if (emails.indexOf(",") >0){
        responseResultSheet.getRange(lastRow, 3).setValue(emails.replace(/,/g, ',\n'));
      } else {
        responseResultSheet.getRange(lastRow, 3).setValue(emails);
      }
      responseResultSheet.getRange(lastRow, 5).setValue("Emails Sent");
      //sheet.getRange("o10").setValue(response);
    }catch(e){
      if (emails.indexOf(",") >0){
        responseResultSheet.getRange(lastRow, 3).setValue(emails.replace(/,/g, ',\n'));
      } else {
        responseResultSheet.getRange(lastRow, 3).setValue("");
      }
      responseResultSheet.getRange(lastRow, 5).setValue(e);
      //sheet.getRange("o10").setValue(e);
    }
  }
  
}

function parseURL(url) {
  var URLS = new Array();
  url = url.replace(/\n/g, ",");
  URLS = url.split(",");
  var URLObject = "";
  //'"' + URLS[0] + '"';
  var regexpression = "^https://www.sumologic.com.*";
 var URLRegex = new RegExp(regexpression);
  //if (URLRegex.test(URLS[0])) { URLObject = '"' + URLS[0] + '"'; }
  for (i=0; i< URLS.length; i++){
    if (URLObject.length > 0) {
      if (URLRegex.test(URLS[i])) { URLObject = URLObject + ',\n"' + URLS[i] + '"'; }
    } else {
      if (URLRegex.test(URLS[i])) { URLObject = '"' + URLS[i] + '"'; }
    }
    
    //URLObject = URLObject + ',\n"' + URLS[i] + '"';
  }
  return URLObject;
}

function getLastCell(){
  var SheetwithData = "Sheet2";
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = sheet.getSheetByName(SheetwithData);
  var lastRow = sheet.getLastRow();
   var lastColumn = sheet.getLastColumn();
  
  Browser.msgBox("Last Row is " + lastRow + "Last COlumn is " + lastColumn);
  
}