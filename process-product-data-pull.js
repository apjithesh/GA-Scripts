/* function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var searchMenuEntries = [ {name: "Search in all files", functionName: "search"}];
  ss.addMenu("Search Google Drive", searchMenuEntries);
}
  */
function search() { //This is a function where I tried and tested many google sheet functions
  // Prompt the user for a search term
  //var searchTerm = Browser.inputBox("Enter the string to search for:");
 var searchTerm = "Copy of Import of";
  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SSID=ss.getId();
  var sheet = ss.getActiveSheet();
    var config_sheet = ss.getSheetByName("config-sheet");
  var regions = new Array();
  regions = config_sheet.getRange('B11').getValue().split(",");
 Browser.msgBox(regions[1]);
  // Set up the spreadsheet to display the results
  var headers = [["File Name", "File Type", "URL"]];
  sheet.clear();
  var anotherstr = "C1";
  sheet.getRange("A1:" + anotherstr).setValues(headers);
 
  // Search the files in the user's Google Drive for the search term
  // See documentation for search parameters you can use
  // https://developers.google.com/apps-script/reference/drive/drive-app#searchFiles(String)
  //var files = DriveApp.searchFiles("title contains '"+searchTerm.replace("'","\'")+"'"); https://drive.google.com/open?id=0B-jvR7IWXrOqVlQxTUlfVUY1RGs
   //var files = DriveApp.getFolderById("0B-jvR7IWXrOqVlQxTUlfVUY1RGs").getFiles();
  //var files = DriveApp.getParents();
  //var files = DriveApp.getFolderById("0B-jvR7IWXrOqVlQxTUlfVUY1RGs").getFilesByName("Copy of Import of");
  var files = DriveApp.getFolderById("0B-jvR7IWXrOqVlQxTUlfVUY1RGs").searchFiles("title contains '"+searchTerm.replace("'","\'")+"'");
  // create an array to store our data to be written to the sheet
  var output = [];
  // Loop through the results and get the file name, file type, and URL
  while (files.hasNext()) {
    var file = files.next();
     
    var name = file.getName();
    var type = file.getMimeType();
    var url = file.getUrl();
    var fileid = file.getId();
    // push the file details to our output array (essentially pushing a row of data)
    output.push([name, type, url, fileid]);
  }
  // write data to the sheet
  
 // var formula = '=importrange("' + fileid + '","signups-us2!A:ZZ")';
  sheet.getRange(2, 1, output.length, 4).setValues(output);
 // sheet.getRange(5, 1).setValue(formula);
  
  //var source_sheet = SpreadsheetApp.openById(fileid);
  //var target = SpreadsheetApp.getActiveSpreadsheet();
//var target_sheet = target.getSheetByName("Sheet1");
  //source_sheet = source_sheet.getSheetByName("signups-us2");
 // var source_range = source_sheet.getDataRange(); //getRange("A:ZZ");
 // var A1Range = source_range.getA1Notation();
    //get the data values in range
 // var SData = source_range.getValues();
///var target_range = target_sheet.getRange("A:ZZ");
 // source_range.copyTo(target_range);
  
  //target_sheet.getRange(5,1,SData.length, SData.width).setValues(SData);
  
  //target_sheet.getRange(A1Range).setValues(SData);
  
}

function getConfigs () {
//Find the folder where the spreadsheets are stored with data
  var config = [
  ];
// Get config spreadsheet
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config_sheet = ss.getSheetByName("config-sheet");
  var regions = new Array();
  regions = config_sheet.getRange('B2').getValue().split(",");
  config['folder-id-for-spreadsheet-data'] =  config_sheet.getRange('B1').getValue();
  //Browser.msgBox("folder id received is " + config['folder-id-for-spreadsheet-data']);
  
  // Loop thru and create the config list
  var rowNum = 3;
  for (i=0; i < regions.length; i++) {
    //config.push(['fileName-'+regions[i],config_sheet.getRange('B'+rowNum).getValue()]);
    config['fileName-'+regions[i]] = config_sheet.getRange('B'+rowNum).getValue();
    //Browser.msgBox("config " + config['fileName-'+regions[i]] + " regions is " + regions[i] + " Value should be " + config_sheet.getRange('B'+rowNum).getValue());
    rowNum++;
    config['sourceSheetName-'+regions[i]] = config_sheet.getRange('B'+rowNum).getValue();
    //Browser.msgBox("config " + config['sourceSheetName-'+regions[i]] + " regions is " + regions[i] + " Value should be " + config_sheet.getRange('B'+(rowNum)).getValue());
    rowNum++;
    config['targetSheetName-'+regions[i]] = config_sheet.getRange('B'+rowNum).getValue();
   // Browser.msgBox("config " + config['targetSheetName-'+regions[i]] + " regions is " + regions[i] + " Value should be" + config_sheet.getRange('B'+(rowNum)).getValue());
    //Browser.msgBox("rownumber " + rowNum)
    rowNum++;
  } 
  config['regions'] = config_sheet.getRange('B2').getValue();
  
  /* config['fileName-US2'] = config_sheet.getRange('B2').getValue();
  config['fileName-SYD'] = config_sheet.getRange('B3').getValue();
  config['fileName-DUB'] = config_sheet.getRange('B4').getValue();
  config['sourceSheetName-US2'] = config_sheet.getRange('B5').getValue();
  config['sourceSheetName-SYD'] = config_sheet.getRange('B6').getValue();
  config['sourceSheetName-DUB'] = config_sheet.getRange('B7').getValue();
  config['targetSheetName-US2'] = config_sheet.getRange('B8').getValue();
  config['targetSheetName-SYD'] = config_sheet.getRange('B9').getValue();
  config['targetSheetName-DUB'] = config_sheet.getRange('B10').getValue(); */
  //Browser.msgBox(config['folder-id-for-spreadsheet-data'] + config['regions'] + " overall configs " + config );
  return config;
}

function getData () {
// Gets data from other files that has the product data and populate into this sheet
  var config = getConfigs();
  
  // Loop thru and process the files and copy data into this spreadsheet
  var regions = config['regions'].split(",");
  //Browser.msgBox(config['folder-id-for-spreadsheet-data'] + config['regions'] + " overall configs " + config );
  for (i=0; i < regions.length ; i++) {
    var files = DriveApp.getFolderById(config['folder-id-for-spreadsheet-data']).searchFiles("title contains '"+config['fileName-' + regions[i]].replace("'","\'")+"'");
    
      // Loop thru the files to get file id for each
    while (files.hasNext()) {
      var file = files.next();
      
      var name = file.getName();
      var type = file.getMimeType();
      var url = file.getUrl();
      var fileId = file.getId();
    }
        var source_sheet = SpreadsheetApp.openById(fileId);
    source_sheet = source_sheet.getSheetByName(config['sourceSheetName-' + regions[i]]);
    var source_range = source_sheet.getDataRange(); //getRange("A:ZZ");
    var A1Range = source_range.getA1Notation();
    //get the data values in range
    var SData = source_range.getValues();
    var target = SpreadsheetApp.getActiveSpreadsheet();
    var target_sheets = target.getSheets();
    //Browser.msgBox("Sheets are # " + target_sheets);
    var target_sheet = target.getSheetByName(config['targetSheetName-' + regions[i]]);
      if (!target_sheet) {
        target_sheet = target.insertSheet(config['targetSheetName-' + regions[i]], target_sheets.length + 1);
      }
    target_sheet.getRange(A1Range).setValues(SData);
    
    //source_sheet = source_sheet.getSheetByName("signups-us2");
    // var source_range = source_sheet.getDataRange(); //getRange("A:ZZ");
    // var A1Range = source_range.getA1Notation();
    //get the data values in range
    // var SData = source_range.getValues();
    ///var target_range = target_sheet.getRange("A:ZZ");
    // source_range.copyTo(target_range);
    
    //target_sheet.getRange(5,1,SData.length, SData.width).setValues(SData);
    
    //target_sheet.getRange(A1Range).setValues(SData);
    
  }
  
  /* var filesUS2 = DriveApp.getFolderById(config['folder-id-for-spreadsheet-data']).searchFiles("title contains '"+config['fileName-US2'].replace("'","\'")+"'");
  var filesSYD = DriveApp.getFolderById(config['folder-id-for-spreadsheet-data']).searchFiles("title contains '"+config['fileName-SYD'].replace("'","\'")+"'");
  var filesDUB = DriveApp.getFolderById(config['folder-id-for-spreadsheet-data']).searchFiles("title contains '"+config['fileName-DUB'].replace("'","\'")+"'");
  
  // Loop thru the files to get file id for each
  while (filesUS2.hasNext()) {
    var fileUS2 = filesUS2.next();
     
    var nameUS2 = fileUS2.getName();
    var typeUS2 = fileUS2.getMimeType();
    var urlUS2 = fileUS2.getUrl();
    var fileIdUS2 = fileUS2.getId();
  }
  
    while (filesSYD.hasNext()) {
    var fileSYD = filesSYD.next();
     
    var nameSYD = fileSYD.getName();
    var typeSYD = fileSYD.getMimeType();
    var urlSYD = fileSYD.getUrl();
    var fileIdSYD = fileSYD.getId();
  }
  
    while (filesDUB.hasNext()) {
    var fileDUB = filesDUB.next();
     
    var nameDUB = fileDUB.getName();
    var typeDUB = fileDUB.getMimeType();
    var urlDUB = fileDUB.getUrl();
    var fileIdDUB = fileDUB.getId();
  }
  
    var source_sheetUS2 = SpreadsheetApp.openById(fileIdUS2);
    source_sheetUS2 = source_sheetUS2.getSheetByName(config['sourceSheetName-US2']);
    var source_rangeUS2 = source_sheetUS2.getDataRange(); //getRange("A:ZZ");
    var targetUS2 = SpreadsheetApp.getActiveSpreadsheet();
    var target_sheetUS2 = targetUS2.getSheetByName(config['targetSheetName-US2']);
    
  //source_sheet = source_sheet.getSheetByName("signups-us2");
 // var source_range = source_sheet.getDataRange(); //getRange("A:ZZ");
 // var A1Range = source_range.getA1Notation();
    //get the data values in range
 // var SData = source_range.getValues();
///var target_range = target_sheet.getRange("A:ZZ");
 // source_range.copyTo(target_range);
  
  //target_sheet.getRange(5,1,SData.length, SData.width).setValues(SData);
  
  //target_sheet.getRange(A1Range).setValues(SData); */
  
  
}






/**
 * Convert Excel file to Sheets
 * @param {Blob} excelFile The Excel file blob data; Required
 * @param {String} filename File name on uploading drive; Required
 * @param {Array} arrParents Array of folder ids to put converted file in; Optional, will default to Drive root folder
 * @return {Spreadsheet} Converted Google Spreadsheet instance
 **/
function convertExcel2Sheets(excelFile, filename, arrParents) {
  
  var parents  = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
  if ( !parents.isArray ) parents = []; // make sure parents is an array, reset to empty array if not
  
  // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
  var uploadParams = {
    method:'post',
    contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
    contentLength: excelFile.getBytes().length,
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: excelFile.getBytes()
  };
  
  // Upload file to Drive root folder and convert to Sheets
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
    
  // Parse upload&convert response data (need this to be able to get id of converted sheet)
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());

  // Create payload (body) data for updating converted file's name and parent folder(s)
  var payloadData = {
    title: filename, 
    parents: []
  };
  if ( parents.length ) { // Add provided parent folder(s) id(s) to payloadData, if any
    for ( var i=0; i<parents.length; i++ ) {
      try {
        var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
        payloadData.parents.push({id: parents[i]});
      }
      catch(e){} // fail silently if no such folder id exists in Drive
    }
  }
  // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
  var updateParams = {
    method:'put',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };
  
  // Update metadata (filename and parent folder(s)) of converted sheet
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);
  
  return SpreadsheetApp.openById(fileDataResponse.id);
}