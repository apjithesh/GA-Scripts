/* function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var searchMenuEntries = [ {name: "Search in all files", functionName: "search"}];
  ss.addMenu("Search Google Drive", searchMenuEntries);
}
  */

/* Add Moment library by going to Resources -> Libraries in script editor and using this key - MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48 */

function TryOut() { //This is a function where I tried and tested many google sheet functions
  // Prompt the user for a search term
  //var searchTerm = Browser.inputBox("Enter the string to search for:");
 var searchTerm = "product-signup";
  var moment = Moment.load();
  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SSID=ss.getId();
  var sheet = ss.getActiveSheet();
    var config_sheet = ss.getSheetByName("config-sheet");
  var regions = new Array();
  regions = config_sheet.getRange('B11').getValue().split(",");
 //Browser.msgBox(regions[1]);
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
  var files = DriveApp.getFolderById("0B-jvR7IWXrOqVlQxTUlfVUY1RGs").searchFiles("title contains '"+searchTerm.replace("'","\'")+"'"); //" + moment().format('YYYY-MM-DD') + "
  // create an array to store our data to be written to the sheet
  var output = [];
    //Browser.msgBox("Files Length " + files.hasNext());
  if (files.hasNext()) {
  //Browser.msgBox("No files found");
  
  // Loop through the results and get the file name, file type, and URL
  while (files.hasNext()) {
    var file = files.next();
     
    var name = file.getName();
    var type = file.getMimeType();
    var url = file.getUrl();
    var fileid = file.getId();
    var filecreateDate = file.getDateCreated();
    // push the file details to our output array (essentially pushing a row of data)
    if ((moment().diff(filecreateDate,"hours")) > 13) {
    file.setTrashed(true);
    } else {
      output.push([name, type, url, fileid, filecreateDate, moment().diff(filecreateDate,"hours"), moment().format('YYYY-MM-DD')]);
    }
  }
  // write data to the sheet
  
 // var formula = '=importrange("' + fileid + '","signups-us2!A:ZZ")';
    if (output.length > 0) {
  sheet.getRange(2, 1, output.length, 7).setValues(output);
    }
 // sheet.getRange(5, 1).setValue(formula);
  }
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

function getProductpullConfigs () {
//Find the folder where the spreadsheets are stored with data
  var config = [
  ];
// Get config spreadsheet
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config_sheet = ss.getSheetByName("product-dataconfig-sheet");
  var regions = new Array();
  regions = config_sheet.getRange('B4').getValue().split(",");
  config['folder-id-for-spreadsheet-data'] =  config_sheet.getRange('B2').getValue();
  //Browser.msgBox("folder id received is " + config['folder-id-for-spreadsheet-data']);
  
  // Loop thru and create the config list
  var rowNum = 5;
  for (i=0; i < regions.length; i++) {
    //config.push(['fileName-'+regions[i],config_sheet.getRange('B'+rowNum).getValue()]);
    config['fileName-'+regions[i]] = config_sheet.getRange('B'+rowNum).getValue();
    //Browser.msgBox("config " + config['fileName-'+regions[i]] + " regions is " + regions[i] + " Value should be " + config_sheet.getRange('B'+rowNum).getValue());
    rowNum++;
    config['sourceSheetName-'+regions[i]] = config_sheet.getRange('B'+rowNum).getValue();
    //Browser.msgBox("config " + config['sourceSheetName-'+regions[i]] + " regions is " + regions[i] + " Value should be " + config_sheet.getRange('B'+(rowNum)).getValue());
    rowNum++;
    config['targetSheetName-'+regions[i]] = config_sheet.getRange('B'+rowNum).getValue();
    //Browser.msgBox("config " + config['targetSheetName-'+regions[i]] + " regions is " + regions[i] + " Value should be" + config_sheet.getRange('B'+(rowNum)).getValue());
    //Browser.msgBox("rownumber " + rowNum)
    rowNum++;
  } 
  config['regions'] = config_sheet.getRange('B4').getValue();
  config['target-sheetID'] = config_sheet.getRange('B3').getValue();
 
  return config;
}

function getProductSignupData () {
// Gets data from other files that has the product data and populate into this sheet
  var config = getProductpullConfigs();
  var moment = Moment.load();
  // Loop thru and process the files and copy data into this spreadsheet
  var regions = config['regions'].split(",");
  //Browser.msgBox(config['folder-id-for-spreadsheet-data'] + config['regions'] + " overall configs " + config );
  for (i=0; i < regions.length ; i++) {
    var files = DriveApp.getFolderById(config['folder-id-for-spreadsheet-data']).searchFiles("title contains '"+config['fileName-' + regions[i]].replace("'","\'")+"'");
    if (files.hasNext()) {
      // Loop thru the files to get file id for each
      while (files.hasNext()) {
        var file = files.next();
        
        var name = file.getName();
        var type = file.getMimeType();
        var url = file.getUrl();
        var fileId = file.getId();
        var filecreateDate = file.getDateCreated();
        
         if ((moment().diff(filecreateDate,"hours")) <= 8) {
        
          //Browser.msgBox("Got into else and file createdate is " + filecreateDate);
          var source_sheet = SpreadsheetApp.openById(fileId);
          source_sheet = source_sheet.getSheetByName(config['sourceSheetName-' + regions[i]]);
          var source_range = source_sheet.getDataRange(); //getRange("A:ZZ");
          var A1Range = source_range.getA1Notation();
          //get the data values in range
          var SData = source_range.getValues();
          var target = SpreadsheetApp.openById(config['target-sheetID'])//getActiveSpreadsheet();
          var target_sheets = target.getSheets();
          //Browser.msgBox("Sheets are # " + target_sheets);
          var target_sheet = target.getSheetByName(config['targetSheetName-' + regions[i]]);
          if (!target_sheet) {
            target_sheet = target.insertSheet(config['targetSheetName-' + regions[i]], target_sheets.length + 1);
          }
          target_sheet.clear();
          target_sheet.getRange(A1Range).setValues(SData);
        
        }
      }

    }
    
  }
  
  var filestoDelete = DriveApp.getFolderById(config['folder-id-for-spreadsheet-data']).searchFiles("title contains 'product-signup'");

  while (filestoDelete.hasNext()) {
        var file = filestoDelete.next();
        
        var name = file.getName();
        var type = file.getMimeType();
        var url = file.getUrl();
        var fileId = file.getId();
        var filecreateDate = file.getDateCreated();
    //Browser.msgBox("Time diff " + moment().diff(filecreateDate,"hours"));
      if ((moment().diff(filecreateDate,"hours")) > 8) {
          file.setTrashed(true);
        }
  }
  MailApp.sendEmail("jithesh@sumologic.com", "Google Script for Extracting product data", "The Script has completed at " + moment().format('YYYY-MM-DD H:mm'));
}

