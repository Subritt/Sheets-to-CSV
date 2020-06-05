/*
############################################################################################################################################################################################################
############################################################################### Creator : Subritt Burlakoti                   ##############################################################################
############################################################################### Email : subritt.burlakoti@es.cloudfactory.com ##############################################################################
############################################################################### Delivery Solution Project Support             ##############################################################################
############################################################################################################################################################################################################
*/

// On Open Trigger
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Export as CSV", functionName: "main_function"}];
  ss.addMenu("CSV", csvMenuEntries);
};

// Returns 'true' if variable d is a date object.
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

// Test if value is a date and if so format
// otherwise, reflect input variable back as-is. 
function isDate(sDate) {
  if (isValidDate(sDate)) {
    sDate = Utilities.formatDate(new Date(sDate), "GMT+5:45", "MM/dd/yyyy");
  }
  return sDate.toString();
}

/*
Caller function
*/
function main_function(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_list = spreadsheet.getSheets(); // extracting all sheet
  var folder_created = DriveApp.createFolder(spreadsheet.getName() + " - " + new Date().getTime()); // drive created
  
  var download_url = [["Sheet Name", "Download URL"]];
  sheet_list.forEach(function(sheet){
    download_url.push([sheet.getName(), saveAsCSV(folder_created, sheet)]);
  });
  
  //  folder_created.createFile("CSV Download URL", "", MimeType.GOOGLE_SHEETS)
  
  var resource = {
    title: "CSV Download URL",
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folder_created.getId() }]
  }
  var url_file = Drive.Files.insert(resource);
  SpreadsheetApp.openById(url_file.id)
                .getActiveSheet()
                .getRange(1, 1, download_url.length, download_url[0].length)
                .setValues(download_url);
  
  Logger.log('CSV files created in a folder named ' + folder_created.getName())
  Browser.msgBox('CSV files created in a folder named ' + folder_created.getName());
}

/*
function to Save as CSV
*/
function saveAsCSV(folder_created, sheet) {
  var filename = sheet.getName(); // CSV file name
  
  var folder = folder_created.getId(); // Folder ID
  
  var csv = "";
  var v = sheet.getDataRange()
               .getValues();
  
  /*
  extracting values
  */
  
  try{
    v.forEach(function(e) {
      //   Logger.log(e[e.length-1]);
      e.forEach(function(c){
        if(c != e[e.length-1]){
          if(isValidDate(c) == true){
            c = isDate(c);
          }
          csv += c + ",";
        } else{
          csv += c;
        }
      });
      csv += "\n";
      //    csv += e.join(",") + "\n";
    });
    Logger.log(csv);
    
    //    DriveApp.getFolderById(folder)
    //    .createFile(filename, csv, MimeType.CSV);
    
    var url = DriveApp.getFolderById(folder)
    .createFile(filename, csv, MimeType.CSV)
    .getDownloadUrl()
    .replace("?e=download&gd=true","");
    return url;
    
  }catch(err){
    Logger.log(err);
    Browser.msgBox(err);
  }
}