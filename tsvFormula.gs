/*
 * script to export data in all sheets in the current spreadsheet as individual tsv files
 * files will be named according to the name of the sheet
 * author: Michael Derazon
 * contributor: xFanatical
*/

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tsvMenuEntries = [{name: "export as TSV files", functionName:   "saveAstsv"}];
  ss.addMenu("tsv", tsvMenuEntries);
};

function saveAstsv() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  // create a folder from the name of the spreadsheet
  var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/   /g,'_') + '_tsv_' + new Date().getTime());
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // append ".tsv" extension to the sheet name
    fileName = sheet.getName() + ".tsv";
    // convert all available sheet data to tsv format
    var tsvFile = convertRangeTotsvFile_(fileName, sheet);
    // create a file in the Docs List with the given name and the tsv data
    folder.createFile(fileName, tsvFile);
  }
  Browser.msgBox('Files are waiting in a folder named ' + folder.getName());
}

function convertRangeTotsvFile_(tsvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var formula = activeRange.getFormulas();
    var tsvFile = undefined;

    // loop through the data in the range and build a string with the tsv  data
    if (data.length > 1) {
      var tsv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (formula[row][col] !== '') {
            data[row][col] = formula[row][col]
          }
          if (data[row][col].toString().indexOf("\t") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          tsv += data[row].join("\t") + "\r\n";
        }
        else {
        tsv += data[row].join("\t");
             }
      }
      tsvFile = tsv;
    }
    return tsvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
