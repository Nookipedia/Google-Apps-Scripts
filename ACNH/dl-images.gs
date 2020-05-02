function DownloadImages() {
  var FOLDERID = 'INSERT_FOLDER_ID'

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var searchRange = sheet.getRange(2, 1, lastRow, lastColumn);
  var searchValues = searchRange.getValues();
  
  for(i = 0; i < lastRow - 1; i++) {
    if (searchValues[i][2]) {
      var imgurl = searchValues[i][2];
      var image = UrlFetchApp.fetch(imgurl).getBlob().getAs('image/png').setName(toTitleCase(searchValues[i][0]) + ' NH Icon.png');
      var folder = DriveApp.getFolderById(FOLDERID);
      folder.createFile(image);
    }
  }
}

// Converts string to title case (used since source sheet uses sentence case)
function toTitleCase(str) {
  return str.toString().replace(/\b[\w']+\b/g, function(txt){
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}
