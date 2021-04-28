function DownloadImages() {
  var FOLDERID = 'INSERT_FOLDER_ID'
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var searchRange = sheet.getRange(2, 1, lastRow, lastColumn);
  var searchFormulas = searchRange.getFormulas();
  var searchValues = searchRange.getValues();

  for (i = 0; i < lastRow - 1; i++) {
    var variationName = '';
    // Append variation name, if exists:
    if (searchValues[i][3] != 'NA') {
      variationName += ' (' + toTitleCase(searchValues[i][3]) + ')';
    }
    var imgurl = getImageUrl(searchFormulas[i][1]);
    var imgurl2 = getImageUrl(searchFormulas[i][2]);
    var image = UrlFetchApp.fetch(imgurl).getBlob().getAs('image/png').setName(toTitleCase(searchValues[i][0]) + variationName + ' NH Icon.png');
    var image2 = UrlFetchApp.fetch(imgurl2).getBlob().getAs('image/png').setName(toTitleCase(searchValues[i][0]) + variationName + ' NH Storage Icon.png');
    var folder = DriveApp.getFolderById(FOLDERID);
    folder.createFile(image);
    console.log('Downloaded ' + toTitleCase(searchValues[i][0]) + variationName + ' NH Icon.png');
    folder.createFile(image2);
    console.log('Downloaded ' + toTitleCase(searchValues[i][0]) + variationName + ' NH Storage Icon.png');
    variationName = '';
  }
}

// Converts string to title case (used since source sheet uses sentence case)
function toTitleCase(str) {
  return str.toString().replace(/\b[\w']+\b/g, function(txt){
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

// Extract image URL from an image formula
function getImageUrl(formula) {
  var regex = /=IMAGE\("(.*)"/i;
  var matches = formula.toString().match(regex);
  return matches ? matches[1] : null;
}
