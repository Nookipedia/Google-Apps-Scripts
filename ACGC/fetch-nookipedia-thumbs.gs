/* 
  Automatically find + place thumbnails from Nookipedia into the spreadsheet you are currently viewing.
  Update the NAME_COLUMN, IMG_COLUMN, and FILE_SUFFIX below as needed.
*/

var NAME_COLUMN = 1    // 0 for column A, 1 for column B, etc.
var IMG_COLUMN = 'C'
var FILE_SUFFIX = '_PG_Field_Sprite.png'

function GrabNookipediaImages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var searchRange = sheet.getRange(2, 1, lastRow, lastColumn);
  var searchValues = searchRange.getValues();
  
  for(i = 0; i < lastRow - 1; i++) {
    var imgurl = 'https://nookipedia.com/wiki/Special:Redirect/file/' + toTitleCase(searchValues[i][NAME_COLUMN]).replace(/ /g, "_") + FILE_SUFFIX;
    var options = {
      'followRedirects': false,
      'muteHttpExceptions': true
    };
    var response = UrlFetchApp.fetch(imgurl, options).getHeaders();
    sheet.getRange(IMG_COLUMN.concat(i + 2)).setValue('=IMAGE("' + response['Location'] + '")');
  }

  // Converts string to title case
  function toTitleCase(str) {
    return str.replace(/(^|\s|-)\S/g, function(t) { return t.toUpperCase() });
  }
}
