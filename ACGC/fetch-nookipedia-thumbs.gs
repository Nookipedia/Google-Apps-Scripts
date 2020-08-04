function GrabNookipediaImages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var searchRange = sheet.getRange(2, 1, lastRow, lastColumn);
  var searchValues = searchRange.getValues();
  
  for(i = 0; i < lastRow - 1; i++) {
    var imgurl = 'https://nookipedia.com/wiki/Special:Redirect/file/' + toTitleCase(searchValues[i][1]).replace(/ /g, "_") + '_PG_Icon.png';
    var options = {
      'followRedirects': false,
      'muteHttpExceptions': true
    };
    var response = UrlFetchApp.fetch(imgurl, options).getHeaders();
    sheet.getRange('C'.concat(i + 2)).setValue('=IMAGE("' + response['Location'] + '")');
  }

  // Converts string to title case
  function toTitleCase(str) {
    return str.replace(/(^|\s|-)\S/g, function(t) { return t.toUpperCase() });
  }
}
