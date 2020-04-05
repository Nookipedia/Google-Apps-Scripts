function GenerateTable() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var searchRange = sheet.getRange(2, 1, lastRow, lastColumn);
  var searchValues = searchRange.getValues();
  
  var content = "{{TableHeader|game=ACNH|type=furniture|collection=other|title=Furniture in {{NH|nolink}}}}\n\n";
  
  for(i = 0; i < lastRow - 1; i++) {
    content += '{{TableContent\n';
    
    // Item name
    content += '| ' + toTitleCase(searchValues[i][0]) + '\n';
    
    // Image
    content += '| [[File:' + toTitleCase(searchValues[i][0]) + ' NH Icon.png|64px|' + toTitleCase(searchValues[i][0]) + ']]\n';
    
    // Collection
    if (!searchValues[i][9]) {
      content += '| -\n';
    } else {
      content += '| ' + searchValues[i][9] + '\n';
    }
    
    // Buy Price
    if (!searchValues[i][6]) {
       content += '| -\n';
    } else if (searchValues[i][6] == 'NFS') {
      content += '| Not for sale\n';
    } else {
      content += '| {{Material|Bells|' + numberWithCommas(searchValues[i][6]) + '}}\n';
    }
    
    // Sell Price
    if (!searchValues[i][7]) {
       content += '| -\n';
    } else {
      content += '| {{Material|Bells|' + numberWithCommas(searchValues[i][7]) + '}}\n';
    }
    
    // Available From
    if (!searchValues[i][8]) {
       content += '| -\n';
    } else if (searchValues[i][8] == "Nook's Cranny") {
      content += '| [[Nook\'s store]]\n';
    } else if (searchValues[i][8] == "Nook Miles Exchange") {
      content += '| [[Nook Miles]]\n';
    } else {
       content += '| [[' + searchValues[i][8] + ']]\n';
    }
    
    // Recipe
    if (searchValues[i][3] == 'Y') {
      content += '| TODO\n';
    } else {
      content += '| N/A\n';
    }
    
    // Customizations
    if (searchValues[i][4] == 'Y') {
      content += '| Yes\n';
    } else if (searchValues[i][4] == 'N') {
      content += '| None\n';
    } else {
      content += '| -\n';
    }
    
    // Colors 
    content += '| -\n';
    
    // HHA Theme
    content += '| -\n';
    
    // Style
    content += '| -\n';
    
    // Size
    content += '| -\n';
    
    // Info
    content += '| -\n';
    
    content += '}}\n';
    content += '\n';
  }
  
  content += '{{TableFooter}}\n';
  
  newFile = DriveApp.createFile('furniture-output.txt', content); // Creates a new text file in Google Drive root
};

// Converts string to title case (used since source sheet uses sentence case)
function toTitleCase(str) {
  return str.toString().replace(/\b[\w']+\b/g, function(txt){
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

// Adds commas to numbers (used for bell prices)
function numberWithCommas(int) {
  return int.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
