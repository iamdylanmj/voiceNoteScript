
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Audio')
    .addItem('Play Audio', 'playAudio')
    .addToUi();
}

function extractFileId(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function playAudio() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  var cellValue = cell.getValue();
  
  var fileId = extractFileId(cellValue);
  
  if (fileId) {
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var base64Content = Utilities.base64Encode(blob.getBytes());
      var contentType = blob.getContentType();
      
      var html = '<audio controls><source src="data:' + contentType + ';base64,' + base64Content + '">Your browser does not support the audio element.</audio>';
      
      var ui = HtmlService.createHtmlOutput(html)
        .setWidth(300)
        .setHeight(100);
      
      SpreadsheetApp.getUi().showModelessDialog(ui, 'Play Audio');
    } catch (e) {
      SpreadsheetApp.getUi().alert('Error accessing the file. Make sure the file exists and you have permission to access it.');
    }
  } else {
    SpreadsheetApp.getUi().alert('Please select a cell with a valid Google Drive link or file ID.');
  }
}
