/**
 * Export the selected range of data as a CSV
 */
function exportSelectionToCSV() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selection = sheet.getActiveRange();
  
  if (!selection || selection.getNumRows() * selection.getNumColumns() <= 1) {
    SpreadsheetApp.getUi().alert('Only one cell selected.  Please select a range of cells.');
    return;
  }
  
  var values = selection.getValues();
  
  Utilities.sleep(100); // Give it a moment

  var csvContent = '';
  values.forEach(function(row) {
    // This handles values containing commas by enclosing them in quotes.
    var cleanRow = row.map(function(cell) {
        var a = '' + cell;
        if (a.indexOf(",") > -1) {
            a = '"' + a.replace(/"/g, '""') + '"';
        }
        return a;
    });
    csvContent += cleanRow.join(',') + '\n';
  });

  var sheetName = sheet.getName();
  var rangeNotation = selection.getA1Notation();
  var fileName = 'Export-' + sheetName + '-' + rangeNotation + '.csv';
  
  var file = DriveApp.getRootFolder().createFile(fileName, csvContent, MimeType.CSV);

  var htmlOutput = HtmlService.createHtmlOutput(
    'Your file "' + fileName + '" is saved in the main folder of your Google Drive. <br><br>' +
    '<a href="' + file.getUrl() + '" target="_blank">Click here to open the file.</a>' +
    '<br><br><input type="button" value="Close" onclick="google.script.host.close()">'
  ).setWidth(400).setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Complete');
}
