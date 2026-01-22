/**
 * Strips all non-digit characters from phone numbers in the selected range(s),
 * leaving only the raw digits (e.g., "(123) 456-7890" → "1234567890").
 * Works with single cell or multiple cells.
 */
function formatPhoneNumber() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const selection = spreadsheet.getSelection();

  const rangeList = selection.getActiveRangeList();

  if (!rangeList) {
    ui.alert('No range is selected. Please select the cell(s) you want to clean.');
    return;
  }

  const ranges = rangeList.getRanges();

  let modifiedCount = 0;
  const regex = /[^\d]/g;  // Matches anything that is NOT a digit

  ranges.forEach(range => {
    const values = range.getValues();

    const newValues = values.map(row =>
      row.map(cell => {
        if (typeof cell === 'string' && cell.trim() !== '') {
          const cleaned = cell.replace(regex, '');
          if (cleaned !== cell) modifiedCount++;
          return cleaned;
        }
        return cell;  // Leave numbers, blanks, etc. unchanged
      })
    );

    if (modifiedCount > 0) {
      range.setValues(newValues);
    }
  });

  if (modifiedCount > 0) {
    ui.alert(`Success! Cleaned ${modifiedCount} cell(s) to raw digits only (e.g., 1234567890).`);
  } else {
    ui.alert('No formatted phone numbers were found in the selected range(s).');
  }
}
