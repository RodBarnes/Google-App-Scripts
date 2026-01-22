/**
 * Export columns A:H of the active sheet as a CSV file with custom fields & events.
 * Fixed column mapping:
 *   A: Title
 *   B: First Name
 *   C: Last Name
 *   D: Zone     → Custom Field "Zone"
 *   E: Stake    → Custom Field "Stake"
 *   F: Release  → Custom Event "Release" (forced to YYYY-MM-DD)
 *   G: Phone
 *   H: Email    → E-mail 1 - Value with Type "Home"
 */
function exportContactsToCSV() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data found in columns A:H (need at least headers + 1 contact).');
    return;
  }

  const values = sheet.getRange(1, 1, lastRow, 8).getValues(); // A1:H[lastRow]

  // Build CSV content with OFFICIAL Google headers
  let csvContent = 'Given Name,Family Name,Name Prefix,Phone 1 - Value,E-mail 1 - Value,E-mail 1 - Type,Custom Field 1 - Label,Custom Field 1 - Value,Custom Field 2 - Label,Custom Field 2 - Value,Event 1 - Label,Event 1 - Value\n';

  for (let i = 1; i < values.length; i++) { // Skip header row
    const row = values[i];

    const title = ('' + (row[0] || '')).trim();
    const firstName = ('' + (row[1] || '')).trim();
    const lastName = ('' + (row[2] || '')).trim();
    const zone = ('' + (row[3] || '')).trim();
    const stake = ('' + (row[4] || '')).trim();
    let release = row[5]; // Keep as-is initially (may be Date or string)
    const phone = ('' + (row[6] || '')).trim();
    const email = ('' + (row[7] || '')).trim();

    if (!firstName && !lastName && !phone && !email) continue;

    // Format full name with title
    let fullName = [title, firstName, lastName].filter(Boolean).join(' ').trim();
    if (!fullName) fullName = 'Contact';

    // Handle Release date: Force to YYYY-MM-DD
    let formattedRelease = '';
    if (release) {
      if (release instanceof Date) {
        // This is the key fix for Date objects from Sheets
        formattedRelease = Utilities.formatDate(release, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        // Handle string dates
        const strRelease = ('' + release).trim();

        // Already YYYY-MM-DD?
        if (/^\d{4}-\d{2}-\d{2}$/.test(strRelease)) {
          formattedRelease = strRelease;
        }
        // Common formats: MM/DD/YYYY, DD/MM/YYYY, etc.
        else if (strRelease.match(/^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}$/)) {
          const parts = strRelease.split(/[\/-]/);
          let month, day, year;

          if (parts[0].length <= 2 && parseInt(parts[0]) <= 12) {
            // Assume MM/DD/YYYY
            month = parts[0].padStart(2, '0');
            day = parts[1].padStart(2, '0');
            year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
          } else {
            // Assume DD/MM/YYYY
            day = parts[0].padStart(2, '0');
            month = parts[1].padStart(2, '0');
            year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
          }

          formattedRelease = `${year}-${month}-${day}`;
        }
        // Fallback
        else {
          formattedRelease = strRelease;
        }
      }
    }

    // Escape for CSV
    const escapeCsv = (str) => `"${('' + str).replace(/"/g, '""')}"`;

    // Add row with correct headers
    csvContent += `${escapeCsv(firstName)},${escapeCsv(lastName)},${escapeCsv(title)},${escapeCsv(phone)},${escapeCsv(email)},Home,Zone,${escapeCsv(zone)},Stake,${escapeCsv(stake)},Release,${escapeCsv(formattedRelease)}\n`;
  }

  if (csvContent === 'Given Name,Family Name,Name Prefix,Phone 1 - Value,E-mail 1 - Value,E-mail 1 - Type,Custom Field 1 - Label,Custom Field 1 - Value,Custom Field 2 - Label,Custom Field 2 - Value,Event 1 - Label,Event 1 - Value\n') {
    ui.alert('No valid contact data found in columns A:H.');
    return;
  }

  // Generate filename with timestamp (using your existing function)
  const fileName = 'Contacts-Service-Missionary-' + getTimestamp() + '.csv';

  const file = DriveApp.getRootFolder().createFile(fileName, csvContent, MimeType.CSV);

  const htmlOutput = HtmlService.createHtmlOutput(
    'Your CSV file "' + fileName + '" has been saved in the main folder of your Google Drive.<br><br>' +
    '<strong>After import, rename the "Imported" label to "Service".</strong><br><br>' +
    '<a href="' + file.getUrl() + '" target="_blank">Click here to open/download the file.</a>' +
    '<br><br><input type="button" value="Close" onclick="google.script.host.close()">'
  ).setWidth(450).setHeight(180);

  ui.showModalDialog(htmlOutput, 'Export Complete');
}
