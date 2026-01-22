/**
 * Export columns A:H of the active sheet as a vCard (.vcf) file with multiple contacts.
 * Fixed column mapping:
 *   A: Title       → Used as Name Prefix in N: (e.g., "Mr.", "Elder")
 *   B: First Name
 *   C: Last Name
 *   D: Zone        → custom field X-ZONE
 *   E: Stake       → custom field X-STAKE
 *   F: Release     → custom significant date X-RELEASE;TYPE=RELEASE
 *   G: Phone
 *   H: Email
 * 
 * All contacts are assigned the fixed category/label: "Service"
 */
function exportContactsToVCF() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Automatically select columns A:H (full columns, from row 1 to last row with data)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {  // Need at least headers + 1 data row
    ui.alert('No data found in columns A:H (need at least headers + 1 contact).');
    return;
  }

  const selection = sheet.getRange(1, 1, lastRow, 8); // A1:H[lastRow]
  const values = selection.getValues();

  const customLabel = "Service";  // Hard-coded label for all contacts

  // Build vCard content
  let vcfContent = '';

  // Start from row 2 (skip headers in row 1)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    // Extract fields from fixed columns (0-based index)
    const title = ('' + (row[0] || '')).trim();    // A: Title → Name Prefix
    const firstName = ('' + (row[1] || '')).trim();    // B: First Name
    const lastName = ('' + (row[2] || '')).trim();    // C: Last Name
    const zone = ('' + (row[3] || '')).trim();    // D: Zone → custom
    const stake = ('' + (row[4] || '')).trim();    // E: Stake → custom
    const release = row[5];                          // F: Release (handle Date or string)
    const phone = ('' + (row[6] || '')).trim();    // G: Phone
    const email = ('' + (row[7] || '')).trim();    // H: Email

    // Skip completely empty rows
    if (!firstName && !lastName && !phone && !email) continue;

    // Format Release date to YYYY-MM-DD if it's a Date object or parsable string
    let formattedRelease = '';
    if (release) {
      if (release instanceof Date) {
        formattedRelease = Utilities.formatDate(release, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        const strRelease = ('' + release).trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(strRelease)) {
          formattedRelease = strRelease;
        } else if (strRelease.match(/^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}$/)) {
          const parts = strRelease.split(/[\/-]/);
          let month, day, year;
          if (parts[0].length <= 2 && parseInt(parts[0]) <= 12) {
            month = parts[0].padStart(2, '0');
            day = parts[1].padStart(2, '0');
            year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
          } else {
            day = parts[0].padStart(2, '0');
            month = parts[1].padStart(2, '0');
            year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
          }
          formattedRelease = `${year}-${month}-${day}`;
        } else {
          formattedRelease = strRelease;
        }
      }
    }

    // Build full name (FN:) - include title/prefix if available
    let fnParts = [title, firstName, lastName].filter(Boolean);
    let fn = fnParts.join(' ').trim();
    if (!fn) fn = 'Contact';

    // Build vCard
    vcfContent += 'BEGIN:VCARD\r\n';
    vcfContent += 'VERSION:3.0\r\n';

    // Structured name (N:) - Last;First;;;
    vcfContent += `N:${escapeVCard(lastName)};${escapeVCard(firstName)};;;\r\n`;

    // Formatted name (FN:)
    vcfContent += `FN:${escapeVCard(fn)}\r\n`;

    if (phone) vcfContent += `TEL;TYPE=CELL,VOICE:${escapeVCard(phone)}\r\n`;
    if (email) vcfContent += `EMAIL;TYPE=INTERNET:${escapeVCard(email)}\r\n`;

    // Custom fields
    if (zone) vcfContent += `X-ZONE:${escapeVCard(zone)}\r\n`;
    if (stake) vcfContent += `X-STAKE:${escapeVCard(stake)}\r\n`;

    // Custom significant date (Release)
    if (formattedRelease) {
      vcfContent += `X-RELEASE;TYPE=RELEASE:${escapeVCard(formattedRelease)}\r\n`;
    }

    // Add the fixed label/group "Service"
    vcfContent += `CATEGORIES:${escapeVCard(customLabel)}\r\n`;

    vcfContent += 'END:VCARD\r\n';
  }

  if (!vcfContent) {
    ui.alert('No valid contact data found in columns A:H.');
    return;
  }

  // File name
  const sheetName = sheet.getName();
  const fileName = 'Contacts-Service-Missionary-' + getTimestamp() + '.vcf';

  // Create file in Drive root
  const file = DriveApp.getRootFolder().createFile(fileName, vcfContent, MimeType.PLAIN_TEXT);

  // Show success dialog
  const htmlOutput = HtmlService.createHtmlOutput(
    'Your vCard file "' + fileName + '" has been saved in the main folder of your Google Drive.<br><br>' +
    '<a href="' + file.getUrl() + '" target="_blank">Click here to open/download the file.</a>' +
    '<br><br><input type="button" value="Close" onclick="google.script.host.close()">'
  ).setWidth(450).setHeight(220);

  ui.showModalDialog(htmlOutput, 'Export Complete');
}
