/**
 * Export the selected range of data as a vCard (.vcf) file with multiple contacts
 * and an optional user-defined label/group applied to all contacts.
 */
function exportSelectionToVCF() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection || selection.getNumRows() * selection.getNumColumns() <= 1) {
    ui.alert('Only one cell selected. Please select a range of cells.');
    return;
  }

  // Prompt user for the label (group name)
  const labelResponse = ui.prompt(
    'Set Contact Label/Group',
    'Enter the label/group name to apply to ALL exported contacts\n' +
    '(e.g., "Customers", "2025 Leads", "VIPs")\n\n' +
    'Leave blank for no label:',
    ui.ButtonSet.OK_CANCEL
  );

  if (labelResponse.getSelectedButton() !== ui.Button.OK) {
    return; // User canceled
  }

  const customLabel = labelResponse.getResponseText().trim();

  const values = selection.getValues();

  // Build vCard content
  let vcfContent = '';

  // Assume first row might be headers
  const firstRow = values[0].map(cell => ('' + cell).toLowerCase().trim());
  const hasHeaders = firstRow.some(cell => 
    cell.includes('name') || cell.includes('phone') || cell.includes('email') ||
    cell.includes('company') || cell.includes('title') || cell.includes('org')
  );

  const startRow = hasHeaders ? 1 : 0;

  for (let i = startRow; i < values.length; i++) {
    const row = values[i];

    // Extract fields
    const fullName  = getCell(row, firstRow, ['full name', 'name', 'fn']) || '';
    const firstName = getCell(row, firstRow, ['first name', 'given name']) || '';
    const lastName  = getCell(row, firstRow, ['last name', 'surname', 'family name']) || '';
    const company   = getCell(row, firstRow, ['company', 'org', 'organization']) || '';
    const title     = getCell(row, firstRow, ['title', 'job title', 'position']) || '';
    const phone     = getCell(row, firstRow, ['phone', 'mobile', 'cell']) || '';
    const email     = getCell(row, firstRow, ['email']) || '';

    // Skip empty rows
    if (!fullName && !firstName && !lastName && !phone && !email) continue;

    // Build vCard
    vcfContent += 'BEGIN:VCARD\r\n';
    vcfContent += 'VERSION:3.0\r\n';

    // Structured name (N:) - Last;First;;;
    if (lastName || firstName) {
      vcfContent += `N:${escapeVCard(lastName)};${escapeVCard(firstName)};;;\r\n`;
    }

    // Formatted name (FN:) - fallback to combining first + last
    const fn = fullName || `${firstName} ${lastName}`.trim();
    if (fn) {
      vcfContent += `FN:${escapeVCard(fn)}\r\n`;
    }

    if (company)  vcfContent += `ORG:${escapeVCard(company)}\r\n`;
    if (title)    vcfContent += `TITLE:${escapeVCard(title)}\r\n`;
    if (phone)    vcfContent += `TEL;TYPE=CELL,VOICE:${escapeVCard(phone)}\r\n`;
    if (email)    vcfContent += `EMAIL;TYPE=INTERNET:${escapeVCard(email)}\r\n`;

    // Add the user-defined label/group (this is the new part!)
    if (customLabel) {
      vcfContent += `CATEGORIES:${escapeVCard(customLabel)}\r\n`;
    }

    vcfContent += 'END:VCARD\r\n';
  }

  if (!vcfContent) {
    ui.alert('No valid contact data found in the selection.');
    return;
  }

  // File name
  const sheetName = sheet.getName();
  const rangeNotation = selection.getA1Notation();
  const fileName = 'Contacts-' + sheetName + '-' + rangeNotation + '.vcf';

  // Create file in Drive root
  const file = DriveApp.getRootFolder().createFile(fileName, vcfContent, MimeType.PLAIN_TEXT);

  // Show success dialog
  const htmlOutput = HtmlService.createHtmlOutput(
    'Your vCard file "' + fileName + '" has been saved in the main folder of your Google Drive.<br><br>' +
    '<strong>Contains ' + (values.length - startRow) + ' contact(s).</strong><br><br>' +
    (customLabel ? '<strong>Label/Group applied: ' + customLabel + '</strong><br><br>' : '') +
    '<a href="' + file.getUrl() + '" target="_blank">Click here to open/download the file.</a>' +
    '<br><br><input type="button" value="Close" onclick="google.script.host.close()">'
  ).setWidth(450).setHeight(220);

  ui.showModalDialog(htmlOutput, 'Export Complete');
}

// Helper: Find column index by possible header names
function getCell(row, headers, possibleNames) {
  for (let name of possibleNames) {
    const index = headers.indexOf(name);
    if (index !== -1 && row[index]) {
      return ('' + row[index]).trim();
    }
  }
  return '';
}

// Escape special characters for vCard ( ; , \n \r )
function escapeVCard(str) {
  return ('' + str)
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '');
}


// Helper: Find column index by possible header names
function getCell(row, headers, possibleNames) {
  for (let name of possibleNames) {
    const index = headers.indexOf(name);
    if (index !== -1 && row[index]) {
      return ('' + row[index]).trim();
    }
  }
  return '';
}

// Escape special characters for vCard ( ; , \n \r )
function escapeVCard(str) {
  return ('' + str)
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '');
}
