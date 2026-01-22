// Escape special characters for vCard (; , \n \r \)
function escapeVCard(str) {
  return ('' + str)
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '');
}
