# General
These functions are written to be included in a Google App Scripts module.  This is accessible from Sheets via the "Extension" menu.  Paste in the code from each module; then save.

## Custom-Action
Adds a menu option "Custom Actions" with these three entries:

## Format Phone Number
Calls the `formatPhoneNumber()` function.  This removes all formatting from the contents of the selected columns.  It is intended for US 10-digit phone numbers.  Anything except digits is removed allowing the sheet to use the default phone number format.

## Save Contacts to CSV
Calls the `exportContactsToCSV()` function.  This outputs the contents of columns A-H as a CSV.  The file is output in a format that works with import into Google Contacts but may work elsewhere.  The full name includes the title, first name, and last name.  The "Zone" and "Stake" become custom fields.  The "Release" becomes a significant date.  Requires the `getTimestamp()` function.

## Save Contacts to VCF
Calls the `exportContactsToVCF()` function.  This outputs the contents of columns A-H as a VCF.  The file should import successfully into any CardDAV compliant contacts app.  The full name includes the title, first name, and last name.  The "Zone", "Stake", and "Release" all become custom fields.  Requires the `getTimestamp()` and `escapeVCard()` functions.