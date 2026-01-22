function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌟 Custom Actions')
    .addItem('📠 Format Phone Number', 'formatPhoneNumber')
    .addItem('📥 Save Contacts to CSV', 'exportContactsToCSV')
    .addItem('📥 Save Contacts to VCF', 'exportContactsToVCF')
    .addToUi();
}
