function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌟 Custom Actions')
    .addItem('Format Phone Number', 'formatPhoneNumber')
    .addItem('📥 Save Selection as CSV', 'exportSelectionToCSV')
    .addItem('📥 Save Selection as VCF', 'exportSelectionToVCF')
    .addToUi();
}

