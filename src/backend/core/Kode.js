function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Arrahmah')
    .addItem('Home', 'fnHome')
    .addItem('Pendaftaran Modern', 'fnDaftarAnggotaModern')
    .addItem('Bulk Import Anggota', 'fnBulkImportAnggota')
    .addItem('Approval Pendaftaran', 'fnApprovalDashboard')
    .addToUi();
}

function fnDaftarAnggotaModern() {
  var template = HtmlService.createTemplateFromFile('daftarAnggotaModern');
  var htmlOutput = template.evaluate().setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Pendaftaran Anggota (Modern)');
}

function fnBulkImportAnggota() {
  var template = HtmlService.createTemplateFromFile('bulkImportAnggota');
  var htmlOutput = template.evaluate().setWidth(1200).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bulk Import Anggota (Excel Paste)');
}

function fnHome() {
  var template = HtmlService.createTemplateFromFile('home');
  var htmlOutput = template.evaluate().setWidth(1200).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function doGet() {
  var template = HtmlService.createTemplateFromFile('daftarAnggotaModern');
  return template.evaluate()
    .setTitle('Pendaftaran Anggota Modern')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function fnApprovalDashboard() {
  var template = HtmlService.createTemplateFromFile('approvalDashboard');
  var htmlOutput = template.evaluate().setWidth(1200).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Approval Pendaftaran Anggota');
}

function getPendingMembersData() {
  return MemberService.getPendingData();
}

function processPendingMember(action, rowIndex) {
  const processor = new InputTransactions({
    method: action === 'approve' ? 'approveMember' : 'rejectMember',
    rowIndex: rowIndex,
    data: {}
  });
  return processor.transactionEntries();
}

function processMembersBulk(membersArray) {
  const processor = new InputTransactions({ method: 'bulk', data: {} });
  return processor.addMembersBulk(membersArray);
}

function processBatchApproval(indices) {
  const processor = new InputTransactions({ method: 'bulk', data: {} });
  return processor.approvePendingBulk(indices);
}