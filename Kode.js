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
  const ss = SpreadsheetApp.openById('1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34');
  const sheet = ss.getSheetByName('Pending_Pendaftaran');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map((row, index) => {
    return {
      rowIndex: index,
      tanggal: row[1],
      nama: row[2],
      nik: row[3],
      wa: row[7],
      pokok: row[17],
      wajib: row[18]
    };
  });
}

function processPendingMember(action, rowIndex) {
  const processor = new InputTransactions({
    method: action === 'approve' ? 'approveMember' : 'rejectMember',
    rowIndex: rowIndex,
    data: {}
  });
  return processor.transactionEntries();
}