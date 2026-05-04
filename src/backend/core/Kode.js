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
    .addItem('Bantu Insert Data Rudi', 'sendRudiData')
    .addToUi();
}

function sendRudiData() {
  const data = {
    tanggalBergabung: "2026-04-17",
    namaAnggota: "Rudi Maulana S",
    nik: "6303112803930001",
    tempatLahir: "Martapura",
    tanggalLahir: "1993-03-28",
    jenisKelamin: "Laki-Laki",
    telponAnggota: "'085348690354",
    email: "rudimaulana2312@gmail.com",
    alamatKTP: "Jl. IR Pangeran Muhammad Noor RT. 05 RW. 02 Kel/Desa Awang Bangkal Barat Kec. Karang Intan 71352",
    alamatTinggal: "Jl. IR Pangeran Muhammad Noor RT. 05 RW. 02 Kel/Desa Awang Bangkal Barat Kec. Karang Intan 71352",
    keluargaSerumah: "Hairil Fitriana",
    hubunganKeluargaSerumah: "Istri",
    telponKeluargaSerumah: "'082350909169",
    keluargaTidakSerumah: "Hairil Fitriana",
    telponKeluargaTidakSerumah: "'082350909169",
    jenisPekerjaan: "Aparatur Sipil Negara (ASN)",
    kantorPekerjaan: "Rutan Kelas II B Barabai",
    alamatKantor: "Hulu Sungai Tengah, Kalimantan Selatan 71352 RT. 00 RW. 00 Kel/Desa Barabai Kec. Barabai 71352",
    namaBank: "BRI",
    noRekBank: "14301026729509",
    anBank: "Rudi Maulana S",
    simpananPokok: "300000",
    simpananWajib: "600000",
    akunPembayaran: "BSI"
  };

  const params = {
    method: 'addMember',
    sheet: 'simpanan',
    data: data
  };

  const result = newTransaction(params);
  SpreadsheetApp.getUi().alert(result.message);
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

function doGet(e) {
  var formMode = e && e.parameter && e.parameter.form;
  var templateFile = formMode ? 'daftarAnggotaModern' : 'landingPage';
  var title = formMode ? 'Pendaftaran Anggota - KKS Arrahmah' : 'KKS Arrahmah - Keuangan Syariah Martapura';
  
  var template = HtmlService.createTemplateFromFile(templateFile);
  return template.evaluate()
    .setTitle(title)
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
  const params = { method: 'addMembersBulk', data: membersArray };
  const tx = new Transactions(params);
  return tx.postTransactions();
}

function processBatchApproval(indices) {
  const params = { method: 'approvePendingBulk', data: indices };
  const tx = new Transactions(params);
  return tx.postTransactions();
}

function getFormUrl() {
  var baseUrl = ScriptApp.getService().getUrl();
  return baseUrl + (baseUrl.indexOf('?') > -1 ? '&' : '?') + 'form=1';
}