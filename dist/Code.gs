
// ===== Kode.js =====

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
  const processor = new InputTransactions({ method: 'bulk', data: {} });
  return processor.addMembersBulk(membersArray);
}

function processBatchApproval(indices) {
  const processor = new InputTransactions({ method: 'bulk', data: {} });
  return processor.approvePendingBulk(indices);
}

function getFormUrl() {
  var baseUrl = ScriptApp.getService().getUrl();
  return baseUrl + (baseUrl.indexOf('?') > -1 ? '&' : '?') + 'form=1';
}

// ===== changeScreen.js =====

class sideMenu {
  constructor(name){
    this.ssMemberSavingsDB = SpreadsheetApp.openById('1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34').getSheetByName('Anggota');
    this.name = name;
  }

  getData(){
    const lrData = this.ssMemberSavingsDB.getLastRow();
    const data = this.ssMemberSavingsDB.getRange(2,1,lrData,7).getValues();

    const jsonData = data
    .filter(row => row[0] && row[1])
    .reduce((obj, row) => {

      let bulan = new Date(row[4]);
      let validDate = new Date(bulan);
      let formattedDate = isNaN(validDate) ? new Date().toISOString().split('T')[0] : validDate.toISOString().split('T')[0];

      obj[row[0]] = {
        nama:row[1],
        pokok:row[2],
        wajib:row[3],
        bulan:formattedDate,
        sukarela:row[5],
        qard:row[6],
      };
      return obj;
    },{})

    return jsonData;
  }

  goToLink(){
    var template = HtmlService.createTemplateFromFile(this.name);
    template.data = this.getData();
    var htmlOutput = template.evaluate().setWidth(1200).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput,' ');
  }
}

function changeScreen(name) {    
  const link = new sideMenu(name);
  link.goToLink();
}

// ===== classInputTransactions.js =====

class InputTransactions {
  constructor(data) {
    this.simpananWajib = CONFIG.SIMPANAN_WAJIB_NOMINAL;
    this.saldoSimpanan = '=IF(INDIRECT("C9:C"&ROW())="SW";SUM(FILTER(INDIRECT("K9:K"&ROW());LEFT(INDIRECT("C9:C"&ROW());2)="SW";INDIRECT("E9:E"&ROW())=INDIRECT("E"&ROW())))-SUM(FILTER(INDIRECT("L9:L"&ROW());LEFT(INDIRECT("C9:C"&ROW());2)="SW";INDIRECT("E9:E"&ROW())=INDIRECT("E"&ROW())));SUMIFS(INDIRECT("K9:K"&ROW());INDIRECT("E9:E"&ROW());INDIRECT("E"&ROW());INDIRECT("C9:C"&ROW());INDIRECT("C"&ROW()))-SUMIFS(INDIRECT("L9:L"&ROW());INDIRECT("E9:E"&ROW());INDIRECT("E"&ROW());INDIRECT("C9:C"&ROW());INDIRECT("C"&ROW())))';
    this.params = data;
    this.method = data.method;
    this.data = data.data;
  }

  // Lazy loading getters (MEMPERCEPAT EXECUTION TIME)
  get sheetTransactions() {
    return SpreadsheetApp.openById(CONFIG.SS_ID_TRANSACTIONS).getSheetByName(CONFIG.SHEET_NAMES.TRANSAKSI);
  }
  get sheetPending() {
    return SpreadsheetApp.openById(CONFIG.SS_ID_MEMBER).getSheetByName(CONFIG.SHEET_NAMES.PENDING_PENDAFTARAN);
  }
  get sheetMaster() {
    return SpreadsheetApp.openById(CONFIG.SS_ID_MEMBER).getSheetByName(CONFIG.SHEET_NAMES.MASTER_ANGGOTA);
  }

  createIdMember(lastRow) {
    const rowCount = lastRow || this.sheetMaster.getLastRow();
    const newIdMember = 'KKSA03' + ("0000" + rowCount).slice(-4);
    return newIdMember;
  }

  private_processAddMember() {
    const idMember = this.createIdMember();
    const formattedDate = DateHelper.formatToDMY(this.data.tanggalBergabung);
    const startMonth = DateHelper.getStartOfMonth(this.data.tanggalBergabung);
    const totalMonth = this.data.simpananWajib / this.simpananWajib;
    
    const lastMonth = new Date(startMonth);
    lastMonth.setMonth(lastMonth.getMonth() + totalMonth - 1);

    // Save Profile to Master Anggota
    const profileRow = MemberService.prepareMasterRow(idMember, this.data);
    this.sheetMaster.appendRow(profileRow);

    const rows = [
      [null, formattedDate, "SP", null, idMember, null, null, null, null, "Pendaftaran Anggota", this.data.simpananPokok, null, this.saldoSimpanan, this.data.akunPembayaran],
      [null, formattedDate, "SW", null, idMember, null, DateHelper.formatToDMY(startMonth), totalMonth, DateHelper.formatToDMY(lastMonth), "Pendaftaran Anggota", this.data.simpananWajib, null, this.saldoSimpanan, this.data.akunPembayaran]
    ];

    this.sheetTransactions.getRange(this.sheetTransactions.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
    this.params.data.idMember = idMember;
  }

  private_processBulkTransactions() {
    const rows = TransactionService.mapTransactionsToRows(this.data, this.saldoSimpanan);
    if (rows.length > 0) {
      this.sheetTransactions.getRange(this.sheetTransactions.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
    }
  }

  private_addPending() {
    const row = MemberService.preparePendingRow(this.data);
    this.sheetPending.appendRow(row);
    return { success: true, message: 'Data pendaftaran berhasil disimpan.' };
  }

  private_approveMember() {
    const rowIndex = this.params.rowIndex + 2;
    const rowData = this.sheetPending.getRange(rowIndex, 2, 1, 19).getValues()[0];
    
    this.data = MemberService.mapRowToMember(rowData);
    this.private_processAddMember();
    
    this.sheetPending.deleteRow(rowIndex);
    return { success: true, message: 'Anggota Berhasil Di-approve!' };
  }

  private_rejectMember() {
    const rowIndex = this.params.rowIndex + 2;
    this.sheetPending.deleteRow(rowIndex);
    return { success: true, message: 'Pendaftaran Ditolak.' };
  }

  transactionEntries() {
    switch (this.method) {
      case 'addMember':
        this.private_processAddMember();
        break;
      case 'transactionsSimp':
        this.private_processBulkTransactions();
        break;
      case 'addPending':
        return this.private_addPending();
      case 'approveMember':
        return this.private_approveMember();
      case 'rejectMember':
        return this.private_rejectMember();
    }
    
    const addTransactions = new Transactions(this.params);
    addTransactions.postSimpanan();
  }

  addMembersBulk(membersArray) {
    const allRows = [];
    const memberRows = [];
    let lastRowMaster = this.sheetMaster.getLastRow();
    const formattedDate = DateHelper.formatToDMY(new Date());
    
    // Ambil data NIK yang sudah ada untuk validasi duplikat
    const existingNiks = MemberService.getExistingNIKs();
    let skipCount = 0;

    membersArray.forEach(member => {
      // Pastikan data member diambil dari properti .data jika ada (struktur dari frontend)
      const memberData = member.data || member;
      
      // Validasi Duplikat NIK
      if (existingNiks.includes(String(memberData.nik))) {
        skipCount++;
        return; // Skip member ini
      }
      
      const idMember = this.createIdMember(lastRowMaster);
      lastRowMaster++; // Increment for next member

      const startMonth = DateHelper.getStartOfMonth(memberData.tanggalBergabung);
      const totalMonth = memberData.simpananWajib / this.simpananWajib;
      const lastMonth = new Date(startMonth);
      lastMonth.setMonth(lastMonth.getMonth() + totalMonth - 1);

      // Profile Row
      memberRows.push(MemberService.prepareMasterRow(idMember, memberData));

      // SP Row
      allRows.push([null, formattedDate, "SP", null, idMember, null, null, null, null, "Pendaftaran Anggota (Bulk)", memberData.simpananPokok, null, this.saldoSimpanan, memberData.akunPembayaran]);
      // SW Row
      allRows.push([null, formattedDate, "SW", null, idMember, null, DateHelper.formatToDMY(startMonth), totalMonth, DateHelper.formatToDMY(lastMonth), "Pendaftaran Anggota (Bulk)", memberData.simpananWajib, null, this.saldoSimpanan, memberData.akunPembayaran]);
      
      // Update member with ID for external API call
      memberData.idMember = idMember;
      
      // Call external API for each member (passing ONLY the data part)
      const memberParams = { method: 'addMember', data: memberData };
      const addTransactions = new Transactions(memberParams);
      addTransactions.postSimpanan();
    });

    // Save Profiles to Master Anggota in Bulk
    if (memberRows.length > 0) {
      this.sheetMaster.getRange(this.sheetMaster.getLastRow() + 1, 1, memberRows.length, memberRows[0].length).setValues(memberRows);
    }

    // Save Transactions in Bulk
    if (allRows.length > 0) {
      this.sheetTransactions.getRange(this.sheetTransactions.getLastRow() + 1, 1, allRows.length, 14).setValues(allRows);
    }

    let msg = `${membersArray.length - skipCount} anggota berhasil diproses.`;
    if (skipCount > 0) msg += ` (${skipCount} data dilewati karena NIK sudah terdaftar)`;

    return { 
      success: true, 
      message: msg, 
      count: membersArray.length - skipCount,
      skipped: skipCount 
    };
  }

  approvePendingBulk(indices) {
    // Sort indices descending to avoid range shift during deletion
    indices.sort((a, b) => b - a);
    
    indices.forEach(index => {
      const rowIndex = index + 2;
      const rowData = this.sheetPending.getRange(rowIndex, 2, 1, 19).getValues()[0];
      this.data = MemberService.mapRowToMember(rowData);
      
      this.private_processAddMember();
      
      // Final API call for synchronization
      const addTransactions = new Transactions(this.params);
      addTransactions.postSimpanan();
      
      this.sheetPending.deleteRow(rowIndex);
    });

    return { success: true, message: `${indices.length} pendaftaran berhasil di-approve secara massal.` };
  }
}

// ===== classTransactions.js =====

class Transactions {
  constructor(data){
    this.dataSend = data;
    this.typeOfSheet = data.sheet;
    this.apiTransactions = 'https://script.google.com/macros/s/AKfycbzuIdsJMATSon8M0HMcPhS5u3oJ2oUAeWAWlziS9ODE9j55Wyd3jMUbxE0rHlYlNJrA/exec';
    this.apiSavingsAccount = 'https://script.google.com/macros/s/AKfycbx24PmBwHdOMjyVHJm3sMJc6OHLkep17Oulc5txyJuKxWCnc2zvH4fdqlsieh4BF0MF/exec';
  }

  async postTransactions(){
    try {      
      let options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(this.dataSend),
      };
      const responseTransactions = await UrlFetchApp.fetch(this.apiTransactions, options);
    this.handleNotification(responseTransactions);
    }
    catch (error) {      
      responseTransactions = { success: false, message: 'Error: ' + error.message };
    }
  }

  async postSimpanan(){
    let options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(this.dataSend),
      };    
    const responseSimpanan = await UrlFetchApp.fetch(this.apiSavingsAccount, options)
    this.handleNotification(responseSimpanan);
  }

  handleNotification(response){
    var template = HtmlService.createTemplateFromFile('notification');
    template.data = response;
    var htmlOutput = template.evaluate()
      .setWidth(500)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
  }
}

// ===== doPost.js =====

function doPost(e) {
  let data = JSON.parse(e.postData.contents);
  let newPost = new InputTransactions(data);
  newPost.transactionEntries();
  }

// ===== newTransaction.js =====

var addTransactions;

function newTransaction(params) {
  addTransactions = new Transactions(params);
  addTransactions.postTransactions();
}


// ===== MemberService.js =====

const MemberService = {
  /**
   */
  preparePendingRow(data) {
    return [
      new Date(),
      data.tanggalBergabung, data.namaAnggota, data.nik, data.tempatLahir,
      data.tanggalLahir, data.jenisKelamin, data.telponAnggota, data.email,
      data.alamatKTP, data.alamatTinggal, data.jenisPekerjaan, data.kantorPekerjaan,
      data.alamatKantor, data.namaBank, data.noRekBank, data.anBank,
      data.simpananPokok, data.simpananWajib, data.akunPembayaran
    ];
  },

  /**
   * Menyiapkan baris untuk sheet Master Anggota
   */
  prepareMasterRow(idMember, data) {
    return [
      idMember,
      data.tanggalBergabung, data.namaAnggota, data.nik, data.tempatLahir,
      data.tanggalLahir, data.jenisKelamin, data.telponAnggota, data.email,
      data.alamatKTP, data.alamatTinggal, data.jenisPekerjaan, data.kantorPekerjaan,
      data.alamatKantor, data.namaBank, data.noRekBank, data.anBank,
      data.simpananPokok, data.simpananWajib, data.akunPembayaran
    ];
  },

  /**
   */
  mapRowToMember(rowData) {
    return {
      tanggalBergabung: rowData[0],
      namaAnggota: rowData[1],
      nik: rowData[2],
      tempatLahir: rowData[3],
      tanggalLahir: rowData[4],
      jenisKelamin: rowData[5],
      telponAnggota: rowData[6],
      email: rowData[7],
      alamatKTP: rowData[8],
      alamatTinggal: rowData[9],
      jenisPekerjaan: rowData[10],
      kantorPekerjaan: rowData[11],
      alamatKantor: rowData[12],
      namaBank: rowData[13],
      noRekBank: rowData[14],
      anBank: rowData[15],
      simpananPokok: rowData[16],
      simpananWajib: rowData[17],
      akunPembayaran: rowData[18]
    };
  },

  getPendingData() {
    const ss = SpreadsheetApp.openById(CONFIG.SS_ID_MEMBER).getSheetByName(CONFIG.SHEET_NAMES.PENDING_PENDAFTARAN);
    const data = ss.getDataRange().getValues();
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
  },

  /**
   * Mendapat semua NIK yang sudah terdaftar di Master & Pending
   */
  getExistingNIKs() {
    const ssMember = SpreadsheetApp.openById(CONFIG.SS_ID_MEMBER);
    const sheetMaster = ssMember.getSheetByName(CONFIG.SHEET_NAMES.MASTER_ANGGOTA);
    const sheetPending = ssMember.getSheetByName(CONFIG.SHEET_NAMES.PENDING_PENDAFTARAN);

    const masterNiks = sheetMaster.getLastRow() > 1 
      ? sheetMaster.getRange(2, 4, sheetMaster.getLastRow() - 1, 1).getValues().flat() 
      : [];
    
    const pendingNiks = sheetPending.getLastRow() > 1 
      ? sheetPending.getRange(2, 4, sheetPending.getLastRow() - 1, 1).getValues().flat() 
      : [];

    return [...new Set([...masterNiks, ...pendingNiks])].map(String);
  }
};

// ===== TransactionService.js =====

const TransactionService = {
  /**
   */
  mapTransactionsToRows(data, saldoFormula) {
    const formattedDate = DateHelper.formatToDMY(new Date());
    let rows = [];

    for (const key in data) {
      const entry = data[key];
      const common = [null, formattedDate, entry.code, null, entry.idMember, null, null];

      switch (entry.code) {
        case "SW":
          rows.push([...common, entry.debitWajib / CONFIG.SIMPANAN_WAJIB_NOMINAL, entry.date, 'Simpanan Wajib', entry.debitWajib, null, saldoFormula, entry.payment]);
          break;

        case "SS":
          if (entry.debitSukarela > 0) {
            rows.push([...common, null, null, 'Simpanan Sukarela', entry.debitSukarela, null, saldoFormula, entry.payment]);
          } else if (entry.kreditSukarela > 0) {
            rows.push([...common, null, null, 'Penarikan Simpanan Sukarela', null, entry.kreditSukarela, saldoFormula, entry.payment]);
          }
          break;

        case "SQ":
          if (entry.debitQard > 0) {
            rows.push([...common, null, null, 'Simpanan Qard', entry.debitQard, null, saldoFormula, entry.paymentDebitQard]);
          }
          if (entry.kreditQard > 0) {
            rows.push([...common, null, null, 'Penarikan Simpanan Qard', null, entry.kreditQard, saldoFormula, entry.paymentKreditQard]);
          }
          break;

        case "KA":
          this._addKeluarAnggotaRows(rows, entry, formattedDate, saldoFormula);
          break;
      }
    }
    return rows;
  },

  _addKeluarAnggotaRows(rows, entry, date, saldoFormula) {
    const types = [
      { code: 'SP', label: 'Pokok', val: entry.kaPokok },
      { code: 'SW', label: 'Wajib', val: entry.kaWajib },
      { code: 'SS', label: 'Sukarela', val: entry.kaSukarela },
      { code: 'SQ', label: 'Qard', val: entry.kaQard }
    ];

    types.forEach(t => {
      if (t.val > 0) {
        rows.push([null, date, t.code, null, entry.idMember, null, null, null, null, `Penarikan Simpanan ${t.label}`, null, t.val, saldoFormula, entry.payment]);
      }
    });
  }
};

// ===== fnNotification.js =====

function fnNotification(data) {
  var template = HtmlService.createTemplateFromFile('notification');
  template.data = data;
  var htmlOutput = template.evaluate()
    .setWidth(500)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

// ===== Config.js =====

const CONFIG = {
  SS_ID_TRANSACTIONS: '1ju5u1Lr-yz54ttw9L4o5AIup7gIaFogHeqyBqwgjLhg',
  SS_ID_MEMBER: '1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34',
  SHEET_NAMES: {
    TRANSAKSI: 'Transaksi',
    MASTER_ANGGOTA: 'Master Anggota',
    PENDING_PENDAFTARAN: 'Pending_Pendaftaran'
  },
  SIMPANAN_WAJIB_NOMINAL: 50000
};

// ===== DateHelper.js =====

const DateHelper = {
  formatToDMY(date) {
    return Utilities.formatDate(new Date(date), "GMT+8", "dd/MM/yyyy");
  },
  
  getStartOfMonth(date) {
    const d = new Date(date);
    return new Date(d.getFullYear(), d.getMonth(), 1);
  }
};
