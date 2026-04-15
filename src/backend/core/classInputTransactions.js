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

      // Ensure phone number has a leading quote for Sheets
      if (memberData.telponAnggota && !String(memberData.telponAnggota).startsWith("'")) {
        memberData.telponAnggota = "'" + memberData.telponAnggota;
      }

      const startMonth = DateHelper.getStartOfMonth(memberData.tanggalBergabung);
      const totalMonth = memberData.simpananWajib / this.simpananWajib;
      const lastMonth = new Date(startMonth);
      lastMonth.setMonth(lastMonth.getMonth() + totalMonth - 1);

      // Save Profile Row immediately (Consistent with individual form)
      const profileRow = MemberService.prepareMasterRow(idMember, memberData);
      this.sheetMaster.appendRow(profileRow);

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

    // Save Transactions in Bulk
    if (allRows.length > 0) {
      this.sheetTransactions.getRange(this.sheetTransactions.getLastRow() + 1, 1, allRows.length, 14).setValues(allRows);
    }

    // Flush all changes to Spreadsheet
    SpreadsheetApp.flush();

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