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
    let lastRowMaster = this.sheetMaster.getLastRow();
    const formattedDate = DateHelper.formatToDMY(new Date());

    membersArray.forEach(member => {
      const idMember = this.createIdMember(lastRowMaster);
      lastRowMaster++; // Increment for next member

      const startMonth = DateHelper.getStartOfMonth(member.tanggalBergabung);
      const totalMonth = member.simpananWajib / this.simpananWajib;
      const lastMonth = new Date(startMonth);
      lastMonth.setMonth(lastMonth.getMonth() + totalMonth - 1);

      // SP Row
      allRows.push([null, formattedDate, "SP", null, idMember, null, null, null, null, "Pendaftaran Anggota (Bulk)", member.simpananPokok, null, this.saldoSimpanan, member.akunPembayaran]);
      // SW Row
      allRows.push([null, formattedDate, "SW", null, idMember, null, DateHelper.formatToDMY(startMonth), totalMonth, DateHelper.formatToDMY(lastMonth), "Pendaftaran Anggota (Bulk)", member.simpananWajib, null, this.saldoSimpanan, member.akunPembayaran]);
      
      // Update member with ID for external API call if needed
      member.idMember = idMember;
      
      // Call external API for each member (as per current design)
      const memberParams = { method: 'addMember', data: member };
      const addTransactions = new Transactions(memberParams);
      addTransactions.postSimpanan();
    });

    if (allRows.length > 0) {
      this.sheetTransactions.getRange(this.sheetTransactions.getLastRow() + 1, 1, allRows.length, 14).setValues(allRows);
    }

    return { success: true, message: `${membersArray.length} anggota berhasil diproses.` };
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