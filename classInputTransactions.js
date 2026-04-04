class InputTransactions {
  constructor(data) {
    this.ssTransactions = SpreadsheetApp.openById('1ju5u1Lr-yz54ttw9L4o5AIup7gIaFogHeqyBqwgjLhg').getSheetByName('Transaksi');
    this.ssMemberDB = SpreadsheetApp.openById('1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34').getSheetByName('Master Anggota');
    this.ssPending = SpreadsheetApp.openById('1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34').getSheetByName('Pending_Pendaftaran') || this.createPendingSheet();
    this.simpananWajib = 50000;
    this.saldoSimpanan = '=IF(INDIRECT("C9:C"&ROW())="SW";SUM(FILTER(INDIRECT("K9:K"&ROW());LEFT(INDIRECT("C9:C"&ROW());2)="SW";INDIRECT("E9:E"&ROW())=INDIRECT("E"&ROW())))-SUM(FILTER(INDIRECT("L9:L"&ROW());LEFT(INDIRECT("C9:C"&ROW());2)="SW";INDIRECT("E9:E"&ROW())=INDIRECT("E"&ROW())));SUMIFS(INDIRECT("K9:K"&ROW());INDIRECT("E9:E"&ROW());INDIRECT("E"&ROW());INDIRECT("C9:C"&ROW());INDIRECT("C"&ROW()))-SUMIFS(INDIRECT("L9:L"&ROW());INDIRECT("E9:E"&ROW());INDIRECT("E"&ROW());INDIRECT("C9:C"&ROW());INDIRECT("C"&ROW())))';
    this.params = data;
    this.method = data.method;
    this.data = data.data;
  }

  createPendingSheet() {
    const ss = SpreadsheetApp.openById('1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34');
    const sheet = ss.insertSheet('Pending_Pendaftaran');
    sheet.appendRow([
      'Timestamp', 'Tanggal Pendaftaran', 'Nama Anggota', 'NIK', 'Tempat Lahir', 'Tanggal Lahir',
      'Jenis Kelamin', 'Telpon', 'Email', 'Alamat KTP', 'Alamat Domisili',
      'Pekerjaan', 'Kantor', 'Alamat Kantor', 'Bank', 'No Rek', 'Atas Nama',
      'S. Pokok', 'S. Wajib', 'Metode Bayar'
    ]);
    return sheet;
  }

  createIdMember() {
    const newIdMember = 'KKSA03' + ("0000" + this.ssMemberDB.getLastRow()).slice(-4);
    return newIdMember;
  }

  transactionEntries() {
    if (this.method === 'addMember') {
      const idMember = this.createIdMember();
      const date = new Date(this.data.tanggalBergabung);
      const formattedDate = Utilities.formatDate(date, "GMT+8", "dd/MM/yyyy");
      const startMonth = new Date(date.getFullYear(), date.getMonth(), 1);
      const formattedStartMonth = Utilities.formatDate(startMonth, "GMT+8", "dd/MM/yyyy");
      const totalMonth = this.data.simpananWajib / this.simpananWajib;
      const lastMonth = new Date(date.getFullYear(), date.getMonth() + totalMonth - 1, 1);
      const formattedLastMonth = Utilities.formatDate(lastMonth, "GMT+8", "dd/MM/yyyy");

      this.ssTransactions.appendRow([, formattedDate, "SP", , idMember, , , , , "Pendaftaran Anggota", this.data.simpananPokok, , this.saldoSimpanan, this.data.akunPembayaran]);

      this.ssTransactions.appendRow([, formattedDate, "SW", , idMember, , formattedStartMonth, totalMonth, formattedLastMonth, "Pendaftaran Anggota", this.data.simpananWajib, , this.saldoSimpanan, this.data.akunPembayaran]);

      this.params.data.idMember = idMember;

    } else if (this.method === 'transactionsSimp') {
      let data = this.data;
      let date = new Date();
      let formattedDate = Utilities.formatDate(date, "GMT+8", "dd/MM/yyyy");

      for (const key in data) {
        if (data[key].code === "SW") {

          this.ssTransactions.appendRow([
            , formattedDate, 'SW', , data[key].idMember, , , data[key].debitWajib / this.simpananWajib, data[key].date, 'Simpanan Wajib', data[key].debitWajib, , this.saldoSimpanan, data[key].payment
          ]);

        } else if (data[key].code == "SS") {

          if (data[key].debitSukarela > 0 && data[key].kreditSukarela == 0) {
            this.ssTransactions.appendRow([
              , formattedDate, 'SS', , data[key].idMember, , , , , 'Simpanan Sukarela', data[key].debitSukarela, , this.saldoSimpanan, data[key].payment
            ]);
          } else if (data[key].kreditSukarela > 0 && data[key].debitSukarela == 0) {
            this.ssTransactions.appendRow([
              , formattedDate, 'SS', , data[key].idMember, , , , , 'Penarikan Simpanan Sukarela', , data[key].kreditSukarela, this.saldoSimpanan, data[key].payment
            ]);
          }
        } else if (data[key].code == "SQ") {

          if (data[key].kreditQard == 0) {
            this.ssTransactions.appendRow([
              , formattedDate, 'SQ', , data[key].idMember, , , , , 'Simpanan Qard', data[key].debitQard, , this.saldoSimpanan, data[key].paymentDebitQard
            ]);
            continue;
          } else if (data[key].debitQard == 0) {
            this.ssTransactions.appendRow([
              , formattedDate, 'SQ', , data[key].idMember, , , , , 'Penarikan Simpanan Qard', , data[key].kreditQard, this.saldoSimpanan, data[key].paymentKreditQard
            ]);
            continue;
          }

          this.ssTransactions.appendRow([
            , formattedDate, 'SQ', , data[key].idMember, , , , , 'Simpanan Qard', data[key].debitQard, , this.saldoSimpanan, data[key].paymentDebitQard
          ]);
          this.ssTransactions.appendRow([
            , formattedDate, 'SQ', , data[key].idMember, , , , , 'Penarikan Simpanan Qard', , data[key].kreditQard, this.saldoSimpanan, data[key].paymentKreditQard
          ]);

        } else if (data[key].code == "KA") {
          this.ssTransactions.appendRow([
            , formattedDate, 'SP', , data[key].idMember, , , , , 'Penarikan Simpanan Pokok', , data[key].kaPokok, this.saldoSimpanan, data[key].payment
          ]);
          this.ssTransactions.appendRow([
            , formattedDate, 'SW', , data[key].idMember, , , , , 'Penarikan Simpanan Wajib', , data[key].kaWajib, this.saldoSimpanan, data[key].payment
          ]);

          data[key].kaSukarela > 0 &&
            this.ssTransactions.appendRow([
              , formattedDate, 'SS', , data[key].idMember, , , , , 'Penarikan Simpanan Sukarela', , data[key].kaSukarela, this.saldoSimpanan, data[key].payment
            ]);
          data[key].kaQard > 0 &&
            this.ssTransactions.appendRow([
              , formattedDate, 'SQ', , data[key].idMember, , , , , 'Penarikan Simpanan Qard', , data[key].kaQard, this.saldoSimpanan, data[key].payment
            ]);

        }
      }
    } else if (this.method === 'addPending') {
      this.ssPending.appendRow([
        new Date(),
        this.data.tanggalBergabung, this.data.namaAnggota, this.data.nik, this.data.tempatLahir,
        this.data.tanggalLahir, this.data.jenisKelamin, this.data.telponAnggota, this.data.email,
        this.data.alamatKTP, this.data.alamatTinggal, this.data.jenisPekerjaan, this.data.kantorPekerjaan,
        this.data.alamatKantor, this.data.namaBank, this.data.noRekBank, this.data.anBank,
        this.data.simpananPokok, this.data.simpananWajib, this.data.akunPembayaran
      ]);
      return { success: true, message: 'Data pendaftaran berhasil disimpan untuk ditinjau admin.' };

    } else if (this.method === 'approveMember') {
      // Row index is 1-based, adjusted from frontend (which is 0-based index of the data array)
      const rowIndex = this.params.rowIndex + 2; // +2 because of header and 1-based index
      const rowData = this.ssPending.getRange(rowIndex, 2, 1, 19).getValues()[0];

      // Re-map to match addMember data structure
      this.data = {
        tanggalBergabung: rowData[0], namaAnggota: rowData[1], nik: rowData[2], tempatLahir: rowData[3],
        tanggalLahir: rowData[4], jenisKelamin: rowData[5], telponAnggota: rowData[6], email: rowData[7],
        alamatKTP: rowData[8], alamatTinggal: rowData[9], jenisPekerjaan: rowData[10], kantorPekerjaan: rowData[11],
        alamatKantor: rowData[12], namaBank: rowData[13], noRekBank: rowData[14], anBank: rowData[15],
        simpananPokok: rowData[16], simpananWajib: rowData[17], akunPembayaran: rowData[18]
      };

      this.method = 'addMember'; // Change method to trigger legacy addMember logic
      this.transactionEntries(); // Run the actual registration

      this.ssPending.deleteRow(rowIndex); // Remove from pending
      return { success: true, message: 'Anggota Berhasil Di-approve!' };

    } else if (this.method === 'rejectMember') {
      const rowIndex = this.params.rowIndex + 2;
      this.ssPending.deleteRow(rowIndex);
      return { success: true, message: 'Pendaftaran Ditolak & Dihapus.' };
    }

    addTransactions = new Transactions(this.params);
    addTransactions.postSimpanan();
  }

  // --- OPTIMIZATION: BULK WRITING ---

  addMembersBulk(membersArray) {
    const startRow = this.ssTransactions.getLastRow() + 1;
    const initialLastRowMember = this.ssMemberDB.getLastRow();
    let rowsToAppend = [];

    membersArray.forEach((member, index) => {
      const data = member.data;
      const idMember = 'KKSA03' + ("0000" + (initialLastRowMember + index)).slice(-4);
      const date = new Date(data.tanggalBergabung);
      const formattedDate = Utilities.formatDate(date, "GMT+8", "dd/MM/yyyy");

      // SP Row
      rowsToAppend.push([
        null, formattedDate, "SP", null, idMember, null, null, null, null,
        "Pendaftaran Anggota (Bulk)", data.simpananPokok, null, this.saldoSimpanan, data.akunPembayaran
      ]);

      // SW Row
      const startMonth = new Date(date.getFullYear(), date.getMonth(), 1);
      const formattedStartMonth = Utilities.formatDate(startMonth, "GMT+8", "dd/MM/yyyy");
      const totalMonth = data.simpananWajib / this.simpananWajib;
      const lastMonth = new Date(date.getFullYear(), date.getMonth() + totalMonth - 1, 1);
      const formattedLastMonth = Utilities.formatDate(lastMonth, "GMT+8", "dd/MM/yyyy");

      rowsToAppend.push([
        null, formattedDate, "SW", null, idMember, null, formattedStartMonth, totalMonth,
        formattedLastMonth, "Pendaftaran Anggota (Bulk)", data.simpananWajib, null, this.saldoSimpanan, data.akunPembayaran
      ]);

      // Call external API for each (simulating legacy behavior, but spreadsheet write is batched)
      let singleParams = { method: 'addMember', data: { ...data, idMember: idMember } };
      new Transactions(singleParams).postSimpanan();
    });

    if (rowsToAppend.length > 0) {
      this.ssTransactions.getRange(startRow, 1, rowsToAppend.length, 14).setValues(rowsToAppend);
    }

    return { success: true, count: membersArray.length };
  }

  approvePendingBulk(indices) {
    // Sort indices descending to delete rows from bottom to top safely
    const sortedIndices = indices.sort((a, b) => b - a);
    let membersToProcess = [];

    sortedIndices.forEach(idx => {
      const rowIndex = idx + 2;
      const rowData = this.ssPending.getRange(rowIndex, 2, 1, 19).getValues()[0];
      membersToProcess.push({
        data: {
          tanggalBergabung: rowData[0], namaAnggota: rowData[1], nik: rowData[2], tempatLahir: rowData[3],
          tanggalLahir: rowData[4], jenisKelamin: rowData[5], telponAnggota: rowData[6], email: rowData[7],
          alamatKTP: rowData[8], alamatTinggal: rowData[9], jenisPekerjaan: rowData[10], kantorPekerjaan: rowData[11],
          alamatKantor: rowData[12], namaBank: rowData[13], noRekBank: rowData[14], anBank: rowData[15],
          simpananPokok: rowData[16], simpananWajib: rowData[17], akunPembayaran: rowData[18]
        }
      });
      this.ssPending.deleteRow(rowIndex);
    });

    return this.addMembersBulk(membersToProcess);
  }


}