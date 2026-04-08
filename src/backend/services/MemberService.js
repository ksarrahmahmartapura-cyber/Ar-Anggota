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
   */
  mapRowToMember(rowData) {
    return {
      tanggalBergabung: rowData[0], 
      namaAnggota: rowData[1], 
      nik: rowData[2],
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
  }
};