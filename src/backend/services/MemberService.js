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
    const ss = SpreadsheetApp.openById(CONFIG.SS_ID_PENDING).getSheetByName(CONFIG.SHEET_NAMES.PENDING_PENDAFTARAN);
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
   * Mendapat semua NIK yang sudah terdaftar beserta ID Member-nya
   * @returns {Object} Mapping NIK -> idMember
   */
  getExistingNIKs() {
    const ssMember = SpreadsheetApp.openById(CONFIG.SS_ID_MEMBER);
    const sheetMaster = ssMember.getSheetByName(CONFIG.SHEET_NAMES.MASTER_ANGGOTA);
    
    if (sheetMaster.getLastRow() <= 1) return {};

    const data = sheetMaster.getRange(2, 1, sheetMaster.getLastRow() - 1, 4).getValues();
    const mapping = {};
    
    data.forEach(row => {
      const id = String(row[0]); // Kolom A
      const nik = String(row[3]); // Kolom D
      if (nik) mapping[nik] = id;
    });

    return mapping;
  }
};
