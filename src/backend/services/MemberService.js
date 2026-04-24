const MemberService = {
  /**
   */
  preparePendingRow(data) {
    return [
      new Date(),
      data.tanggalBergabung, data.namaAnggota, data.nik, data.tempatLahir,
      data.tanggalLahir, data.jenisKelamin, data.telponAnggota, data.email,
      data.alamatKTP, data.alamatTinggal, 
      data.keluargaSerumah, data.hubunganKeluargaSerumah, data.telponKeluargaSerumah,
      data.keluargaTidakSerumah, data.telponKeluargaTidakSerumah,
      data.jenisPekerjaan, data.kantorPekerjaan, data.alamatKantor, 
      data.namaBank, data.noRekBank, data.anBank,
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
      data.alamatKTP, data.alamatTinggal, 
      data.keluargaSerumah, data.hubunganKeluargaSerumah, data.telponKeluargaSerumah,
      data.keluargaTidakSerumah, data.telponKeluargaTidakSerumah,
      data.jenisPekerjaan, data.kantorPekerjaan, data.alamatKantor, 
      data.namaBank, data.noRekBank, data.anBank,
      data.simpananPokok, data.simpananWajib, data.akunPembayaran
    ];
  },

  /**
   */
  mapRowToMember(rowData) {
    return {
      tanggalBergabung: rowData[1],
      namaAnggota: rowData[2],
      nik: rowData[3],
      tempatLahir: rowData[4],
      tanggalLahir: rowData[5],
      jenisKelamin: rowData[6],
      telponAnggota: rowData[7],
      email: rowData[8],
      alamatKTP: rowData[9],
      alamatTinggal: rowData[10],
      keluargaSerumah: rowData[11],
      hubunganKeluargaSerumah: rowData[12],
      telponKeluargaSerumah: rowData[13],
      keluargaTidakSerumah: rowData[14],
      telponKeluargaTidakSerumah: rowData[15],
      jenisPekerjaan: rowData[16],
      kantorPekerjaan: rowData[17],
      alamatKantor: rowData[18],
      namaBank: rowData[19],
      noRekBank: rowData[20],
      anBank: rowData[21],
      simpananPokok: rowData[22],
      simpananWajib: rowData[23],
      akunPembayaran: rowData[24]
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
