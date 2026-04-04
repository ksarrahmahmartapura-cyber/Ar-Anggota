class InputTransactions {
  constructor(data){
    this.ssTransactions = SpreadsheetApp.openById('1ju5u1Lr-yz54ttw9L4o5AIup7gIaFogHeqyBqwgjLhg').getSheetByName('Transaksi');
    this.ssMemberDB = SpreadsheetApp.openById('1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34').getSheetByName('Master Anggota');
    this.simpananWajib = 50000;
    this.saldoSimpanan = '=IF(INDIRECT("C9:C"&ROW())="SW";SUM(FILTER(INDIRECT("K9:K"&ROW());LEFT(INDIRECT("C9:C"&ROW());2)="SW";INDIRECT("E9:E"&ROW())=INDIRECT("E"&ROW())))-SUM(FILTER(INDIRECT("L9:L"&ROW());LEFT(INDIRECT("C9:C"&ROW());2)="SW";INDIRECT("E9:E"&ROW())=INDIRECT("E"&ROW())));SUMIFS(INDIRECT("K9:K"&ROW());INDIRECT("E9:E"&ROW());INDIRECT("E"&ROW());INDIRECT("C9:C"&ROW());INDIRECT("C"&ROW()))-SUMIFS(INDIRECT("L9:L"&ROW());INDIRECT("E9:E"&ROW());INDIRECT("E"&ROW());INDIRECT("C9:C"&ROW());INDIRECT("C"&ROW())))';
    this.params = data;
    this.method = data.method;
    this.data = data.data;
  }

  createIdMember(){
    const newIdMember = 'KKSA03' + ("0000" + this.ssMemberDB.getLastRow()).slice(-4);
    return newIdMember;
  }

  transactionEntries(){
    if(this.method === 'addMember'){
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

    } else if(this.method === 'transactionsSimp'){
        let data = this.data;
        let date = new Date();
        let formattedDate = Utilities.formatDate(date, "GMT+8", "dd/MM/yyyy");

        for(const key in data){
          if(data[key].code === "SW"){

            this.ssTransactions.appendRow([
              ,formattedDate, 'SW',,data[key].idMember,,,data[key].debitWajib / this.simpananWajib, data[key].date, 'Simpanan Wajib', data[key].debitWajib,,this.saldoSimpanan,data[key].payment
            ]);
          
          } else if(data[key].code == "SS"){
          
            if(data[key].debitSukarela > 0 && data[key].kreditSukarela == 0){
              this.ssTransactions.appendRow([
                ,formattedDate, 'SS',,data[key].idMember,,,,, 'Simpanan Sukarela', data[key].debitSukarela,,this.saldoSimpanan,data[key].payment
              ]);
            } else if(data[key].kreditSukarela > 0 && data[key].debitSukarela == 0) {
              this.ssTransactions.appendRow([
                ,formattedDate, 'SS',,data[key].idMember,,,,, 'Penarikan Simpanan Sukarela',,data[key].kreditSukarela,this.saldoSimpanan,data[key].payment
              ]);
            } 
          } else if(data[key].code == "SQ"){

            if(data[key].kreditQard == 0){
              this.ssTransactions.appendRow([
                ,formattedDate, 'SQ',,data[key].idMember,,,,, 'Simpanan Qard',data[key].debitQard,,this.saldoSimpanan,data[key].paymentDebitQard
              ]);
              continue;
            } else if(data[key].debitQard == 0) {
              this.ssTransactions.appendRow([
                ,formattedDate, 'SQ',,data[key].idMember,,,,, 'Penarikan Simpanan Qard',,data[key].kreditQard,this.saldoSimpanan,data[key].paymentKreditQard
              ]);
              continue;
            }

            this.ssTransactions.appendRow([
              ,formattedDate, 'SQ',,data[key].idMember,,,,, 'Simpanan Qard',data[key].debitQard,,this.saldoSimpanan,data[key].paymentDebitQard
            ]);
            this.ssTransactions.appendRow([
              ,formattedDate, 'SQ',,data[key].idMember,,,,, 'Penarikan Simpanan Qard',,data[key].kreditQard,this.saldoSimpanan,data[key].paymentKreditQard
            ]);
            
          } else if(data[key].code == "KA"){
            this.ssTransactions.appendRow([
              ,formattedDate, 'SP',,data[key].idMember,,,,, 'Penarikan Simpanan Pokok',,data[key].kaPokok,this.saldoSimpanan,data[key].payment
            ]);
            this.ssTransactions.appendRow([
              ,formattedDate, 'SW',,data[key].idMember,,,,, 'Penarikan Simpanan Wajib',,data[key].kaWajib,this.saldoSimpanan,data[key].payment
            ]);

            data[key].kaSukarela > 0 &&
            this.ssTransactions.appendRow([
              ,formattedDate, 'SS',,data[key].idMember,,,,, 'Penarikan Simpanan Sukarela',,data[key].kaSukarela,this.saldoSimpanan,data[key].payment
            ]);
            data[key].kaQard > 0 &&
            this.ssTransactions.appendRow([
              ,formattedDate, 'SQ',,data[key].idMember,,,,, 'Penarikan Simpanan Qard',,data[key].kaQard,this.saldoSimpanan,data[key].payment
            ]);
            
          }
        }
    }
    addTransactions = new Transactions(this.params);
    addTransactions.postSimpanan();
  }

  
}