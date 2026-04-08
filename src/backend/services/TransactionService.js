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