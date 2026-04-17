var addTransactions;

function newTransaction(params) {
  addTransactions = new InputTransactions(params);
  return addTransactions.transactionEntries();
}
