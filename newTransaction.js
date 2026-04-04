var addTransactions;

function newTransaction(params) {
  addTransactions = new Transactions(params);
  addTransactions.postTransactions();
}
