function doPost(e) {
  let data = JSON.parse(e.postData.contents);
  let newPost = new InputTransactions(data);
  newPost.transactionEntries();
  }