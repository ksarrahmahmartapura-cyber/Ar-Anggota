function doPost(e) {
  let data = JSON.parse(e.postData.contents);
  let newPost = new InputTransactions(data);
  let result = newPost.transactionEntries();
  return ContentService.createTextOutput(JSON.stringify(result || { success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}