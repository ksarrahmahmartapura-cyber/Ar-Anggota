function fnNotification(data) {
  var template = HtmlService.createTemplateFromFile('notification');
  template.data = data;
  var htmlOutput = template.evaluate()
    .setWidth(500)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}