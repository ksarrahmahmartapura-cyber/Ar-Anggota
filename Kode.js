function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Arrahmah')
    .addItem('Home', 'fnHome')
    .addToUi();
}

function fnHome(){
  var template = HtmlService.createTemplateFromFile('home');
  var htmlOutput = template.evaluate().setWidth(1200).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput,' ');
}