class Transactions {
  constructor(data){
    this.dataSend = data;
    this.typeOfSheet = data.sheet;
    this.apiTransactions = 'https://script.google.com/macros/s/AKfycbzuIdsJMATSon8M0HMcPhS5u3oJ2oUAeWAWlziS9ODE9j55Wyd3jMUbxE0rHlYlNJrA/exec';
    this.apiSavingsAccount = 'https://script.google.com/macros/s/AKfycbx24PmBwHdOMjyVHJm3sMJc6OHLkep17Oulc5txyJuKxWCnc2zvH4fdqlsieh4BF0MF/exec';
  }

  async postTransactions(){
    try {      
      let options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(this.dataSend),
      };
      const responseTransactions = await UrlFetchApp.fetch(this.apiTransactions, options);
    this.handleNotification(responseTransactions);
    }
    catch (error) {      
      responseTransactions = { success: false, message: 'Error: ' + error.message };
    }
  }

  async postSimpanan(){
    let options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(this.dataSend),
      };    
    const responseSimpanan = await UrlFetchApp.fetch(this.apiSavingsAccount, options)
    this.handleNotification(responseSimpanan);
  }

  handleNotification(response){
    var template = HtmlService.createTemplateFromFile('notification');
    template.data = response;
    var htmlOutput = template.evaluate()
      .setWidth(500)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
  }
}