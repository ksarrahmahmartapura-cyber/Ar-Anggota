class sideMenu {
  constructor(name){
    this.ssMemberSavingsDB = SpreadsheetApp.openById('1-czelMtKWcMe5lEw0WyxUKZrCZvP2cRohNECeHudD34').getSheetByName('Anggota');
    this.name = name;
  }

  getData(){
    const lrData = this.ssMemberSavingsDB.getLastRow();
    const data = this.ssMemberSavingsDB.getRange(2,1,lrData,7).getValues();

    const jsonData = data
    .filter(row => row[0] && row[1])
    .reduce((obj, row) => {

      let bulan = new Date(row[4]);
      let validDate = new Date(bulan);
      let formattedDate = isNaN(validDate) ? new Date().toISOString().split('T')[0] : validDate.toISOString().split('T')[0];

      obj[row[0]] = {
        nama:row[1],
        pokok:row[2],
        wajib:row[3],
        bulan:formattedDate,
        sukarela:row[5],
        qard:row[6],
      };
      return obj;
    },{})

    return jsonData;
  }

  goToLink(){
    var template = HtmlService.createTemplateFromFile(this.name);
    template.data = this.getData();
    var htmlOutput = template.evaluate().setWidth(1200).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput,' ');
  }
}

function changeScreen(name) {    
  const link = new sideMenu(name);
  link.goToLink();
}