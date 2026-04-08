const DateHelper = {
  formatToDMY(date) {
    return Utilities.formatDate(new Date(date), "GMT+8", "dd/MM/yyyy");
  },
  
  getStartOfMonth(date) {
    const d = new Date(date);
    return new Date(d.getFullYear(), d.getMonth(), 1);
  }
};