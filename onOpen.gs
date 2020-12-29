function onOpen() {    
  var ui = SpreadsheetApp.getUi();   
  ui.createMenu("Меню работника")
  .addItem("Копирование данных из Google Calendar", "main")
  .addSeparator()
  .addItem("Подведение итогов месяца", "secondary") 
  .addToUi();  
}
