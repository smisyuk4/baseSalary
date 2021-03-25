function main() {        
  var list = SpreadsheetApp.openById("*****").getSheetByName("Данные из календаря");       
  var calendar = CalendarApp.getCalendarById("******"); 
  //проверка в каком режиме выполнять загрузку событий из календаря
  var modeLoadingEvents = list.getRange(1, 2).isChecked(); 
  var oneDay;
  
  if (modeLoadingEvents == true){
    //автоматический режим
     oneDay = new Date();      
  }
  else{
    //ручной режим - поиск даты для загрузки данных   
     oneDay = list.getRange(1, 4).getValue();   
  }    

  pullEvents(oneDay, calendar, list);
}

function secondary () {    
  info = { 
    list : SpreadsheetApp.openById("*********").getSheetByName("Данные из календаря"),
    analiticList : SpreadsheetApp.openById("*********").getSheetByName("Статистика")
  };    
  
  createRangePerMonth (info);
  sumTimePerMonth (info);
  info.list.getRange(info.lastRow, 2).setValue(converterTime (info.minMonth));

  var formulaSum = "=SUM(" + info.rangeSalary + ")"; //"=SUM(C3:C121)"  
  info.list.getRange(info.lastRow, 3).setFormula(formulaSum);  

  questionForTransfer(info);
}

function pullEvents (oneDay, calendar, list){
  oneDay.setHours(oneDay.getHours() + 7);  // +7 часов
  Logger.log(oneDay);

  costHour = list.getRange(1, 6).getValue();   
  costMin = costHour/60;

  //отлавливание ошибок из-за отсутствия событий в календарном дне
  try{   
    var events = calendar.getEventsForDay(oneDay);    
    //поиск последней строки в таблице
    var lastRow = list.getLastRow();    

    //проверка на последний день месяца
    finishMonth (oneDay, list, lastRow);

    for (var i=0; i<events.length; i++){
      var startTime = events[i].getStartTime();
      var endTime = events[i].getEndTime();
      var min = (endTime - startTime)/60000;

      list.getRange(lastRow+1, 1).setValue(events[i].getStartTime()); //записывает дату рабочего дня

      //вычитает из общих часов - перерыв 30мин
      if (min > 480) {
        min = min - 30;
        list.getRange(lastRow+1, 4).setValue("Перерыв 30 мин учтён");    
      }    
           
      list.getRange(lastRow+1, 2).setValue(converterTime (min));  //переводит минуты в часы и записывает в таблицу      
      list.getRange(lastRow+1, 3).setValue(Math.round(min * costMin)); //считает зарплату за день
    }
  } //конец try
  catch (e){
    Logger.log("В календаре нет событий для загрузки");
    return;
  }
}

function finishMonth (oneDay, list, lastRow){    
  //выяснение дня недели, месяца, года 
  var todayDate = oneDay.getDate();   
  var numMonth = oneDay.getMonth() + 1; 
  var fullYear = oneDay.getFullYear(); 
  
  /*
  0 - Воскресенье, 1 - Понедельник, ..., 6 - Суббота    
  0 - Январь, 1 - Февраль, ..., 11 - Декабрь
  */
   
  //учитывает высокосность года  
  var lastDayMonth = 28 + ((numMonth + Math.floor(numMonth / 8)) % 2) + 2 % numMonth + 
    Math.floor((1 + (1 - (fullYear % 4 + 2) % (fullYear % 4 + 1)) * 
      ((fullYear % 100 + 2) % (fullYear % 100 + 1)) + (1 - (fullYear % 400 + 2) % (fullYear % 400 + 1))) / numMonth) + 
        Math.floor(1/numMonth) - Math.floor(((1 - (fullYear % 4 + 2) % (fullYear % 4 + 1)) * 
          ((fullYear % 100 + 2) % (fullYear % 100 + 1)) + (1 - (fullYear % 400 + 2) % (fullYear % 400 + 1)))/numMonth);    
    
  if (todayDate == lastDayMonth){
    list.getRange(lastRow+2, 4).setValue("Транспорт");    
    list.getRange(lastRow+3, 4).setValue("Налоги и отчисления");    
    list.getRange(lastRow+4, 1, 1, 6).setBackground("#ffff00");  
    list.getRange(lastRow+4, 5).setValue("<----- Значения за месяц");    
  }    
}

function converterTime (min){
  var outputHour = Math.floor(min / 60);
  var outputMin = min % 60;
  return outputHour + "ч " + (outputMin != 0 ? outputMin + "мин" : "");
}

function createRangePerMonth (info){
  //поиск нижней желтой строки (<----- Значения за месяц)
   var bottomPoint = info.list.getLastRow();     
    while(info.list.getRange(bottomPoint, 2).getBackground() !== "#ffff00"){
      bottomPoint--;    
    }
  
  //поиск верхней желтой строки (<----- Значения за месяц)
    var topPoint = bottomPoint;      
      do {
      topPoint--;    
      }while(info.list.getRange(topPoint, 2).getBackground() !== "#ffff00")

  //формирование имени ячеек диапазона   
    var sumRangeTop = info.list.getRange(topPoint + 1, 2).getA1Notation();
    var sumRangeBottom = info.list.getRange(bottomPoint - 1, 2).getA1Notation();    
        rangeHour = sumRangeTop + ":" + sumRangeBottom; //B3:B121 
        Logger.log(rangeHour); 
        info.rangeHour = rangeHour;        

    var sumRangeTop = info.list.getRange(topPoint + 1, 3).getA1Notation();
    var sumRangeBottom = info.list.getRange(bottomPoint - 1, 3).getA1Notation();    
        rangeSalary = sumRangeTop + ":" + sumRangeBottom; //С3:С121 
        Logger.log(rangeSalary); 
        info.rangeSalary = rangeSalary;  
        info.lastRow = bottomPoint;
  return info;
}

function sumTimePerMonth (info){  
  var timePerDay = info.list.getRange(info.rangeHour).getValues();  
  var regex = /\d+/;

  var arrayData = timePerDay.flat();
  Logger.log(arrayData);

  var minDay = 0;
      minMonth = 0;
  
  //сделать конверсию в минуты
  for (var i=0; i<arrayData.length; i++){
    var arr = arrayData[i].split(" ");
    var hourStr = arr[0]; //ч
    var hourNum = (arr[0].match(regex))*1;        
        minStr = arr[1]; //мин
        try{
          minNum = (arr[1].match(regex))*1;   
        } 
        catch (e){
          Logger.log(e);
        }
        minDay = hourNum * 60 + minNum;
        minMonth += minDay;    
  }
  Logger.log(minMonth);  
  info.minMonth = minMonth;
  return info;
}

function questionForTransfer(info){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Сделать копирование \"итогов месяца\" в Статистику?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    transferDataToAnaliticList(info);
    Logger.log('The user clicked "Yes."');
  } else {   
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
  }  
}

function transferDataToAnaliticList(info) {  
  //поиск даты в list
  var date;   
  var i=0;
  do{
    i++;
    date = info.list.getRange(info.lastRow-i, 1).getValue(); //если она пустая, то взять выше на одну строку
  }while(date == '')
    
  Logger.log(date);
  //var arrayDate = date.split(" ");   
  //var newArrayDate = arrayDate[0]+ " " +arrayDate[1]+ " " + arrayDate[2]+ " " + arrayDate[3]; //30 ноября 2019 г.
   
  //поиск ячейки в analiticList и запись формулы  
  var lastRowAnaliticList = info.analiticList.getLastRow();  
  info.analiticList.getRange(lastRowAnaliticList+1, 1).setValue(date);  
  
  for (var i=0; i<3; i++){
  //поиск ячейки в list и её названия, сделать формулу
  var count = info.list.getRange(info.lastRow, 2+i).getA1Notation();  
  var formulaTransfer = "='Данные из календаря'!" + count; //='Данные из календаря'!B11   
  info.analiticList.getRange(lastRowAnaliticList+1, 2+i).setValue(formulaTransfer);
  }
  
}
