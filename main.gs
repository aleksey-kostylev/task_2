// Создаем меню для более удобного управления
function initMenu(){
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu('Мои Макросы');
  menu.addItem('Загрузить картинки', 'setImage');
  menu.addToUi();

}

// Делаем подгрузку меню при запуске файла
function onOpen(){
  
  initMenu()

}

// Создаем главную функцию для парсинга VK API
function vkAPI_community_parse(community_id, num_posts){
  var token = vkToken(); // импортируем VK API Token из другого файла при помощи функции
  var url = `https:\/\/api.vk.com/method/wall.get?owner_id=-${community_id}&count=${num_posts}&filter=all&v=5.95&access_token=${token}`;
  var response = UrlFetchApp.fetch(url);
  var content = response.getContentText();
  var data = JSON.parse(content)["response"];
  
  var arr = [['ID', 'Date', 'Text', 'URL']]; // creating table header
  
  for(var row = 0; row < num_posts; row++){
    if (data["items"][row].hasOwnProperty('attachments')){
      if(data["items"][row]["attachments"][0]['type'] == 'photo'){
        arr = arr.concat([[data["items"][row]["id"], new Date(data["items"][row]["date"]*1000),
                           data["items"][row]["text"], data["items"][row]["attachments"][0]['photo']['sizes'][2]['url']]])
      }
    }
  }
  return arr
}


// Функция для загрузки картинок (находится в меню "Мои макросы" на кнопке "Загрузить картинки")
function setImage(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('E5').setFormula('=IMAGE(D5)')
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var fillDownRange = sheet.getRange(5, lastColumn, lastRow-4)
  sheet.getRange('E5').copyTo(fillDownRange)
}