/**
 * Google Apps Script — Форма записи на курс "Эффективная команда"
 *
 * ИНСТРУКЦИЯ:
 * 1. Откройте sheets.google.com → создайте пустую таблицу
 * 2. В первой строке напишите: Дата | Имя | Фамилия | E-mail
 * 3. Скопируйте ID таблицы из адресной строки (между /d/ и /edit)
 * 4. Вставьте ID в переменную SPREADSHEET_ID ниже
 * 5. Расширения → Apps Script → вставьте этот код → Сохраните
 * 6. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone → Deploy
 * 7. Скопируйте URL и вставьте в index.html в строку const SCRIPT_URL='...'
 */

// ============ НАСТРОЙКА ============
const SPREADSHEET_ID = 'ВСТАВЬТЕ_СЮДА_ID_ВАШЕЙ_ТАБЛИЦЫ';
const SHEET_NAME = 'Заявки';
// ====================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);

    // Если листа нет — создаём с заголовками
    if (!sheet) {
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_NAME);
      sheet.appendRow(['Дата и время', 'Имя', 'Фамилия', 'E-mail', 'IP адрес', 'Источник']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#2E3443').setFontColor('#BC9C69');
      sheet.setColumnWidth(1, 180);
      sheet.setColumnWidth(2, 150);
      sheet.setColumnWidth(3, 150);
      sheet.setColumnWidth(4, 220);
      sheet.setColumnWidth(5, 140);
      sheet.setColumnWidth(6, 160);
    }

    // Парсим данные
    var data = e.parameter;
    var now = new Date();
    var timeZone = Session.getScriptTimeZone();
    var formattedDate = Utilities.formatDate(now, timeZone, 'dd.MM.yyyy HH:mm:ss');

    var firstName = (data.firstName || '').trim();
    var lastName  = (data.lastName  || '').trim();
    var email     = (data.email     || '').trim();
    var ip        = e.clientAddress || 'не определён';
    var source    = data.source || 'лендинг';

    // Валидация
    if (!firstName || !lastName || !email) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Обязательные поля не заполнены' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (!email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Некорректный e-mail' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Проверка на дубликат по e-mail
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var emails = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
      for (var i = 0; i < emails.length; i++) {
        if (emails[i][0].toString().toLowerCase() === email.toLowerCase()) {
          return ContentService.createTextOutput(
            JSON.stringify({ success: false, error: 'Этот e-mail уже зарегистрирован' })
          ).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // Записываем строку
    sheet.appendRow([formattedDate, firstName, lastName, email, ip, source]);

    // Отправляем уведомление на почту (опционально)
    sendNotification(formattedDate, firstName, lastName, email);

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, message: 'Заявка успешно принята' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Отправка email-уведомления о новой заявке
 * Замените YOUR_EMAIL на вашу почту
 */
function sendNotification(date, firstName, lastName, email) {
  var YOUR_EMAIL = 'ВСТАВЬТЕ_ВАШУ_ПОЧТУ'; // например: ivan@gmail.com

  if (!YOUR_EMAIL || YOUR_EMAIL === 'ВСТАВЬТЕ_ВАШУ_ПОЧТУ') return;

  var subject = 'Новая заявка на курс — ' + firstName + ' ' + lastName;
  var body =
    'НОВАЯ ЗАЯВКА НА КУРС\n\n' +
    'Дата: ' + date + '\n' +
    'Имя: ' + firstName + '\n' +
    'Фамилия: ' + lastName + '\n' +
    'E-mail: ' + email + '\n\n' +
    'Данные автоматически добавлены в Google Таблицу.';

  MailApp.sendEmail(YOUR_EMAIL, subject, body);
}


/**
 * Открывает веб-страницу с формой (опционально)
 * Можно использовать как тестовую страницу
 */
function doGet(e) {
  var html =
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>Тест формы</title>' +
    '<style>body{font-family:sans-serif;max-width:400px;margin:40px auto;padding:20px}' +
    'input{display:block;width:100%;padding:10px;margin:8px 0;border:1px solid #ccc;border-radius:6px}' +
    'button{background:#BC9C69;color:#fff;border:none;padding:12px 24px;border-radius:24px;cursor:pointer;font-size:16px;width:100%;margin-top:12px}' +
    '</style></head><body>' +
    '<h2>Тест: Запись на курс</h2>' +
    '<form method="POST" action="' + ScriptApp.getService().getUrl() + '">' +
    '<input name="firstName" placeholder="Имя" required>' +
    '<input name="lastName" placeholder="Фамилия" required>' +
    '<input name="email" type="email" placeholder="E-mail" required>' +
    '<input type="hidden" name="source" value="тест">' +
    '<button type="submit">Отправить</button>' +
    '</form></body></html>';

  return HtmlService.createHtmlOutput(html);
}


/**
 * Утилита: удалить дубликаты по e-mail (запустить один раз вручную)
 */
function removeDuplicates() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var seen = {};
  var rowsToRemove = [];

  for (var i = 1; i < data.length; i++) {
    var email = data[i][3].toString().toLowerCase();
    if (seen[email]) {
      rowsToRemove.push(i + 1); // +1 потому что нумерация с 1
    } else {
      seen[email] = true;
    }
  }

  // Удаляем с конца, чтобы не сбивались индексы
  for (var j = rowsToRemove.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToRemove[j]);
  }

  LoggerApp.log('Удалено дубликатов: ' + rowsToRemove.length);
}


/**
 * Утилита: экспорт всех заявок в CSV (для скачивания)
 */
function exportToCSV() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var csv = '';

  for (var i = 0; i < data.length; i++) {
    csv += data[i].join(';') + '\n';
  }

  var blob = Utilities.newBlob(csv, 'text/csv', 'zayavki.csv');
  var file = DriveApp.createFile(blob);
  LoggerApp.log('CSV создан: ' + file.getUrl());
}
