// === НАСТРОЙКИ ===
// ⚠️ ВАЖНО: ниже плейсхолдер. Если у тебя уже есть рабочий Apps Script
// с реальным ID папки — НЕ заменяй его этим файлом, возьми рабочую версию
// из самого Apps Script (Google Таблица → Расширения → Apps Script).
// Этот файл — справочная копия для Claude Code, чтобы видеть структуру.

const DRIVE_FOLDER_ID = 'ВСТАВЬТЕ_ID_ПАПКИ_СЮДА';
const SHEET_NAME = 'Заявки';

function doGet() {
  return ContentService.createTextOutput(
    'Этот адрес принимает заявки с сайта. Форма находится на отдельной странице.'
  );
}

// Точка входа: сюда приходит POST-запрос от формы на сайте
function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    result = submitRegistration(data);
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Дата заявки', 'Имя и фамилия родителя', 'Имя и фамилия ребёнка', 'Ссылка на чек']);
  }
  return sheet;
}

function submitRegistration(data) {
  try {
    if (!data.parentName || !data.childName || !data.fileData) {
      throw new Error('Заполните все обязательные поля и приложите чек.');
    }

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const decoded = Utilities.base64Decode(data.fileData);
    const blob = Utilities.newBlob(decoded, data.mimeType, data.fileName);
    const safeName = (data.parentName + '_' + data.childName + '_' + data.fileName)
      .replace(/[\\\/:*?"<>|]/g, '_');
    blob.setName(safeName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const sheet = getOrCreateSheet();
    sheet.appendRow([
      new Date(),
      data.parentName,
      data.childName,
      file.getUrl()
    ]);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
