/**
 * Дозапись строк учёта времени в таблицу «Учет времени по проектам».
 *
 * Установка (один раз):
 *   1. Открыть таблицу → Расширения → Apps Script.
 *   2. Вставить этот файл, задать SHEET_NAME и SHARED_SECRET.
 *   3. Развернуть → Новое развёртывание → Веб-приложение,
 *      «Запуск от имени: я», «Доступ: все».
 *   4. Скопировать URL развёртывания в timesheet/rules.json → "webhook_url".
 *
 * После этого:
 *   python3 timesheet/build_timesheet.py 2026-08-10 --push
 */

var SHEET_NAME = 'Оксана';
var SHARED_SECRET = 'ЗАМЕНИТЕ_НА_СВОЙ_СЕКРЕТ';

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.secret !== SHARED_SECRET) {
      return json({ status: 'error', message: 'неверный секрет' });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return json({ status: 'error', message: 'нет листа ' + SHEET_NAME });
    }

    // Повторный запуск за ту же дату не должен задваивать часы:
    // сначала убираем строки этой даты, потом дописываем новые.
    var removed = removeRowsForDate(sheet, payload.date);

    payload.rows.forEach(function (row) {
      sheet.appendRow([
        payload.date,
        row.hours,
        row.project,
        row.work_type,
        row.comment
      ]);
    });

    return json({
      status: 'ok',
      added: payload.rows.length,
      replaced: removed
    });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
}

/** Удаляет ранее записанные строки за дату, возвращает их количество. */
function removeRowsForDate(sheet, date) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var dates = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  var removed = 0;

  // идём снизу вверх, чтобы удаление не сдвигало ещё не проверенные строки
  for (var i = lastRow - 1; i >= 1; i--) {
    if (dates[i][0] === date) {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }
  return removed;
}

function json(object) {
  return ContentService
    .createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}
