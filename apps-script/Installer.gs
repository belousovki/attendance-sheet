/* -------------------- v3 alpha installer -------------------- */

const INSTALLER_APP_VERSION = '3.0.0-alpha.1';
const INSTALLER_SCHEMA_VERSION = 3;

function installAttendanceWorkbook() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Откройте Google-таблицу и запустите установщик из связанного Apps Script.');
  }

  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Установить журнал посещаемости?',
    'Установщик создаст и настроит служебные листы в этой таблице. ' +
    'Существующие листы с другими именами не удаляются.',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return;

  const disciplineResponse = ui.prompt(
    'Название дисциплины',
    'Введите название дисциплины. Его можно изменить позже в листе «Настройки».',
    ui.ButtonSet.OK_CANCEL
  );

  if (disciplineResponse.getSelectedButton() !== ui.Button.OK) return;

  const discipline = String(disciplineResponse.getResponseText() || '').trim() || 'Название дисциплины';

  const result = installAttendanceWorkbook_(spreadsheet, {
    discipline: discipline,
    locale: 'ru'
  });

  ui.alert(
    'Установка завершена',
    'Создано листов: ' + result.createdSheets + '.\n' +
    'Версия приложения: ' + INSTALLER_APP_VERSION + '.\n\n' +
    'Дальше:\n' +
    '1. Заполните «Список группы».\n' +
    '2. Выполните «Посещаемость → Синхронизировать состав группы».\n' +
    '3. Разверните Apps Script как Web App и вставьте /exec через меню настройки URL.',
    ui.ButtonSet.OK
  );
}

function installAttendanceWorkbook_(spreadsheet, options) {
  options = options || {};
  const created = [];

  const students = ensureInstallerSheet_(spreadsheet, SHEETS.STUDENTS, created);
  const journal = ensureInstallerSheet_(spreadsheet, SHEETS.JOURNAL, created);
  const lessons = ensureInstallerSheet_(spreadsheet, SHEETS.LESSONS, created);
  const types = ensureInstallerSheet_(spreadsheet, SHEETS.TYPES, created);
  const settings = ensureInstallerSheet_(spreadsheet, SHEETS.SETTINGS, created);
  const marks = ensureInstallerSheet_(spreadsheet, SHEETS.MARKS, created);

  setupStudentsSheet_(students);
  setupJournalSheet_(journal, String(options.discipline || 'Название дисциплины'));
  setupLessonsSheet_(lessons, types);
  setupLessonTypesSheet_(types);
  setupSettingsSheet_(settings, spreadsheet, options);
  setupMarksSheet_(marks);

  props_().setProperty(PROPS.BOUND_SPREADSHEET_ID, spreadsheet.getId());
  props_().deleteProperty(PROPS.ACTIVE_LESSON);
  props_().deleteProperty(PROPS.ACTIVE_CHECK);

  bindToContainer_();
  rebuildJournalFromRegistryInSpreadsheet_(spreadsheet);
  cleanupUnusedBlankSheet_(spreadsheet);

  SpreadsheetApp.flush();

  return {
    createdSheets: created.length,
    sheetNames: created,
    appVersion: INSTALLER_APP_VERSION,
    schemaVersion: INSTALLER_SCHEMA_VERSION
  };
}

function ensureInstallerSheet_(spreadsheet, name, created) {
  let sh = spreadsheet.getSheetByName(name);
  if (!sh) {
    sh = spreadsheet.insertSheet(name);
    created.push(name);
  }
  return sh;
}

function resetInstallerSheetArea_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < cols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, rows, cols)
    .clearContent()
    .clearFormat()
    .clearDataValidations()
    .clearNote();
}

function setupStudentsSheet_(sheet) {
  resetInstallerSheetArea_(sheet, 1000, 3);

  sheet.getRange('A1:C1')
    .setValues([['№', 'ФИО', 'Активен']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1)
    .setDataValidation(checkboxRule);

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 360);
  sheet.setColumnWidth(3, 100);
}

function setupJournalSheet_(sheet, discipline) {
  resetInstallerSheetArea_(sheet, 300, 5);

  sheet.getRange('A1:E1')
    .merge()
    .setValue('Журнал посещаемости — ' + discipline)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('left');

  sheet.getRange('A2:E2')
    .setValues([['№', 'ФИО', 'Баллы за посещение', 'Баллы за работу', 'Итого']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('A3:E3').clearContent().setBackground('#EAEAEA');
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(2);
  sheet.setColumnWidth(1, 64);
  sheet.setColumnWidth(2, 360);
  sheet.setColumnWidths(3, 3, 145);
}

function setupLessonsSheet_(sheet, typesSheet) {
  resetInstallerSheetArea_(sheet, 300, 14);

  const headers = [[
    'ID занятия', 'Дата', 'Формат', 'Тип занятия', '№ по типу', 'Тема',
    'Статус', 'Начало', 'Завершение', 'Колонка Пос.', 'Колонка Оц.',
    'Проверок', 'Заметка', 'Создано'
  ]];

  sheet.getRange(1, 1, 1, 14)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#EAEAEA')
    .setWrap(true);

  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Очно', 'Дистанционно'], true)
    .setAllowInvalid(false)
    .build();

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['active', 'finished'], true)
    .setAllowInvalid(true)
    .build();

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(typesSheet.getRange('A2:A100'), true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(2, LESSON_COL.MODE, sheet.getMaxRows() - 1, 1).setDataValidation(modeRule);
  sheet.getRange(2, LESSON_COL.TYPE, sheet.getMaxRows() - 1, 1).setDataValidation(typeRule);
  sheet.getRange(2, LESSON_COL.STATUS, sheet.getMaxRows() - 1, 1).setDataValidation(statusRule);

  sheet.getRange(1, LESSON_COL.TYPE).setNote(
    'Можно менять после занятия; нумерация, подпись и цвет в Журнале обновляются.'
  );
  sheet.getRange(1, LESSON_COL.TYPE_NUMBER).setNote('Заполняется автоматически.');
  sheet.getRange(1, LESSON_COL.TOPIC).setNote('Необязательное поле; можно редактировать после занятия.');
  sheet.getRange(1, LESSON_COL.NOTES).setNote('Свободная заметка преподавателя.');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 320);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidths(8, 2, 150);
  sheet.setColumnWidths(10, 3, 110);
  sheet.setColumnWidth(13, 320);
  sheet.setColumnWidth(14, 150);
}

function setupLessonTypesSheet_(sheet) {
  resetInstallerSheetArea_(sheet, 100, 4);

  sheet.getRange('A1:D1')
    .setValues([['Тип занятия', 'Цвет (HEX)', 'Нумеровать', 'Активен']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const defaults = [
    ['Лекция', '#DDEBF7', true, true],
    ['Практическое занятие', '#E2F0D9', true, true],
    ['Семинар', '#FFF2CC', true, true],
    ['Лабораторная работа', '#FCE4D6', true, true],
    ['Другое', '#EDEDED', false, true]
  ];

  sheet.getRange(2, 1, defaults.length, 4).setValues(defaults);

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 2)
    .setDataValidation(checkboxRule);

  defaults.forEach((row, idx) => {
    sheet.getRange(idx + 2, 1, 1, 2).setBackground(row[1]);
  });

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidths(3, 2, 120);
}

function setupSettingsSheet_(sheet, spreadsheet, options) {
  resetInstallerSheetArea_(sheet, 100, 3);

  sheet.getRange('A1:C1')
    .setValues([['Ключ', 'Значение', 'Описание']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const rows = [
    ['discipline', String(options.discipline || 'Название дисциплины'), 'Название дисциплины'],
    ['locale', String(options.locale || 'ru'), 'Язык интерфейса: ru / en (подготовка к интернационализации)'],
    ['app_version', INSTALLER_APP_VERSION, 'Версия приложения'],
    ['schema_version', INSTALLER_SCHEMA_VERSION, 'Версия структуры данных'],
    ['regular_code_seconds', 60, 'Время действия обычного кода, секунд'],
    ['regular_code_cycles', 2, 'Пока не используется: коды сменяются непрерывно до ручной остановки'],
    ['late_code_seconds', 60, 'Время действия кода опоздания, секунд'],
    ['late_code_cycles', 1, 'Пока не используется: коды сменяются непрерывно до ручной остановки'],
    ['attendance_present_points', 5, 'Баллы за присутствие'],
    ['attendance_absent_points', 0, 'Баллы за подтверждённое отсутствие'],
    ['auto_absent_on_finish', true, 'При завершении занятия выставлять балл отсутствия неотметившимся'],
    ['late_points', 5, 'Баллы при отметке в режиме опоздания'],
    ['code_digits', 4, 'Количество цифр в коде'],
    ['student_search_min_chars', 3, 'Минимум символов для поиска студента'],
    ['remember_student_in_browser', true, 'Запоминать выбранного студента в браузере'],
    ['schedule_days', '', 'Пока не используется'],
    ['first_lesson_date', '', 'Пока не используется'],
    ['web_app_url', '', 'Точный URL рабочего развертывания /exec'],
    ['student_url', '', 'Прямая ссылка на страницу студента; строится из web_app_url'],
    ['display_url', '', 'Прямая ссылка на экран кода; строится из web_app_url'],
    ['sheet_url', spreadsheet.getUrl(), 'Прямая ссылка на этот журнал'],
    ['teacher_url', '', 'Секретная ссылка на мобильный пульт'],
    ['teacher_key', newTeacherKey_(), 'Секретный ключ пульта; создаётся автоматически'],
    ['student_short_url', '', 'Необязательная короткая ссылка на страницу студента'],
    ['display_short_url', '', 'Необязательная короткая ссылка на экран кода'],
    ['sheet_short_url', '', 'Необязательная короткая ссылка на журнал'],
    ['instance_spreadsheet_id', spreadsheet.getId(), 'ID экземпляра таблицы']
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 430);
  sheet.setColumnWidth(3, 560);
}

function setupMarksSheet_(sheet) {
  resetInstallerSheetArea_(sheet, 3000, 12);

  sheet.getRange(1, 1, 1, 12)
    .setValues([[
      'mark_id', 'lesson_id', 'student_no', 'student_name', 'check_type',
      'code_cycle', 'submitted_at', 'attendance_points', 'status',
      'client_token', 'source', 'notes'
    ]])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  sheet.setFrozenRows(1);
  sheet.hideSheet();
}

function cleanupUnusedBlankSheet_(spreadsheet) {
  const expected = new Set([
    SHEETS.STUDENTS,
    SHEETS.JOURNAL,
    SHEETS.LESSONS,
    SHEETS.TYPES,
    SHEETS.SETTINGS,
    SHEETS.MARKS
  ]);

  const sheets = spreadsheet.getSheets();
  if (sheets.length <= expected.size) return;

  sheets.forEach(sh => {
    if (expected.has(sh.getName())) return;
    if (spreadsheet.getSheets().length <= 1) return;

    const isBlank = sh.getLastRow() <= 1 && sh.getLastColumn() <= 1 && sh.getRange('A1').getValue() === '';
    if (isBlank) spreadsheet.deleteSheet(sh);
  });
}

function diagnoseAttendanceInstallation() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lines = [];
  const required = [
    SHEETS.STUDENTS,
    SHEETS.JOURNAL,
    SHEETS.LESSONS,
    SHEETS.TYPES,
    SHEETS.SETTINGS,
    SHEETS.MARKS
  ];

  required.forEach(name => {
    lines.push((spreadsheet.getSheetByName(name) ? '✓ ' : '✗ ') + name);
  });

  const settingsSheet = spreadsheet.getSheetByName(SHEETS.SETTINGS);
  if (settingsSheet) {
    lines.push('');
    lines.push('app_version: ' + String(getSettingValueInSpreadsheet_(spreadsheet, 'app_version') || 'не задана'));
    lines.push('schema_version: ' + String(getSettingValueInSpreadsheet_(spreadsheet, 'schema_version') || 'не задана'));
    lines.push('web_app_url: ' + (String(getSettingValueInSpreadsheet_(spreadsheet, 'web_app_url') || '').trim() ? 'настроен' : 'не настроен'));
  }

  SpreadsheetApp.getUi().alert('Диагностика установки', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
