/* -------------------- v3.0 alpha.6 installer -------------------- */

const INSTALLER_APP_VERSION = '3.0.0-alpha.6';
const INSTALLER_SCHEMA_VERSION = 3;

function isAttendanceInstallationComplete_(spreadsheet) {
  const settings = spreadsheet.getSheetByName(SHEETS.SETTINGS);
  if (!settings) return false;

  const schema = String(
    getSettingValueInSpreadsheet_(spreadsheet, 'schema_version') || ''
  ).trim();

  const status = String(
    getSettingValueInSpreadsheet_(spreadsheet, 'install_status') || ''
  ).trim();

  return schema === String(INSTALLER_SCHEMA_VERSION) && status === 'ready';
}

function syncVersionMetadata_(spreadsheet) {
  const settings = spreadsheet.getSheetByName(SHEETS.SETTINGS);
  if (!settings) return;

  setSettingValueInSpreadsheet_(
    spreadsheet,
    'app_version',
    INSTALLER_APP_VERSION,
    'Версия приложения'
  );

  setSettingValueInSpreadsheet_(
    spreadsheet,
    'code_version',
    INSTALLER_APP_VERSION,
    'Версия установленного кода'
  );

  const schema = String(
    getSettingValueInSpreadsheet_(spreadsheet, 'schema_version') || ''
  ).trim();

  if (!schema) {
    setSettingValueInSpreadsheet_(
      spreadsheet,
      'schema_version',
      INSTALLER_SCHEMA_VERSION,
      'Версия структуры данных'
    );
  }
}

function migrateLegacyLessonStatuses_(spreadsheet) {
  const sh = spreadsheet.getSheetByName(SHEETS.LESSONS);
  if (!sh) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const range = sh.getRange(2, LESSON_COL.STATUS, lastRow - 1, 1);
  const values = range.getValues();
  let changed = 0;

  values.forEach(row => {
    if (String(row[0] || '').trim().toLowerCase() === 'finished') {
      row[0] = 'closed';
      changed++;
    }
  });

  if (changed) range.setValues(values);
  return changed;
}

function applyLessonDateFormats_(spreadsheet) {
  const sh = spreadsheet.getSheetByName(SHEETS.LESSONS);
  if (!sh) return;

  const rows = Math.max(sh.getMaxRows() - 1, 1);

  sh.getRange(2, LESSON_COL.DATE, rows, 1)
    .setNumberFormat('dd.MM.yyyy');

  [LESSON_COL.STARTED, LESSON_COL.ENDED, LESSON_COL.CREATED]
    .forEach(col => {
      sh.getRange(2, col, rows, 1)
        .setNumberFormat('dd.MM.yyyy HH:mm:ss');
    });
}

function applyLessonValidationRules_(spreadsheet) {
  const sh = spreadsheet.getSheetByName(SHEETS.LESSONS);
  const types = spreadsheet.getSheetByName(SHEETS.TYPES);
  if (!sh || !types) return;

  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Очно', 'Дистанционно'], true)
    .setAllowInvalid(false)
    .build();

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['active', 'closed'], true)
    .setAllowInvalid(false)
    .build();

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(types.getRange('A2:A100'), true)
    .setAllowInvalid(true)
    .build();

  const rows = Math.max(sh.getMaxRows() - 1, 1);

  sh.getRange(2, LESSON_COL.MODE, rows, 1).setDataValidation(modeRule);
  sh.getRange(2, LESSON_COL.TYPE, rows, 1).setDataValidation(typeRule);
  sh.getRange(2, LESSON_COL.STATUS, rows, 1).setDataValidation(statusRule);
}

function countInstallerStudents_(spreadsheet, activeOnly) {
  const sh = spreadsheet.getSheetByName(SHEETS.STUDENTS);
  if (!sh || sh.getLastRow() < 2) return 0;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();

  return values.filter(r => {
    const hasStudent = r[0] !== '' && r[0] != null && String(r[1] || '').trim();
    if (!hasStudent) return false;
    if (!activeOnly) return true;
    return settingBool_(r[2], true);
  }).length;
}

function getInstallationHealth_(spreadsheet) {
  const required = [
    SHEETS.STUDENTS,
    SHEETS.JOURNAL,
    SHEETS.LESSONS,
    SHEETS.TYPES,
    SHEETS.SETTINGS,
    SHEETS.MARKS
  ];

  const missing = required.filter(name => !spreadsheet.getSheetByName(name));
  const settings = spreadsheet.getSheetByName(SHEETS.SETTINGS);

  const storedVersion = settings
    ? String(getSettingValueInSpreadsheet_(spreadsheet, 'app_version') || '').trim()
    : '';

  const storedSchema = settings
    ? String(getSettingValueInSpreadsheet_(spreadsheet, 'schema_version') || '').trim()
    : '';

  const installStatus = settings
    ? String(getSettingValueInSpreadsheet_(spreadsheet, 'install_status') || '').trim()
    : '';

  const webAppUrl = settings
    ? String(getSettingValueInSpreadsheet_(spreadsheet, 'web_app_url') || '').trim()
    : '';

  const lessonSheet = spreadsheet.getSheetByName(SHEETS.LESSONS);
  let lessons = 0;
  let activeLessons = 0;

  if (lessonSheet && lessonSheet.getLastRow() >= 2) {
    const rows = lessonSheet.getRange(
      2,
      1,
      lessonSheet.getLastRow() - 1,
      LESSON_COL.STATUS
    ).getValues();

    rows.forEach(r => {
      const id = String(r[LESSON_COL.ID - 1] || '').trim();
      if (!id) return;
      lessons++;
      if (String(r[LESSON_COL.STATUS - 1] || '').trim() === 'active') {
        activeLessons++;
      }
    });
  }

  return {
    missing,
    storedVersion,
    storedSchema,
    installStatus,
    webAppConfigured: !!webAppUrl,
    students: countInstallerStudents_(spreadsheet, false),
    activeStudents: countInstallerStudents_(spreadsheet, true),
    lessons,
    activeLessons
  };
}

function installAttendanceWorkbook() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error(
      'Откройте Google-таблицу и запустите установщик из связанного Apps Script.'
    );
  }

  const ui = SpreadsheetApp.getUi();
  const alreadyReady = isAttendanceInstallationComplete_(spreadsheet);

  const answer = ui.alert(
    alreadyReady ? 'Переустановить структуру журнала?' : 'Установить журнал посещаемости?',
    alreadyReady
      ? 'Структура уже отмечена как установленная. Переустановка очистит служебные листы.'
      : 'Установщик создаст или восстановит служебные листы. Частично созданную структуру можно безопасно запустить повторно.',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return;

  const disciplineResponse = ui.prompt(
    'Название дисциплины',
    'Введите название дисциплины.',
    ui.ButtonSet.OK_CANCEL
  );

  if (disciplineResponse.getSelectedButton() !== ui.Button.OK) return;

  const discipline = String(disciplineResponse.getResponseText() || '').trim()
    || 'Название дисциплины';

  try {
    const result = installAttendanceWorkbook_(spreadsheet, {
      discipline,
      locale: 'ru'
    });

    onOpen();

    ui.alert(
      'Установка завершена',
      'Создано новых листов: ' + result.createdSheets + '.\n' +
      'Версия: ' + result.appVersion + '.\n' +
      'Схема: ' + result.schemaVersion + '.\n\n' +
      'Теперь заполните «Список группы».',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      'Установка остановлена',
      String(e && e.message ? e.message : e) +
      '\n\nМожно исправить проблему и запустить установщик повторно: ' +
      'частичная структура будет пересоздана.',
      ui.ButtonSet.OK
    );
    throw e;
  }
}

function runInstallerStep_(name, fn) {
  try {
    return fn();
  } catch (e) {
    throw new Error(
      'Ошибка на этапе «' + name + '»: ' +
      String(e && e.message ? e.message : e)
    );
  }
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

  runInstallerStep_('Типы занятий', () => setupInstallerLessonTypes_(types));
  runInstallerStep_('Список группы', () => setupInstallerStudents_(students));
  runInstallerStep_('Журнал', () =>
    setupInstallerJournal_(journal, String(options.discipline || 'Название дисциплины'))
  );
  runInstallerStep_('Занятия: структура', () => setupInstallerLessonsBase_(lessons));
  runInstallerStep_('Настройки', () => setupInstallerSettings_(settings, spreadsheet, options));
  runInstallerStep_('Отметки', () => setupInstallerMarks_(marks));
  runInstallerStep_('Занятия: списки выбора', () =>
    applyInstallerLessonValidations_(lessons, types)
  );

  runInstallerStep_('Привязка экземпляра', () => {
    props_().setProperty(PROPS.BOUND_SPREADSHEET_ID, spreadsheet.getId());
    props_().deleteProperty(PROPS.ACTIVE_LESSON);
    props_().deleteProperty(PROPS.ACTIVE_CHECK);
    bindToContainer_();
  });

  runInstallerStep_('Перестройка журнала', () =>
    rebuildJournalFromRegistryInSpreadsheet_(spreadsheet)
  );

  runInstallerStep_('Очистка исходного пустого листа', () =>
    cleanupInstallerBlankSheets_(spreadsheet)
  );

  runInstallerStep_('Завершение установки', () => {
    setSettingValueInSpreadsheet_(
      spreadsheet,
      'install_status',
      'ready',
      'Состояние установки'
    );

    setSettingValueInSpreadsheet_(
      spreadsheet,
      'app_version',
      INSTALLER_APP_VERSION,
      'Версия приложения'
    );

    setSettingValueInSpreadsheet_(
      spreadsheet,
      'schema_version',
      INSTALLER_SCHEMA_VERSION,
      'Версия структуры данных'
    );

    syncVersionMetadata_(spreadsheet);
    migrateLegacyLessonStatuses_(spreadsheet);
    applyLessonValidationRules_(spreadsheet);
    applyLessonDateFormats_(spreadsheet);
  });

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

function ensureInstallerGrid_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < cols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  }

  const range = sheet.getRange(1, 1, rows, cols);
  if (range.isPartOfMerge()) range.breakApart();

  range
    .clearContent()
    .clearFormat()
    .clearDataValidations()
    .clearNote();
}

function setupInstallerStudents_(sheet) {
  ensureInstallerGrid_(sheet, 1000, 3);

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

function setupInstallerJournal_(sheet, discipline) {
  ensureInstallerGrid_(sheet, 300, 5);

  sheet.getRange('A1:E1')
    .breakApart()
    .clearContent()
    .clearFormat();

  sheet.getRange('A1')
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

function setupInstallerLessonsBase_(sheet) {
  ensureInstallerGrid_(sheet, 300, 14);

  sheet.getRange(1, 1, 1, 14)
    .setValues([[
      'ID занятия', 'Дата', 'Формат', 'Тип занятия', '№ по типу', 'Тема',
      'Статус', 'Начало', 'Завершение', 'Колонка Пос.', 'Колонка Оц.',
      'Проверок', 'Заметка', 'Создано'
    ]])
    .setFontWeight('bold')
    .setBackground('#EAEAEA')
    .setWrap(true);

  sheet.getRange(1, LESSON_COL.TYPE)
    .setNote('Можно менять после занятия; нумерация, подпись и цвет в Журнале обновляются.');
  sheet.getRange(1, LESSON_COL.TYPE_NUMBER)
    .setNote('Заполняется автоматически.');
  sheet.getRange(1, LESSON_COL.TOPIC)
    .setNote('Необязательное поле; можно редактировать после занятия.');
  sheet.getRange(1, LESSON_COL.NOTES)
    .setNote('Свободная заметка преподавателя.');

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

  sheet.getRange(2, LESSON_COL.DATE, sheet.getMaxRows() - 1, 1)
    .setNumberFormat('dd.MM.yyyy');

  [LESSON_COL.STARTED, LESSON_COL.ENDED, LESSON_COL.CREATED]
    .forEach(col => {
      sheet.getRange(2, col, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('dd.MM.yyyy HH:mm:ss');
    });
}

function applyInstallerLessonValidations_(sheet, typesSheet) {
  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Очно', 'Дистанционно'], true)
    .setAllowInvalid(false)
    .build();

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['active', 'closed'], true)
    .setAllowInvalid(false)
    .build();

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(typesSheet.getRange('A2:A100'), true)
    .setAllowInvalid(true)
    .build();

  const rows = Math.max(sheet.getMaxRows() - 1, 1);

  sheet.getRange(2, LESSON_COL.MODE, rows, 1).setDataValidation(modeRule);
  sheet.getRange(2, LESSON_COL.TYPE, rows, 1).setDataValidation(typeRule);
  sheet.getRange(2, LESSON_COL.STATUS, rows, 1).setDataValidation(statusRule);
}

function setupInstallerLessonTypes_(sheet) {
  ensureInstallerGrid_(sheet, 100, 4);

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

function setupInstallerSettings_(sheet, spreadsheet, options) {
  ensureInstallerGrid_(sheet, 100, 3);

  sheet.getRange('A1:C1')
    .setValues([['Ключ', 'Значение', 'Описание']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const rows = [
    ['discipline', String(options.discipline || 'Название дисциплины'), 'Название дисциплины'],
    ['locale', String(options.locale || 'ru'), 'Язык интерфейса: ru / en'],
    ['app_version', INSTALLER_APP_VERSION, 'Версия приложения'],
    ['code_version', INSTALLER_APP_VERSION, 'Версия установленного кода'],
    ['schema_version', INSTALLER_SCHEMA_VERSION, 'Версия структуры данных'],
    ['install_status', 'installing', 'Состояние установки'],
    ['regular_code_seconds', 60, 'Время действия обычного кода, секунд'],
    ['regular_code_cycles', 2, 'Пока не используется'],
    ['late_code_seconds', 60, 'Время действия кода опоздания, секунд'],
    ['late_code_cycles', 1, 'Пока не используется'],
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
    ['student_url', '', 'Прямая ссылка на страницу студента'],
    ['display_url', '', 'Прямая ссылка на экран кода'],
    ['sheet_url', spreadsheet.getUrl(), 'Прямая ссылка на этот журнал'],
    ['teacher_url', '', 'Секретная ссылка на мобильный пульт'],
    ['teacher_key', newTeacherKey_(), 'Секретный ключ пульта'],
    ['student_short_url', '', 'Необязательная короткая ссылка'],
    ['display_short_url', '', 'Необязательная короткая ссылка'],
    ['sheet_short_url', '', 'Необязательная короткая ссылка'],
    ['instance_spreadsheet_id', spreadsheet.getId(), 'ID экземпляра таблицы']
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 430);
  sheet.setColumnWidth(3, 560);
}

function setupInstallerMarks_(sheet) {
  ensureInstallerGrid_(sheet, 3000, 12);

  sheet.getRange(1, 1, 1, 12)
    .setValues([[
      'mark_id', 'lesson_id', 'student_no', 'student_name', 'check_type',
      'code_cycle', 'submitted_at', 'attendance_points', 'status',
      'client_token', 'source', 'notes'
    ]])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  sheet.setFrozenRows(1);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
}

function cleanupInstallerBlankSheets_(spreadsheet) {
  const expected = new Set([
    SHEETS.STUDENTS,
    SHEETS.JOURNAL,
    SHEETS.LESSONS,
    SHEETS.TYPES,
    SHEETS.SETTINGS,
    SHEETS.MARKS
  ]);

  spreadsheet.getSheets().slice().forEach(sh => {
    if (expected.has(sh.getName())) return;
    if (spreadsheet.getSheets().length <= 1) return;

    const isBlank =
      sh.getLastRow() <= 1 &&
      sh.getLastColumn() <= 1 &&
      sh.getRange('A1').getValue() === '';

    if (isBlank) spreadsheet.deleteSheet(sh);
  });
}

function diagnoseAttendanceInstallation() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!spreadsheet) {
    ui.alert('Нет активной Google-таблицы.');
    return;
  }

  if (spreadsheet.getSheetByName(SHEETS.SETTINGS)) {
    syncVersionMetadata_(spreadsheet);
    migrateLegacyLessonStatuses_(spreadsheet);
    applyLessonValidationRules_(spreadsheet);
    applyLessonDateFormats_(spreadsheet);
  }

  const health = getInstallationHealth_(spreadsheet);
  const lines = [];

  lines.push(health.missing.length ? 'Структура: ПРОБЛЕМА' : 'Структура: OK');

  if (health.missing.length) {
    lines.push('Не хватает листов: ' + health.missing.join(', '));
  }

  lines.push('Код: ' + INSTALLER_APP_VERSION);
  lines.push('Записанная версия: ' + (health.storedVersion || 'не задана'));
  lines.push(
    'Схема данных: ' +
    (health.storedSchema || 'не задана') +
    ' / ожидается ' +
    INSTALLER_SCHEMA_VERSION
  );
  lines.push('Состояние установки: ' + (health.installStatus || 'не задано'));
  lines.push('Web App URL: ' + (health.webAppConfigured ? 'настроен' : 'не настроен'));
  lines.push('');
  lines.push(
    'Студентов: ' + health.students +
    ' (активных: ' + health.activeStudents + ')'
  );
  lines.push(
    'Занятий: ' + health.lessons +
    ' (активных: ' + health.activeLessons + ')'
  );
  lines.push('');
  lines.push(
    health.missing.length ||
    health.installStatus !== 'ready' ||
    health.storedSchema !== String(INSTALLER_SCHEMA_VERSION)
      ? 'Итог: требуется проверка'
      : 'Итог: OK'
  );

  ui.alert(
    'Диагностика Attendance Sheet',
    lines.join('\n'),
    ui.ButtonSet.OK
  );
}
