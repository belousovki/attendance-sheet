/* -------------------- v3.0 alpha.9: safe installer / recovery / reset -------------------- */

const INSTALLER_APP_VERSION = '3.0.0-alpha.9';
const INSTALLER_SCHEMA_VERSION = 3;

function isAttendanceInstallationComplete_(spreadsheet) {
  const settings = spreadsheet.getSheetByName(SHEETS.SETTINGS);
  if (!settings) return false;
  const schema = String(getSettingValueInSpreadsheet_(spreadsheet, 'schema_version') || '').trim();
  const status = String(getSettingValueInSpreadsheet_(spreadsheet, 'install_status') || '').trim();
  return schema === String(INSTALLER_SCHEMA_VERSION) && status === 'ready';
}

function getDefaultScoreScaleRows_() {
  return [
    [0, 0, '#D9EAF7', '0 — синий', true],
    [1, 2, '#D9EAD3', '1–2 — зелёный', true],
    [3, 4, '#E2F0D9', '3–4 — светло-зелёный', true],
    [5, 5, '#FFF2CC', '5 — жёлтый', true],
    [6, 7, '#FCE5CD', '6–7 — жёлто-оранжевый', true],
    [8, 9, '#F9CB9C', '8–9 — оранжевый', true],
    [10, '', '#F4CCCC', '10+ — красный', true]
  ];
}

function ensureScoreScaleSheet_(spreadsheet) {
  let sh = spreadsheet.getSheetByName(SHEETS.SCORE_SCALE);
  if (!sh) {
    sh = spreadsheet.insertSheet(SHEETS.SCORE_SCALE);
    setupScoreScaleSheet_(sh);
  } else if (sh.getLastRow() < 2) {
    setupScoreScaleSheet_(sh);
  }
  return sh;
}

function setupScoreScaleSheet_(sheet) {
  ensureInstallerGrid_(sheet, 100, 5);
  sheet.getRange('A1:E1')
    .setValues([['От', 'До', 'Цвет (HEX)', 'Описание', 'Активен']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const defaults = getDefaultScoreScaleRows_();
  sheet.getRange(2, 1, defaults.length, 5).setValues(defaults);

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setDataValidation(checkboxRule);

  defaults.forEach((row, idx) => {
    sheet.getRange(idx + 2, 3).setBackground(row[2]);
  });

  sheet.getRange('B1').setNote(
    'Пустое значение означает «и выше». Например, 10 + пустое «До» = 10+.'
  );
  sheet.getRange('C1').setNote(
    'Фиксированная шкала интенсивности балла: холодные цвета — низкие значения, тёплые — высокие.'
  );
  sheet.getRange('D1').setNote('Описание только для человека; на расчёты не влияет.');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 100);
}

function getScoreScale_(spreadsheet) {
  const sh = ensureScoreScaleSheet_(spreadsheet);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .map((r, idx) => {
      const minRaw = r[0];
      const maxRaw = r[1];
      const color = String(r[2] || '').trim();
      const active = settingBool_(r[4], true);
      if (
        minRaw === '' ||
        minRaw == null ||
        !Number.isFinite(Number(minRaw)) ||
        !/^#[0-9A-Fa-f]{6}$/.test(color) ||
        !active
      ) return null;
      const hasMax =
        maxRaw !== '' &&
        maxRaw != null &&
        Number.isFinite(Number(maxRaw));
      return {
        row: idx + 2,
        min: Number(minRaw),
        max: hasMax ? Number(maxRaw) : null,
        color,
        label: String(r[3] || '').trim()
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.min - b.min);
}

function getLessonScoreRanges_(spreadsheet) {
  const journal = spreadsheet.getSheetByName(SHEETS.JOURNAL);
  const lessons = spreadsheet.getSheetByName(SHEETS.LESSONS);
  if (!journal || !lessons || lessons.getLastRow() < 2) return [];

  const rows = lessons.getRange(
    2, 1, lessons.getLastRow() - 1, LESSON_COL.GRADE_COL
  ).getValues();

  const seen = new Set();
  const ranges = [];
  const height = Math.max(
    journal.getMaxRows() - JOURNAL_LAYOUT.DATA_START_ROW + 1,
    1
  );

  rows.forEach(r => {
    if (!String(r[LESSON_COL.ID - 1] || '').trim()) return;
    [r[LESSON_COL.ATTENDANCE_COL - 1], r[LESSON_COL.GRADE_COL - 1]].forEach(rawCol => {
      const col = Number(rawCol);
      if (
        !Number.isInteger(col) ||
        col < 3 ||
        col > journal.getMaxColumns() ||
        seen.has(col)
      ) return;
      seen.add(col);
      ranges.push(
        journal.getRange(JOURNAL_LAYOUT.DATA_START_ROW, col, height, 1)
      );
    });
  });
  return ranges;
}

function applyScoreScaleFormatting_(spreadsheet) {
  const journal = spreadsheet.getSheetByName(SHEETS.JOURNAL);
  if (!journal) return 0;
  const ranges = getLessonScoreRanges_(spreadsheet);
  const scale = getScoreScale_(spreadsheet);

  if (!ranges.length || !scale.length) {
    journal.setConditionalFormatRules([]);
    return 0;
  }

  const rules = scale.map(item => {
    let builder = SpreadsheetApp.newConditionalFormatRule();
    if (item.max == null) {
      builder = builder.whenNumberGreaterThanOrEqualTo(item.min);
    } else if (item.min === item.max) {
      builder = builder.whenNumberEqualTo(item.min);
    } else {
      builder = builder.whenNumberBetween(item.min, item.max);
    }
    return builder.setBackground(item.color).setRanges(ranges).build();
  });

  journal.setConditionalFormatRules(rules);
  return rules.length;
}

function syncVersionMetadata_(spreadsheet) {
  if (!spreadsheet.getSheetByName(SHEETS.SETTINGS)) return;
  setSettingValueInSpreadsheet_(
    spreadsheet, 'app_version', INSTALLER_APP_VERSION, 'Версия приложения'
  );
  setSettingValueInSpreadsheet_(
    spreadsheet, 'code_version', INSTALLER_APP_VERSION, 'Версия установленного кода'
  );
  if (
    String(getSettingValueInSpreadsheet_(spreadsheet, 'schema_version') || '').trim() === ''
  ) {
    setSettingValueInSpreadsheet_(
      spreadsheet, 'schema_version', INSTALLER_SCHEMA_VERSION, 'Версия структуры данных'
    );
  }
}

function migrateLegacyLessonStatuses_(spreadsheet) {
  const sh = spreadsheet.getSheetByName(SHEETS.LESSONS);
  if (!sh || sh.getLastRow() < 2) return 0;
  const range = sh.getRange(2, LESSON_COL.STATUS, sh.getLastRow() - 1, 1);
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
  sh.getRange(2, LESSON_COL.DATE, rows, 1).setNumberFormat('dd.MM.yyyy');
  [LESSON_COL.STARTED, LESSON_COL.ENDED, LESSON_COL.CREATED].forEach(col => {
    sh.getRange(2, col, rows, 1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
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
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().filter(r => {
    const hasStudent =
      r[0] !== '' &&
      r[0] != null &&
      String(r[1] || '').trim();
    if (!hasStudent) return false;
    return activeOnly ? settingBool_(r[2], true) : true;
  }).length;
}

function getInstallationHealth_(spreadsheet) {
  const required = [
    SHEETS.STUDENTS,
    SHEETS.JOURNAL,
    SHEETS.LESSONS,
    SHEETS.TYPES,
    SHEETS.SETTINGS,
    SHEETS.MARKS,
    SHEETS.SCORE_SCALE
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

  let lessons = 0;
  let activeLessons = 0;
  const lessonSheet = spreadsheet.getSheetByName(SHEETS.LESSONS);
  if (lessonSheet && lessonSheet.getLastRow() >= 2) {
    lessonSheet.getRange(
      2, 1, lessonSheet.getLastRow() - 1, LESSON_COL.STATUS
    ).getValues().forEach(r => {
      if (!String(r[LESSON_COL.ID - 1] || '').trim()) return;
      lessons++;
      if (String(r[LESSON_COL.STATUS - 1] || '').trim() === 'active') activeLessons++;
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
    throw new Error('Откройте Google-таблицу и запустите установщик из связанного Apps Script.');
  }

  if (isAttendanceInstallationComplete_(spreadsheet)) {
    return recoverAttendanceWorkbookFromMenu();
  }

  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Установить / восстановить журнал посещаемости?',
    'Приложение создаст недостающие листы и восстановит структуру.\n\n' +
    'Уже существующие студенты, занятия, оценки и история посещаемости не очищаются.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  let discipline = String(
    getSettingValueInSpreadsheet_(spreadsheet, 'discipline') || ''
  ).trim();

  if (!discipline) {
    const response = ui.prompt(
      'Название дисциплины',
      'Введите название дисциплины.',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    discipline = String(response.getResponseText() || '').trim() || 'Название дисциплины';
  }

  try {
    const result = installAttendanceWorkbook_(spreadsheet, {
      discipline,
      locale: 'ru'
    });
    onOpen();
    ui.alert(
      'Установка / восстановление завершены',
      'Создано новых листов: ' + result.createdSheets + '.\n' +
      'Версия: ' + result.appVersion + '.\n' +
      'Схема: ' + result.schemaVersion + '.\n\n' +
      'Существующие учебные данные не очищались.',
      ui.ButtonSet.OK
    );
    return result;
  } catch (e) {
    ui.alert(
      'Восстановление остановлено',
      String(e && e.message ? e.message : e) +
      '\n\nМожно устранить причину и запустить восстановление повторно. ' +
      'Операция не выполняет полный сброс данных.',
      ui.ButtonSet.OK
    );
    throw e;
  }
}

function installAttendanceWorkbook_(spreadsheet, options) {
  return repairAttendanceWorkbook_(spreadsheet, options || {});
}

function recoverAttendanceWorkbookFromMenu() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (!spreadsheet) {
    ui.alert('Нет активной Google-таблицы.');
    return;
  }

  const answer = ui.alert(
    'Восстановить структуру журнала?',
    'Операция создаст недостающие служебные элементы, восстановит проверки данных ' +
    'и перестроит производный лист «Журнал».\n\n' +
    'Список студентов, реестр занятий, оценки и история посещаемости не очищаются.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const discipline = String(
    getSettingValueInSpreadsheet_(spreadsheet, 'discipline') || 'Название дисциплины'
  ).trim() || 'Название дисциплины';

  try {
    const result = repairAttendanceWorkbook_(spreadsheet, {
      discipline,
      locale: String(getSettingValueInSpreadsheet_(spreadsheet, 'locale') || 'ru')
    });
    onOpen();
    ui.alert(
      'Восстановление завершено',
      'Создано недостающих листов: ' + result.createdSheets + '.\n' +
      'Учебные данные сохранены.\n' +
      'Версия: ' + result.appVersion + '.',
      ui.ButtonSet.OK
    );
    return result;
  } catch (e) {
    ui.alert(
      'Восстановление остановлено',
      String(e && e.message ? e.message : e) + '\n\nПолный сброс не выполнялся.',
      ui.ButtonSet.OK
    );
    throw e;
  }
}

function resetAttendanceWorkbookFromMenu() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (!spreadsheet) {
    ui.alert('Нет активной Google-таблицы.');
    return;
  }

  const first = ui.alert(
    'ПОЛНЫЙ СБРОС ЖУРНАЛА',
    'Будут удалены студенты, занятия, оценки, отметки посещаемости, ' +
    'настройки и текущая структура журнала.\n\n' +
    'Перед сбросом автоматически создаётся резервная копия текущего файла.\n\n' +
    'Для обычного исправления структуры используйте «Восстановить структуру журнала…».',
    ui.ButtonSet.YES_NO
  );
  if (first !== ui.Button.YES) return;

  const confirm = ui.prompt(
    'Подтверждение полного сброса',
    'Введите слово СБРОС заглавными буквами.',
    ui.ButtonSet.OK_CANCEL
  );
  if (
    confirm.getSelectedButton() !== ui.Button.OK ||
    String(confirm.getResponseText() || '').trim() !== 'СБРОС'
  ) {
    ui.alert('Полный сброс отменён.');
    return;
  }

  let backup;
  try {
    backup = createAttendanceBackup_(spreadsheet, 'перед полным сбросом');
  } catch (e) {
    ui.alert(
      'Полный сброс отменён',
      'Не удалось создать резервную копию:\n' +
      String(e && e.message ? e.message : e) +
      '\n\nИсходный журнал не изменён.',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.prompt(
    'Название дисциплины после сброса',
    'Можно оставить текущее название или указать новое.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Полный сброс отменён. Резервная копия уже создана:\n' + backup.url);
    return;
  }

  const currentDiscipline = String(
    getSettingValueInSpreadsheet_(spreadsheet, 'discipline') || 'Название дисциплины'
  ).trim();
  const discipline =
    String(response.getResponseText() || '').trim() ||
    currentDiscipline ||
    'Название дисциплины';

  const result = resetAttendanceWorkbookDestructive_(spreadsheet, {
    discipline,
    locale: 'ru'
  });
  onOpen();

  ui.alert(
    'Полный сброс завершён',
    'Журнал создан заново.\n\nРезервная копия исходного файла:\n' +
    backup.name + '\n' + backup.url,
    ui.ButtonSet.OK
  );
  return result;
}

function createAttendanceBackup_(spreadsheet, reason) {
  const file = DriveApp.getFileById(spreadsheet.getId());
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const tz =
    spreadsheet.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'GMT';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH-mm-ss');
  const suffix = reason
    ? ' — резервная копия (' + reason + ') '
    : ' — резервная копия ';
  const copy = file.makeCopy(spreadsheet.getName() + suffix + stamp, parent);
  return {
    id: copy.getId(),
    name: copy.getName(),
    url: copy.getUrl()
  };
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

function ensureInstallerSheet_(spreadsheet, name, created) {
  let sh = spreadsheet.getSheetByName(name);
  if (!sh) {
    sh = spreadsheet.insertSheet(name);
    if (created) created.push(name);
  }
  return sh;
}

function ensureInstallerGridSize_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < cols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  }
}

function ensureInstallerGrid_(sheet, rows, cols) {
  ensureInstallerGridSize_(sheet, rows, cols);
  const range = sheet.getRange(1, 1, rows, cols);
  if (range.isPartOfMerge()) range.breakApart();
  range.clearContent().clearFormat().clearDataValidations().clearNote();
}

function isRangeColumnEmpty_(sheet, column, startRow) {
  startRow = startRow || 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return true;
  return sheet.getRange(
    startRow, column, lastRow - startRow + 1, 1
  ).getDisplayValues().flat().every(v => !String(v || '').trim());
}

function setupInstallerStudents_(sheet) {
  ensureInstallerGrid_(sheet, 1000, 3);
  ensureRepairStudents_(sheet);
}

function ensureRepairStudents_(sheet) {
  ensureInstallerGridSize_(sheet, 1000, 3);
  sheet.getRange('A1:C1')
    .setValues([['№', 'ФИО', 'Активен']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1).setDataValidation(checkboxRule);

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 360);
  sheet.setColumnWidth(3, 100);
}

function setupInstallerJournal_(sheet, discipline) {
  ensureInstallerGrid_(sheet, 300, 5);
  sheet.getRange('A1:E1').breakApart().clearContent().clearFormat();
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
  ensureRepairLessons_(sheet);
}

function ensureRepairLessons_(sheet) {
  ensureInstallerGridSize_(sheet, 300, 14);
  sheet.getRange(1, 1, 1, 14)
    .setValues([[
      'ID занятия',
      'Дата',
      'Формат',
      'Тип занятия',
      '№ по типу',
      'Тема',
      'Статус',
      'Начало',
      'Завершение',
      'Колонка Пос.',
      'Колонка Оц.',
      'Проверок',
      'Заметка',
      'Создано'
    ]])
    .setFontWeight('bold')
    .setBackground('#EAEAEA')
    .setWrap(true);

  sheet.getRange(1, LESSON_COL.TYPE)
    .setNote('Можно менять после занятия; нумерация, подпись и цвет в Журнале обновляются.');
  sheet.getRange(1, LESSON_COL.TYPE_NUMBER).setNote('Заполняется автоматически.');
  sheet.getRange(1, LESSON_COL.TOPIC)
    .setNote('Необязательное поле; можно редактировать после занятия.');
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
  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 2).setDataValidation(checkboxRule);

  defaults.forEach((row, idx) => {
    sheet.getRange(idx + 2, 1, 1, 2).setBackground(row[1]);
  });

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidths(3, 2, 120);
}

function ensureRepairLessonTypes_(sheet) {
  ensureInstallerGridSize_(sheet, 100, 4);
  if (isRangeColumnEmpty_(sheet, 1, 2)) {
    setupInstallerLessonTypes_(sheet);
    return;
  }

  sheet.getRange('A1:D1')
    .setValues([['Тип занятия', 'Цвет (HEX)', 'Нумеровать', 'Активен']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 2).setDataValidation(checkboxRule);

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 4).getValues().forEach((r, idx) => {
      const typeName = String(r[0] || '').trim();
      const color = String(r[1] || '').trim();
      if (typeName && /^#[0-9A-Fa-f]{6}$/.test(color)) {
        sheet.getRange(idx + 2, 1, 1, 2).setBackground(color);
      }
    });
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidths(3, 2, 120);
}

function ensureRepairScoreScale_(sheet) {
  ensureInstallerGridSize_(sheet, 100, 5);
  if (isRangeColumnEmpty_(sheet, 1, 2)) {
    setupScoreScaleSheet_(sheet);
    return;
  }

  sheet.getRange('A1:E1')
    .setValues([['От', 'До', 'Цвет (HEX)', 'Описание', 'Активен']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setDataValidation(checkboxRule);

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 5).getValues().forEach((r, idx) => {
      const color = String(r[2] || '').trim();
      if (/^#[0-9A-Fa-f]{6}$/.test(color)) {
        sheet.getRange(idx + 2, 3).setBackground(color);
      }
    });
  }

  sheet.getRange('B1').setNote(
    'Пустое значение означает «и выше». Например, 10 + пустое «До» = 10+.'
  );
  sheet.getRange('C1').setNote(
    'Фиксированная шкала интенсивности балла: холодные цвета — низкие значения, тёплые — высокие.'
  );
  sheet.getRange('D1').setNote('Описание только для человека; на расчёты не влияет.');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 100);
}

function setupInstallerSettings_(sheet, spreadsheet, options) {
  ensureInstallerGrid_(sheet, 100, 3);
  sheet.getRange('A1:C1')
    .setValues([['Ключ', 'Значение', 'Описание']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  const rows = defaultInstallerSettingsRows_(spreadsheet, options, true);
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 430);
  sheet.setColumnWidth(3, 560);
}

function defaultInstallerSettingsRows_(spreadsheet, options, includeManaged) {
  const rows = [
    ['discipline', String(options.discipline || 'Название дисциплины'), 'Название дисциплины'],
    ['locale', String(options.locale || 'ru'), 'Язык интерфейса: ru / en'],
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
    ['teacher_url', '', 'Секретная ссылка на мобильный пульт'],
    ['student_short_url', '', 'Необязательная короткая ссылка'],
    ['display_short_url', '', 'Необязательная короткая ссылка'],
    ['sheet_short_url', '', 'Необязательная короткая ссылка']
  ];

  if (includeManaged) {
    rows.splice(2, 0,
      ['app_version', INSTALLER_APP_VERSION, 'Версия приложения'],
      ['code_version', INSTALLER_APP_VERSION, 'Версия установленного кода'],
      ['schema_version', INSTALLER_SCHEMA_VERSION, 'Версия структуры данных'],
      ['install_status', 'installing', 'Состояние установки']
    );
    rows.push(
      ['sheet_url', spreadsheet.getUrl(), 'Прямая ссылка на этот журнал'],
      ['teacher_key', newTeacherKey_(), 'Секретный ключ пульта'],
      ['instance_spreadsheet_id', spreadsheet.getId(), 'ID экземпляра таблицы']
    );
  }
  return rows;
}

function ensureSettingMissing_(spreadsheet, key, value, description) {
  if (getSettingValueInSpreadsheet_(spreadsheet, key) !== null) return false;
  setSettingValueInSpreadsheet_(spreadsheet, key, value, description);
  return true;
}

function ensureRepairSettings_(sheet, spreadsheet, options) {
  ensureInstallerGridSize_(sheet, 100, 3);
  sheet.getRange('A1:C1')
    .setValues([['Ключ', 'Значение', 'Описание']])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  defaultInstallerSettingsRows_(spreadsheet, options, false).forEach(row => {
    ensureSettingMissing_(spreadsheet, row[0], row[1], row[2]);
  });
  ensureSettingMissing_(
    spreadsheet, 'teacher_key', newTeacherKey_(), 'Секретный ключ пульта'
  );
  ensureSettingMissing_(
    spreadsheet, 'instance_spreadsheet_id', spreadsheet.getId(), 'ID экземпляра таблицы'
  );
  ensureSettingMissing_(
    spreadsheet, 'sheet_url', spreadsheet.getUrl(), 'Прямая ссылка на этот журнал'
  );

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 430);
  sheet.setColumnWidth(3, 560);
}

function setupInstallerMarks_(sheet) {
  ensureInstallerGrid_(sheet, 3000, 12);
  ensureRepairMarks_(sheet);
}

function ensureRepairMarks_(sheet) {
  ensureInstallerGridSize_(sheet, 3000, 12);
  sheet.getRange(1, 1, 1, 12)
    .setValues([[
      'mark_id',
      'lesson_id',
      'student_no',
      'student_name',
      'check_type',
      'code_cycle',
      'submitted_at',
      'attendance_points',
      'status',
      'client_token',
      'source',
      'notes'
    ]])
    .setFontWeight('bold')
    .setBackground('#EAEAEA');

  sheet.setFrozenRows(1);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
}

function repairAttendanceWorkbook_(spreadsheet, options) {
  options = options || {};
  const created = [];

  const students = ensureInstallerSheet_(spreadsheet, SHEETS.STUDENTS, created);
  const journal = ensureInstallerSheet_(spreadsheet, SHEETS.JOURNAL, created);
  const lessons = ensureInstallerSheet_(spreadsheet, SHEETS.LESSONS, created);
  const types = ensureInstallerSheet_(spreadsheet, SHEETS.TYPES, created);
  const settings = ensureInstallerSheet_(spreadsheet, SHEETS.SETTINGS, created);
  const marks = ensureInstallerSheet_(spreadsheet, SHEETS.MARKS, created);
  const scoreScale = ensureInstallerSheet_(spreadsheet, SHEETS.SCORE_SCALE, created);

  runInstallerStep_('Список группы', () => ensureRepairStudents_(students));
  runInstallerStep_('Типы занятий', () => ensureRepairLessonTypes_(types));
  runInstallerStep_('Шкала оценок', () => ensureRepairScoreScale_(scoreScale));
  runInstallerStep_('Занятия', () => ensureRepairLessons_(lessons));
  runInstallerStep_('Настройки', () => ensureRepairSettings_(settings, spreadsheet, options));
  runInstallerStep_('Отметки', () => ensureRepairMarks_(marks));
  runInstallerStep_('Журнал: размер', () => ensureInstallerGridSize_(journal, 300, 5));
  runInstallerStep_(
    'Занятия: списки выбора',
    () => applyInstallerLessonValidations_(lessons, types)
  );

  runInstallerStep_('Привязка экземпляра', () => {
    props_().setProperty(PROPS.BOUND_SPREADSHEET_ID, spreadsheet.getId());
    bindToContainer_();
  });

  runInstallerStep_(
    'Перестройка журнала',
    () => rebuildJournalFromRegistryInSpreadsheet_(spreadsheet)
  );
  runInstallerStep_('Очистка bootstrap-листа', () => cleanupInstallerBlankSheets_(spreadsheet));

  runInstallerStep_('Завершение восстановления', () => {
    setSettingValueInSpreadsheet_(
      spreadsheet, 'install_status', 'ready', 'Состояние установки'
    );
    setSettingValueInSpreadsheet_(
      spreadsheet, 'schema_version', INSTALLER_SCHEMA_VERSION, 'Версия структуры данных'
    );
    syncVersionMetadata_(spreadsheet);
    migrateLegacyLessonStatuses_(spreadsheet);
    applyLessonValidationRules_(spreadsheet);
    applyLessonDateFormats_(spreadsheet);
    ensureScoreScaleSheet_(spreadsheet);
    applyScoreScaleFormatting_(spreadsheet);
  });

  SpreadsheetApp.flush();
  return {
    createdSheets: created.length,
    sheetNames: created,
    appVersion: INSTALLER_APP_VERSION,
    schemaVersion: INSTALLER_SCHEMA_VERSION,
    mode: 'repair'
  };
}

function resetAttendanceWorkbookDestructive_(spreadsheet, options) {
  options = options || {};
  const created = [];

  const students = ensureInstallerSheet_(spreadsheet, SHEETS.STUDENTS, created);
  const journal = ensureInstallerSheet_(spreadsheet, SHEETS.JOURNAL, created);
  const lessons = ensureInstallerSheet_(spreadsheet, SHEETS.LESSONS, created);
  const types = ensureInstallerSheet_(spreadsheet, SHEETS.TYPES, created);
  const settings = ensureInstallerSheet_(spreadsheet, SHEETS.SETTINGS, created);
  const marks = ensureInstallerSheet_(spreadsheet, SHEETS.MARKS, created);
  const scoreScale = ensureInstallerSheet_(spreadsheet, SHEETS.SCORE_SCALE, created);

  runInstallerStep_('Типы занятий', () => setupInstallerLessonTypes_(types));
  runInstallerStep_('Шкала оценок', () => setupScoreScaleSheet_(scoreScale));
  runInstallerStep_('Список группы', () => setupInstallerStudents_(students));
  runInstallerStep_(
    'Журнал',
    () => setupInstallerJournal_(journal, String(options.discipline || 'Название дисциплины'))
  );
  runInstallerStep_('Занятия: структура', () => setupInstallerLessonsBase_(lessons));
  runInstallerStep_('Настройки', () => setupInstallerSettings_(settings, spreadsheet, options));
  runInstallerStep_('Отметки', () => setupInstallerMarks_(marks));
  runInstallerStep_(
    'Занятия: списки выбора',
    () => applyInstallerLessonValidations_(lessons, types)
  );

  runInstallerStep_('Привязка экземпляра', () => {
    props_().setProperty(PROPS.BOUND_SPREADSHEET_ID, spreadsheet.getId());
    props_().deleteProperty(PROPS.ACTIVE_LESSON);
    props_().deleteProperty(PROPS.ACTIVE_CHECK);
    bindToContainer_();
  });

  runInstallerStep_(
    'Перестройка журнала',
    () => rebuildJournalFromRegistryInSpreadsheet_(spreadsheet)
  );
  runInstallerStep_('Очистка bootstrap-листа', () => cleanupInstallerBlankSheets_(spreadsheet));

  runInstallerStep_('Завершение сброса', () => {
    setSettingValueInSpreadsheet_(
      spreadsheet, 'install_status', 'ready', 'Состояние установки'
    );
    setSettingValueInSpreadsheet_(
      spreadsheet, 'schema_version', INSTALLER_SCHEMA_VERSION, 'Версия структуры данных'
    );
    syncVersionMetadata_(spreadsheet);
    migrateLegacyLessonStatuses_(spreadsheet);
    applyLessonValidationRules_(spreadsheet);
    applyLessonDateFormats_(spreadsheet);
    ensureScoreScaleSheet_(spreadsheet);
    applyScoreScaleFormatting_(spreadsheet);
  });

  SpreadsheetApp.flush();
  return {
    createdSheets: created.length,
    sheetNames: created,
    appVersion: INSTALLER_APP_VERSION,
    schemaVersion: INSTALLER_SCHEMA_VERSION,
    mode: 'reset'
  };
}

function cleanupInstallerBlankSheets_(spreadsheet) {
  const sh = spreadsheet.getSheetByName('Установщик');
  if (!sh) return false;
  if (spreadsheet.getSheets().length <= 1) return false;

  const expectedShape =
    sh.getMaxRows() === 100 &&
    sh.getMaxColumns() === 10;
  const isBlank = sh.getDataRange()
    .getDisplayValues()
    .flat()
    .every(v => !String(v || '').trim());

  if (expectedShape && isBlank) {
    spreadsheet.deleteSheet(sh);
    return true;
  }
  return false;
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
    'Студентов: ' + health.students + ' (активных: ' + health.activeStudents + ')'
  );
  lines.push(
    'Занятий: ' + health.lessons + ' (активных: ' + health.activeLessons + ')'
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
