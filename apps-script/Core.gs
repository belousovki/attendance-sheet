/**
 * Универсальный журнал посещаемости v3.0 alpha.9 (ядро v2.7 + безопасный установщик).
 * Одна дисциплина / одна группа / один файл.
 * Настраиваемые типы занятий, темы, заметки и цвета.
 */


const SHEETS = {
  STUDENTS: 'Список группы',
  JOURNAL: 'Журнал',
  LESSONS: 'Занятия',
  TYPES: 'Типы занятий',
  MARKS: 'Отметки',
  SETTINGS: 'Настройки',
  SCORE_SCALE: 'Шкала оценок'
};

const JOURNAL_LAYOUT = {
  TITLE_ROW: 1,
  META_ROW: 2,
  SUBHEADER_ROW: 3,
  DATA_START_ROW: 4
};

const LESSON_COL = {
  ID: 1,
  DATE: 2,
  MODE: 3,
  TYPE: 4,
  TYPE_NUMBER: 5,
  TOPIC: 6,
  STATUS: 7,
  STARTED: 8,
  ENDED: 9,
  ATTENDANCE_COL: 10,
  GRADE_COL: 11,
  CHECK_COUNT: 12,
  NOTES: 13,
  CREATED: 14
};

const PROPS = {
  ACTIVE_LESSON: 'ATT_ACTIVE_LESSON',
  ACTIVE_CHECK: 'ATT_ACTIVE_CHECK',
  BOUND_SPREADSHEET_ID: 'ATT_BOUND_SPREADSHEET_ID'
};

function ss_() {
  const id =
    props_().getProperty(
      PROPS.BOUND_SPREADSHEET_ID
    );

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  /*
   * Fallback only for calls made from the spreadsheet UI.
   * In a web-app execution there may be no "active spreadsheet",
   * so the ID must normally have been saved by onOpen().
   */
  const active =
    SpreadsheetApp.getActiveSpreadsheet();

  if (active) {
    props_().setProperty(
      PROPS.BOUND_SPREADSHEET_ID,
      active.getId()
    );

    return active;
  }

  throw new Error(
    'Экземпляр не инициализирован. ' +
    'Откройте связанную Google-таблицу один раз, ' +
    'затем обновите развертывание.'
  );
}

function props_() {
  return PropertiesService.getScriptProperties();
}


function getSettingValueInSpreadsheet_(spreadsheet, key) {
  const sh =
    spreadsheet.getSheetByName(
      SHEETS.SETTINGS
    );

  if (!sh) {
    return null;
  }

  const lastRow =
    sh.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const rows =
    sh
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();

  for (const row of rows) {
    if (
      String(row[0] || '').trim() === key
    ) {
      return row[1];
    }
  }

  return null;
}

function setSettingValueInSpreadsheet_(
  spreadsheet,
  key,
  value,
  description
) {
  const sh =
    spreadsheet.getSheetByName(
      SHEETS.SETTINGS
    );

  if (!sh) {
    throw new Error(
      'Не найден лист "Настройки".'
    );
  }

  const lastRow =
    Math.max(sh.getLastRow(), 1);

  if (lastRow >= 2) {
    const keys =
      sh
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getDisplayValues()
        .flat();

    const idx =
      keys.findIndex(
        k =>
          String(k).trim() === key
      );

    if (idx >= 0) {
      sh
        .getRange(
          idx + 2,
          2
        )
        .setValue(value);

      if (description) {
        sh
          .getRange(
            idx + 2,
            3
          )
          .setValue(description);
      }

      return;
    }
  }

  sh.appendRow([
    key,
    value,
    description || ''
  ]);
}


function dedupeSettingKeyInSpreadsheet_(spreadsheet, key) {
  const sh = spreadsheet.getSheetByName(SHEETS.SETTINGS);
  if (!sh) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const keys = sh.getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues()
    .flat()
    .map(v => String(v || '').trim());

  const matches = [];
  keys.forEach((k, idx) => {
    if (k === key) matches.push(idx + 2);
  });

  if (matches.length <= 1) return 0;

  // Оставляем первую строку, лишние удаляем снизу вверх.
  matches.slice(1)
    .sort((a, b) => b - a)
    .forEach(row => sh.deleteRow(row));

  return matches.length - 1;
}

function dedupeManagedUrlSettings_(spreadsheet) {
  [
    'web_app_url',
    'student_url',
    'display_url',
    'sheet_url',
    'teacher_url',
    'teacher_key',
    'student_short_url',
    'display_short_url',
    'sheet_short_url'
  ].forEach(key => dedupeSettingKeyInSpreadsheet_(spreadsheet, key));
}

function newTeacherKey_() {
  return (
    Utilities.getUuid()
      .replace(/-/g, '') +
    Utilities.getUuid()
      .replace(/-/g, '')
  ).slice(0, 32);
}

function bindToContainer_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) return null;

  const currentId = active.getId();
  props_().setProperty(PROPS.BOUND_SPREADSHEET_ID, currentId);

  const storedInstanceId = String(
    getSettingValueInSpreadsheet_(active, 'instance_spreadsheet_id') || ''
  ).trim();

  const isCopy = !!storedInstanceId && storedInstanceId !== currentId;

  if (!storedInstanceId || isCopy) {
    setSettingValueInSpreadsheet_(
      active,
      'instance_spreadsheet_id',
      currentId,
      'ID экземпляра; меняется автоматически при копировании файла'
    );
  }

  if (isCopy) {
    // Новая копия получает собственные секреты и собственные URL.
    setSettingValueInSpreadsheet_(active, 'teacher_key', newTeacherKey_(),
      'Секретный ключ пульта; создаётся автоматически для каждой копии');

    ['web_app_url','student_url','display_url','teacher_url','student_short_url','display_short_url','sheet_short_url']
      .forEach(key => setSettingValueInSpreadsheet_(active, key, '', null));

    props_().deleteProperty(PROPS.ACTIVE_LESSON);
    props_().deleteProperty(PROPS.ACTIVE_CHECK);
  }

  const teacherKey = String(
    getSettingValueInSpreadsheet_(active, 'teacher_key') || ''
  ).trim();

  if (!teacherKey) {
    setSettingValueInSpreadsheet_(active, 'teacher_key', newTeacherKey_(),
      'Секретный ключ пульта; создаётся автоматически для каждой копии');
  }

  setSettingValueInSpreadsheet_(active, 'sheet_url', active.getUrl(),
    'Прямая ссылка на этот журнал; обновляется автоматически');

  dedupeManagedUrlSettings_(active);
  syncAppLinksInSpreadsheet_(active);
  updateJournalTitleInSpreadsheet_(active);
  return active;
}

function onOpen() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) return;

  if (!isAttendanceInstallationComplete_(active)) {
    SpreadsheetApp.getUi()
      .createMenu('Посещаемость')
      .addItem('Установить / восстановить журнал…', 'installAttendanceWorkbook')
      .addItem('Диагностика установки', 'diagnoseAttendanceInstallation')
      .addToUi();
    return;
  }

  bindToContainer_();
  syncVersionMetadata_(active);
  migrateLegacyLessonStatuses_(active);
  applyLessonDateFormats_(active);
  ensureScoreScaleSheet_(active);
  applyScoreScaleFormatting_(active);

  SpreadsheetApp.getUi()
    .createMenu('Посещаемость')
    .addItem('Начать занятие…', 'startLessonFromMenu_')
    .addItem('Завершить активное занятие', 'finishLessonFromMenu')
    .addSeparator()
    .addItem('Открыть панель', 'showAttendanceSidebar')
    .addItem('Синхронизировать состав группы', 'syncStudentsFromMenu')
    .addItem('Перестроить журнал из исходных данных', 'rebuildJournalFromMenu')
    .addItem('Восстановить структуру журнала…', 'recoverAttendanceWorkbookFromMenu')
    .addSeparator()
    .addItem('Удалить занятие по выбранному столбцу…', 'deleteLessonFromJournalSelection_')
    .addItem('Удалить занятие по выбранной строке…', 'deleteSelectedLessonFromMenu')
    .addSeparator()
    .addItem('Настроить URL веб-приложения', 'setWebAppUrlFromMenu')
    .addItem('Обновить ссылки приложения', 'refreshAppLinksFromMenu')
    .addSeparator()
    .addItem('Подготовить как пустой шаблон…', 'prepareBlankTemplateFromMenu')
    .addItem('Полный сброс журнала…', 'resetAttendanceWorkbookFromMenu')
    .addItem('Диагностика установки', 'diagnoseAttendanceInstallation')
    .addToUi();
}

function showAttendanceSidebar() {
  bindToContainer_();

  syncStudents_();
  recomputeAllTotals_();

  const tpl =
    HtmlService.createTemplateFromFile(
      'Sidebar'
    );

  tpl.teacherKey =
    String(
      getSettings_().teacher_key || ''
    );

  const html =
    tpl.evaluate()
      .setTitle(
        'Посещаемость'
      );

  SpreadsheetApp
    .getUi()
    .showSidebar(html);
}



function normalizeWebAppUrl_(raw) {
  let value = String(raw || '').trim();

  if (!value) {
    return '';
  }

  /*
   * Пользователь может случайно вставить student/display/teacher URL.
   * Оставляем только базовый адрес развертывания.
   */
  value = value.split('#')[0].split('?')[0].replace(/\/+$/, '');

  const match = value.match(
    /^https:\/\/script\.google\.com\/macros\/s\/[^\/]+\/(exec|dev)$/
  );

  if (!match) {
    throw new Error(
      'Нужен URL веб-приложения вида ' +
      'https://script.google.com/macros/s/.../exec'
    );
  }

  if (match[1] !== 'exec') {
    throw new Error(
      'Используйте URL рабочего развертывания /exec, а не тестовый /dev.'
    );
  }

  return value;
}

function setWebAppUrlFromMenu() {
  assertBoundSpreadsheetUi_();

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  dedupeManagedUrlSettings_(
    spreadsheet
  );

  const currentRaw = String(
    getSettingValueInSpreadsheet_(
      spreadsheet,
      'web_app_url'
    ) || ''
  ).trim();

  let current = '';

  if (currentRaw) {
    try {
      current =
        normalizeWebAppUrl_(
          currentRaw
        );
    } catch (e) {
      current = currentRaw;
    }
  }

  const ui =
    SpreadsheetApp.getUi();

  const response =
    ui.prompt(
      'URL веб-приложения',
      'После Развернуть → Управление развертываниями ' +
      'скопируйте URL, заканчивающийся на /exec, и вставьте его сюда.' +
      (current ? '\n\nТекущий URL:\n' + current : ''),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  let base;

  try {
    base = normalizeWebAppUrl_(
      response.getResponseText()
    );
  } catch (e) {
    ui.alert(
      String(e.message || e)
    );
    return;
  }

  const deploymentChanged =
    !!current &&
    current !== base;

  /*
   * Прямые ссылки всегда пересоздаём.
   */
  [
    'student_url',
    'display_url',
    'teacher_url'
  ].forEach(key => {
    setSettingValueInSpreadsheet_(
      spreadsheet,
      key,
      '',
      null
    );
  });

  /*
   * Если deployment ID изменился, короткие ссылки
   * на student/display почти наверняка ведут на старое развертывание.
   * Поэтому их тоже очищаем.
   *
   * sheet_short_url сохраняем: Google-таблица та же.
   */
  if (deploymentChanged) {
    [
      'student_short_url',
      'display_short_url'
    ].forEach(key => {
      setSettingValueInSpreadsheet_(
        spreadsheet,
        key,
        '',
        null
      );
    });
  }

  setSettingValueInSpreadsheet_(
    spreadsheet,
    'web_app_url',
    base,
    'Базовый URL текущего рабочего развертывания /exec; задаётся после развертывания'
  );

  dedupeManagedUrlSettings_(
    spreadsheet
  );

  const links =
    syncAppLinksInSpreadsheet_(
      spreadsheet
    );

  let message =
    'URL сохранён. Прямые ссылки студента, дисплея и пульта пересозданы.';

  if (deploymentChanged) {
    message +=
      '\n\nСтарые короткие ссылки student/display очищены, ' +
      'потому что они могли вести на прежнее развертывание.';
  }

  ui.alert(message);

  return links;
}

function refreshAppLinksFromMenu() {
  assertBoundSpreadsheetUi_();

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const base = String(
    getSettingValueInSpreadsheet_(
      spreadsheet,
      'web_app_url'
    ) || ''
  ).trim();

  if (!base) {
    SpreadsheetApp.getUi().alert(
      'Сначала настройте URL веб-приложения.'
    );
    return;
  }

  try {
    normalizeWebAppUrl_(base);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Сохранённый URL веб-приложения некорректен. ' +
      'Запустите «Настроить URL веб-приложения» и укажите рабочий /exec.'
    );
    return;
  }

  const links =
    syncAppLinksInSpreadsheet_(
      spreadsheet
    );

  SpreadsheetApp.getUi().alert(
    'Ссылки приложения обновлены.'
  );

  return links;
}

function startLessonFromMenu_() {
  assertBoundSpreadsheetUi_();

  const ui =
    SpreadsheetApp.getUi();

  if (getActiveLesson_()) {
    ui.alert(
      'Уже есть активное занятие. Сначала завершите его.'
    );
    return;
  }

  const types =
    getLessonTypes_(true);

  if (!types.length) {
    ui.alert(
      'На листе «Типы занятий» нет активных типов занятий.'
    );
    return;
  }

  const typeNames =
    types.map(t => t.name);

  const typeResponse =
    ui.prompt(
      'Начать занятие',
      'Введите тип занятия:\n' +
      typeNames.join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    typeResponse.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const lessonType =
    String(
      typeResponse.getResponseText() || ''
    ).trim();

  const typeConfig =
    types.find(
      t => t.name === lessonType
    );

  if (!typeConfig) {
    ui.alert(
      'Тип занятия не найден среди активных типов.'
    );
    return;
  }

  const modeResponse =
    ui.prompt(
      'Формат занятия',
      'Введите «Очно» или «Дистанционно».',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    modeResponse.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const rawMode =
    String(
      modeResponse.getResponseText() || ''
    ).trim();

  const normalizedMode =
    rawMode.toLowerCase();

  let mode;

  if (
    normalizedMode === 'очно'
  ) {
    mode = 'Очно';
  } else if (
    normalizedMode === 'дистанционно'
  ) {
    mode = 'Дистанционно';
  } else {
    ui.alert(
      'Формат должен быть «Очно» или «Дистанционно».'
    );
    return;
  }

  const topicResponse =
    ui.prompt(
      'Тема занятия',
      'Введите тему занятия (можно оставить пустой).',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    topicResponse.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const topic =
    String(
      topicResponse.getResponseText() || ''
    ).trim();

  try {
    const result =
      startLesson_({
        lessonType,
        mode,
        topic,
        notes: ''
      });

    ui.alert(
      'Занятие начато',
      (result.typeNumber
        ? lessonType + ' №' + result.typeNumber
        : lessonType) +
      (topic ? '\nТема: ' + topic : '') +
      '\n\nВ «Журнале» создана пара столбцов Пос./Оц.',
      ui.ButtonSet.OK
    );

    return result;
  } catch (e) {
    ui.alert(
      'Не удалось начать занятие',
      String(
        e && e.message
          ? e.message
          : e
      ),
      ui.ButtonSet.OK
    );
    throw e;
  }
}

function assertBoundSpreadsheetUi_() {
  const active =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!active) {
    throw new Error(
      'Нет активной Google-таблицы.'
    );
  }

  props_().setProperty(
    PROPS.BOUND_SPREADSHEET_ID,
    active.getId()
  );

  return active;
}


function showInstallStatus_() {
  const ss = ss_();
  const ui = SpreadsheetApp.getUi();

  const health =
    getInstallationHealth_(ss);

  const lines = [];

  lines.push(
    health.missing.length
      ? 'Не хватает листов: ' +
        health.missing.join(', ')
      : 'Структура: OK'
  );

  lines.push(
    'Приложение: ' +
    (health.storedVersion || 'не задано')
  );

  lines.push(
    'Схема: ' +
    (health.storedSchema || 'не задана')
  );

  lines.push(
    'Статус: ' +
    (health.installStatus || 'не задан')
  );

  ui.alert(
    'Состояние установки',
    lines.join('\n'),
    ui.ButtonSet.OK
  );
}
