/* -------------------- Настройки -------------------- */

function getSettings_() {
  const sh = ss_().getSheetByName(SHEETS.SETTINGS);
  if (!sh) throw new Error('Не найден лист "Настройки".');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const result = {};
  rows.forEach(([key, value]) => {
    if (key !== '' && key != null) result[String(key).trim()] = value;
  });
  return result;
}


function syncAppLinksInSpreadsheet_(spreadsheet) {
  /*
   * ВАЖНО:
   * ScriptApp.getService().getUrl() здесь намеренно НЕ используется.
   * В проекте может существовать несколько versioned deployments,
   * и getUrl() не даёт нам надёжного способа выбрать именно последнее
   * развертывание. Поэтому рабочий /exec хранится явно в web_app_url.
   */
  const configuredBase = String(
    getSettingValueInSpreadsheet_(
      spreadsheet,
      'web_app_url'
    ) || ''
  ).trim();

  const base =
    configuredBase
      ? normalizeWebAppUrl_(configuredBase)
      : '';

  let teacherKey = String(
    getSettingValueInSpreadsheet_(
      spreadsheet,
      'teacher_key'
    ) || ''
  ).trim();

  if (!teacherKey) {
    teacherKey =
      newTeacherKey_();

    setSettingValueInSpreadsheet_(
      spreadsheet,
      'teacher_key',
      teacherKey,
      'Секретный ключ пульта; создаётся автоматически для каждой копии'
    );
  }

  /*
   * Всегда сначала очищаем производные прямые ссылки.
   * Если deployment ID сменился, старые адреса не должны
   * оставаться в таблице ни при каких условиях.
   */
  ['student_url', 'display_url', 'teacher_url']
    .forEach(key => {
      setSettingValueInSpreadsheet_(
        spreadsheet,
        key,
        '',
        null
      );
    });

  const studentDirectUrl =
    base
      ? base + '?view=student'
      : '';

  const displayDirectUrl =
    base
      ? base + '?view=display'
      : '';

  const teacherDirectUrl =
    base
      ? (
          base +
          '?view=teacher&key=' +
          encodeURIComponent(teacherKey)
        )
      : '';

  const sheetDirectUrl =
    spreadsheet.getUrl();

  if (base) {
    setSettingValueInSpreadsheet_(
      spreadsheet,
      'web_app_url',
      base,
      'Базовый URL текущего рабочего развертывания /exec; задаётся после развертывания'
    );
  }

  setSettingValueInSpreadsheet_(
    spreadsheet,
    'student_url',
    studentDirectUrl,
    'Прямая ссылка на страницу студента; строится из web_app_url'
  );

  setSettingValueInSpreadsheet_(
    spreadsheet,
    'display_url',
    displayDirectUrl,
    'Прямая ссылка на экран кода; строится из web_app_url'
  );

  setSettingValueInSpreadsheet_(
    spreadsheet,
    'teacher_url',
    teacherDirectUrl,
    'Секретная ссылка на мобильный пульт; строится из web_app_url и teacher_key'
  );

  setSettingValueInSpreadsheet_(
    spreadsheet,
    'sheet_url',
    sheetDirectUrl,
    'Прямая ссылка на этот журнал; обновляется автоматически'
  );

  return {
    deployed: !!base,
    baseUrl: base,
    studentDirectUrl,
    displayDirectUrl,
    teacherDirectUrl,
    sheetDirectUrl
  };
}

function settingBool_(value, defaultValue) {
  if (value === '' || value == null) return !!defaultValue;
  if (value === true || value === false) return value;
  const s = String(value).trim().toLowerCase();
  if (['true','1','yes','да'].includes(s)) return true;
  if (['false','0','no','нет'].includes(s)) return false;
  return !!defaultValue;
}

function updateJournalTitleInSpreadsheet_(spreadsheet) {
  const settings = getSettingsFromSpreadsheet_(spreadsheet);
  const discipline = String(settings.discipline || 'Название дисциплины');
  const sh = spreadsheet.getSheetByName(SHEETS.JOURNAL);
  if (sh) sh.getRange(JOURNAL_LAYOUT.TITLE_ROW, 1).setValue('Журнал посещаемости — ' + discipline);
}

function getSettingsFromSpreadsheet_(spreadsheet) {
  const sh = spreadsheet.getSheetByName(SHEETS.SETTINGS);
  if (!sh) return {};
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};
  const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const result = {};
  rows.forEach(([key, value]) => {
    if (key !== '' && key != null) result[String(key).trim()] = value;
  });
  return result;
}

function getAppLinks_() {
  const spreadsheet = ss_();
  const settings = getSettings_();

  const configuredBase =
    String(
      settings.web_app_url || ''
    ).trim();

  const base =
    configuredBase
      ? normalizeWebAppUrl_(
          configuredBase
        )
      : '';

  const teacherKey =
    String(
      settings.teacher_key || ''
    ).trim();

  const studentDirect =
    base
      ? base + '?view=student'
      : String(
          settings.student_url || ''
        ).trim();

  const displayDirect =
    base
      ? base + '?view=display'
      : String(
          settings.display_url || ''
        ).trim();

  const teacherDirect =
    base && teacherKey
      ? (
          base +
          '?view=teacher&key=' +
          encodeURIComponent(
            teacherKey
          )
        )
      : String(
          settings.teacher_url || ''
        ).trim();

  const sheetDirect =
    spreadsheet.getUrl();

  const studentShort =
    String(
      settings.student_short_url || ''
    ).trim();

  const displayShort =
    String(
      settings.display_short_url || ''
    ).trim();

  const sheetShort =
    String(
      settings.sheet_short_url || ''
    ).trim();

  return {
    deployed: !!base,
    studentUrl:
      studentShort ||
      studentDirect,
    studentDirectUrl:
      studentDirect,
    displayUrl:
      displayShort ||
      displayDirect,
    displayDirectUrl:
      displayDirect,
    teacherUrl:
      teacherDirect,
    sheetUrl:
      sheetShort ||
      sheetDirect
  };
}


function getPublicConfig() {
  const settings = getSettings_();
  return {
    codeDigits: Number(settings.code_digits || 4),
    codeSeconds: Number(settings.regular_code_seconds || 60),
    discipline: String(settings.discipline || 'Дисциплина'),
    searchMinChars: Number(settings.student_search_min_chars || 3),
    rememberStudent: settingBool_(settings.remember_student_in_browser, true),
    instanceId: String(settings.instance_spreadsheet_id || '')
  };
}


