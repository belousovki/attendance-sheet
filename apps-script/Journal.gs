function captureJournalLessonValues_(spreadsheet) {
  const journal = spreadsheet.getSheetByName(SHEETS.JOURNAL);
  const lastCol = journal.getLastColumn();
  const lastRow = journal.getLastRow();
  const result = new Map();

  if (lastCol < 3 || lastRow < JOURNAL_LAYOUT.DATA_START_ROW) return result;

  const subHeaders = journal
    .getRange(JOURNAL_LAYOUT.SUBHEADER_ROW, 1, 1, lastCol)
    .getDisplayValues()[0];
  const notes = journal
    .getRange(JOURNAL_LAYOUT.SUBHEADER_ROW, 1, 1, lastCol)
    .getNotes()[0];

  const lessonCols = {};
  for (let col = 3; col <= lastCol; col++) {
    const note = String(notes[col - 1] || '');
    const m = note.match(/^lesson_id=(.+)$/);
    if (!m) continue;
    const lessonId = m[1];
    lessonCols[lessonId] = lessonCols[lessonId] || {};
    if (String(subHeaders[col - 1]) === 'Пос.') lessonCols[lessonId].att = col;
    if (String(subHeaders[col - 1]) === 'Оц.') lessonCols[lessonId].grade = col;
  }

  const rowCount = lastRow - JOURNAL_LAYOUT.DATA_START_ROW + 1;
  const values = journal
    .getRange(JOURNAL_LAYOUT.DATA_START_ROW, 1, rowCount, lastCol)
    .getValues();

  values.forEach(row => {
    if (row[0] === '' || row[0] == null) return;
    const studentNo = String(row[0]);

    Object.keys(lessonCols).forEach(lessonId => {
      const cols = lessonCols[lessonId];
      if (!cols.att || !cols.grade) return;
      result.set(studentNo + '\u0001' + lessonId, {
        attendance: row[cols.att - 1],
        grade: row[cols.grade - 1]
      });
    });
  });

  return result;
}

function readLessonRegistryInSpreadsheet_(spreadsheet) {
  const sh = spreadsheet.getSheetByName(SHEETS.LESSONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const rows = sh.getRange(2, 1, lastRow - 1, LESSON_COL.CREATED).getValues();
  return rows
    .map((r, idx) => ({ row: idx + 2, values: r }))
    .filter(x => String(x.values[LESSON_COL.ID - 1] || '').trim());
}

function rebuildJournalFromRegistryInSpreadsheet_(spreadsheet) {
  const journal = spreadsheet.getSheetByName(SHEETS.JOURNAL);
  const lessonsSheet = spreadsheet.getSheetByName(SHEETS.LESSONS);
  const saved = captureJournalLessonValues_(spreadsheet);

  // Нумерация по типам пересчитывается перед построением шапки.
  recalculateLessonTypeNumbers_();
  const lessons = readLessonRegistryInSpreadsheet_(spreadsheet);

  /*
   * v2.6:
   * Нельзя удалять ВСЕ незакреплённые столбцы Google Sheets.
   * Раньше использовалось:
   *   deleteColumns(3, maxCols - 2)
   * при A:B frozen, и Google выдавал:
   * «Невозможно удалить все незакрепленные столбцы».
   *
   * Теперь сначала рассчитываем конечную ширину Journal,
   * оставляем минимум C:E и только подгоняем число столбцов.
   */
  const desiredCols = 2 + lessons.length * 2 + 3; // A:B + пары занятий + 3 итога
  let maxCols = journal.getMaxColumns();

  // Снимаем старые объединения заголовков занятий перед очисткой.
  if (maxCols > 2) {
    journal
      .getRange(
        JOURNAL_LAYOUT.META_ROW,
        3,
        1,
        maxCols - 2
      )
      .breakApart();

    // Очищаем старую производную часть Journal.
    journal
      .getRange(
        JOURNAL_LAYOUT.META_ROW,
        3,
        Math.max(
          journal.getMaxRows() - JOURNAL_LAYOUT.META_ROW + 1,
          1
        ),
        maxCols - 2
      )
      .clearContent()
      .clearNote();
  }

  // Сначала при необходимости добавляем столбцы.
  if (maxCols < desiredCols) {
    journal.insertColumnsAfter(
      maxCols,
      desiredCols - maxCols
    );
    maxCols = journal.getMaxColumns();
  }

  // Потом удаляем только ЛИШНИЕ, но никогда не все незакреплённые.
  if (maxCols > desiredCols) {
    journal.deleteColumns(
      desiredCols + 1,
      maxCols - desiredCols
    );
  }

  // Базовые заголовки.
  journal
    .getRange(
      JOURNAL_LAYOUT.META_ROW,
      1,
      2,
      2
    )
    .setValues([
      ['№', 'ФИО'],
      ['', '']
    ])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#EAEAEA');

  let col = 3;
  const tz = spreadsheet.getSpreadsheetTimeZone();
  const lessonIdSet = new Set();

  lessons.forEach(item => {
    const r = item.values;
    const lessonId = String(r[LESSON_COL.ID - 1] || '').trim();
    lessonIdSet.add(lessonId);

    const date = r[LESSON_COL.DATE - 1];
    const typeName = String(r[LESSON_COL.TYPE - 1] || '').trim();
    const typeNumber = r[LESSON_COL.TYPE_NUMBER - 1];
    const topic = String(r[LESSON_COL.TOPIC - 1] || '').trim();

    const dateDisplay =
      date instanceof Date
        ? Utilities.formatDate(date, tz, 'dd.MM')
        : String(date || '');

    const config = getLessonType_(typeName);

    const bg =
      /^#[0-9A-Fa-f]{6}$/.test(
        String(config.color || '')
      )
        ? config.color
        : '#EDEDED';

    const metaRange =
      journal.getRange(
        JOURNAL_LAYOUT.META_ROW,
        col,
        1,
        2
      );

    metaRange.merge();

    metaRange
      .setValue(
        formatLessonMetaLabel_(
          dateDisplay,
          typeName,
          typeNumber,
          topic
        )
      )
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBackground(bg)
      .setNote('lesson_id=' + lessonId);

    journal
      .getRange(
        JOURNAL_LAYOUT.SUBHEADER_ROW,
        col,
        1,
        2
      )
      .setValues([['Пос.', 'Оц.']])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground(bg)
      .setNotes([[
        'lesson_id=' + lessonId,
        'lesson_id=' + lessonId
      ]]);

    journal.setColumnWidth(col, 78);
    journal.setColumnWidth(col + 1, 78);

    lessonsSheet
      .getRange(
        item.row,
        LESSON_COL.ATTENDANCE_COL
      )
      .setValue(col);

    lessonsSheet
      .getRange(
        item.row,
        LESSON_COL.GRADE_COL
      )
      .setValue(col + 1);

    col += 2;
  });

  const summaryCol = col;

  journal
    .getRange(
      JOURNAL_LAYOUT.META_ROW,
      summaryCol,
      1,
      3
    )
    .setValues([[
      'Баллы за посещение',
      'Баллы за работу',
      'Итого'
    ]])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground('#EAEAEA');

  journal
    .getRange(
      JOURNAL_LAYOUT.SUBHEADER_ROW,
      summaryCol,
      1,
      3
    )
    .clearContent()
    .setBackground('#EAEAEA');

  journal.setFrozenRows(
    JOURNAL_LAYOUT.SUBHEADER_ROW
  );
  journal.setFrozenColumns(2);

  journal.setColumnWidth(1, 64);
  journal.setColumnWidth(2, 360);
  journal.setColumnWidths(
    summaryCol,
    3,
    145
  );

  // Возвращаем компактный список студентов.
  syncStudentsInSpreadsheet_(spreadsheet);

  // Посещаемость восстанавливается из Marks; оценки — из прежнего Journal.
  const latestMarks =
    getLatestAttendanceMarksMap_();

  const students =
    getStudents_(false);

  students.forEach((student, idx) => {
    const row =
      JOURNAL_LAYOUT.DATA_START_ROW +
      idx;

    lessons.forEach(item => {
      const r = item.values;

      const lessonId =
        String(
          r[LESSON_COL.ID - 1] || ''
        ).trim();

      const key =
        student.no +
        '\u0001' +
        lessonId;

      const savedValue =
        saved.get(key);

      const latestMark =
        latestMarks.get(key);

      const attCol =
        Number(
          lessonsSheet
            .getRange(
              item.row,
              LESSON_COL.ATTENDANCE_COL
            )
            .getValue()
        );

      const gradeCol =
        Number(
          lessonsSheet
            .getRange(
              item.row,
              LESSON_COL.GRADE_COL
            )
            .getValue()
        );

      if (
        latestMark &&
        latestMark.points !== '' &&
        latestMark.points != null &&
        Number.isFinite(
          Number(latestMark.points)
        )
      ) {
        journal
          .getRange(row, attCol)
          .setValue(
            Number(latestMark.points)
          );
      } else if (
        savedValue &&
        savedValue.attendance !== '' &&
        savedValue.attendance != null
      ) {
        journal
          .getRange(row, attCol)
          .setValue(
            savedValue.attendance
          );
      }

      if (
        savedValue &&
        savedValue.grade !== '' &&
        savedValue.grade != null
      ) {
        journal
          .getRange(row, gradeCol)
          .setValue(
            savedValue.grade
          );
      }
    });
  });

  recomputeAllTotalsInSpreadsheet_(spreadsheet);
  clearPhantomJournalTotals_(spreadsheet);
  updateJournalTitleInSpreadsheet_(spreadsheet);

  const active =
    getActiveLesson_();

  if (active) {
    if (
      !lessonIdSet.has(
        String(active.lessonId)
      )
    ) {
      props_().deleteProperty(
        PROPS.ACTIVE_LESSON
      );
      props_().deleteProperty(
        PROPS.ACTIVE_CHECK
      );
    } else {
      const row =
        findLessonRow_(
          active.lessonId
        );

      if (row) {
        active.attendanceCol =
          Number(
            lessonsSheet
              .getRange(
                row,
                LESSON_COL.ATTENDANCE_COL
              )
              .getValue()
          );

        active.gradeCol =
          Number(
            lessonsSheet
              .getRange(
                row,
                LESSON_COL.GRADE_COL
              )
              .getValue()
          );

        saveActiveLesson_(active);
      }
    }
  }

  SpreadsheetApp.flush();

  return {
    students: students.length,
    lessons: lessons.length
  };
}



function deleteLessonByIdFromUi_(spreadsheet, lessonId) {
  const ui = SpreadsheetApp.getUi();

  const lessons =
    spreadsheet.getSheetByName(
      SHEETS.LESSONS
    );

  const row =
    findLessonRow_(
      lessonId
    );

  if (!row) {
    ui.alert(
      'Занятие не найдено на листе «Занятия». ' +
      'Возможно, Журнал устарел. Выполните «Перестроить журнал из исходных данных».'
    );
    return false;
  }

  const date =
    lessons
      .getRange(
        row,
        LESSON_COL.DATE
      )
      .getDisplayValue();

  const type =
    lessons
      .getRange(
        row,
        LESSON_COL.TYPE
      )
      .getDisplayValue();

  const topic =
    lessons
      .getRange(
        row,
        LESSON_COL.TOPIC
      )
      .getDisplayValue();

  const label =
    [date, type, topic]
      .filter(Boolean)
      .join(' · ');

  const answer =
    ui.alert(
      'Удалить занятие?',
      (label ? label + '\n\n' : '') +
      'Будут удалены запись занятия, его технические отметки посещаемости ' +
      'и соответствующая пара столбцов Пос./Оц. в Журнале. Действие необратимо.',
      ui.ButtonSet.YES_NO
    );

  if (
    answer !==
    ui.Button.YES
  ) {
    return false;
  }

  const active =
    getActiveLesson_();

  if (
    active &&
    String(active.lessonId) ===
      String(lessonId)
  ) {
    props_().deleteProperty(
      PROPS.ACTIVE_LESSON
    );

    props_().deleteProperty(
      PROPS.ACTIVE_CHECK
    );
  }

  const removedMarks =
    purgeMarksForLessonInSpreadsheet_(
      spreadsheet,
      lessonId
    );

  /*
   * Удаляем строку занятия из реестра только через скрипт,
   * затем Journal полностью перестраивается по оставшимся занятиям.
   */
  lessons.deleteRow(row);

  const result =
    rebuildJournalFromRegistryInSpreadsheet_(
      spreadsheet
    );

  ui.alert(
    'Занятие удалено.\n' +
    'Удалено технических отметок: ' +
    removedMarks +
    '.\n' +
    'В Журнале осталось занятий: ' +
    result.lessons +
    '.'
  );

  return true;
}

function getLessonIdFromJournalColumn_(spreadsheet, column) {
  const journal =
    spreadsheet.getSheetByName(
      SHEETS.JOURNAL
    );

  if (
    column < 3 ||
    column > journal.getLastColumn()
  ) {
    return '';
  }

  /*
   * В строке SUBHEADER_ROW обе колонки пары Пос./Оц.
   * содержат note вида lesson_id=...
   * Поэтому функция работает независимо от того,
   * выбрана Пос. или Оц. и независимо от строки выбранной ячейки.
   */
  const note =
    String(
      journal
        .getRange(
          JOURNAL_LAYOUT.SUBHEADER_ROW,
          column
        )
        .getNote() || ''
    ).trim();

  const match =
    note.match(
      /^lesson_id=(.+)$/
    );

  return match
    ? String(match[1]).trim()
    : '';
}

function deleteLessonFromJournalSelection_() {
  assertBoundSpreadsheetUi_();

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const range =
    spreadsheet.getActiveRange();

  if (
    !range ||
    range.getSheet().getName() !==
      SHEETS.JOURNAL
  ) {
    SpreadsheetApp.getUi().alert(
      'Сначала выделите любую ячейку в колонке Пос. или Оц. нужного занятия на листе «Журнал».'
    );
    return;
  }

  const lessonId =
    getLessonIdFromJournalColumn_(
      spreadsheet,
      range.getColumn()
    );

  if (!lessonId) {
    SpreadsheetApp.getUi().alert(
      'Выбранный столбец не относится к занятию. ' +
      'Выберите любую ячейку в паре Пос./Оц. нужного занятия.'
    );
    return;
  }

  deleteLessonByIdFromUi_(
    spreadsheet,
    lessonId
  );
}

function deleteSelectedLessonFromMenu() {
  assertBoundSpreadsheetUi_();

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const range =
    spreadsheet.getActiveRange();

  if (
    !range ||
    range.getSheet().getName() !==
      SHEETS.LESSONS ||
    range.getRow() < 2
  ) {
    SpreadsheetApp.getUi().alert(
      'Сначала выделите любую ячейку в строке занятия на листе «Занятия».'
    );
    return;
  }

  const lessons =
    spreadsheet.getSheetByName(
      SHEETS.LESSONS
    );

  const lessonId =
    String(
      lessons
        .getRange(
          range.getRow(),
          LESSON_COL.ID
        )
        .getValue() || ''
    ).trim();

  if (!lessonId) {
    SpreadsheetApp.getUi().alert(
      'В выбранной строке нет занятия.'
    );
    return;
  }

  deleteLessonByIdFromUi_(
    spreadsheet,
    lessonId
  );
}

function rebuildJournalFromMenu() {
  assertBoundSpreadsheetUi_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const result = rebuildJournalFromRegistryInSpreadsheet_(spreadsheet);
  SpreadsheetApp.getUi().alert(
    'Журнал перестроен. Студентов: ' + result.students +
    '; занятий: ' + result.lessons + '.'
  );
}

function prepareBlankTemplateFromMenu() {
  assertBoundSpreadsheetUi_();
  const ui = SpreadsheetApp.getUi();

  const answer = ui.alert(
    'Подготовить пустой шаблон?',
    'Будут удалены список студентов, занятия, отметки посещаемости/оценки, ' +
    'служебная история и URL развертывания. Типы занятий и общие настройки сохранятся.',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const students = spreadsheet.getSheetByName(SHEETS.STUDENTS);
  const lessons = spreadsheet.getSheetByName(SHEETS.LESSONS);
  const marks = spreadsheet.getSheetByName(SHEETS.MARKS);

  if (students.getMaxRows() > 1) {
    students.getRange(2, 1, students.getMaxRows() - 1, 3).clearContent();
  }
  if (lessons.getMaxRows() > 1) {
    lessons.getRange(2, 1, lessons.getMaxRows() - 1, LESSON_COL.CREATED).clearContent();
  }
  if (marks.getMaxRows() > 1) {
    marks.getRange(2, 1, marks.getMaxRows() - 1, 12).clearContent();
  }

  props_().deleteProperty(PROPS.ACTIVE_LESSON);
  props_().deleteProperty(PROPS.ACTIVE_CHECK);

  setSettingValueInSpreadsheet_(spreadsheet, 'discipline', 'Название дисциплины', 'Название дисциплины');
  setSettingValueInSpreadsheet_(spreadsheet, 'web_app_url', '', null);
  setSettingValueInSpreadsheet_(spreadsheet, 'student_url', '', null);
  setSettingValueInSpreadsheet_(spreadsheet, 'display_url', '', null);
  setSettingValueInSpreadsheet_(spreadsheet, 'teacher_url', '', null);
  setSettingValueInSpreadsheet_(spreadsheet, 'student_short_url', '', null);
  setSettingValueInSpreadsheet_(spreadsheet, 'display_short_url', '', null);
  setSettingValueInSpreadsheet_(spreadsheet, 'sheet_short_url', '', null);
  setSettingValueInSpreadsheet_(spreadsheet, 'teacher_key', newTeacherKey_(),
    'Секретный ключ пульта; создаётся автоматически для каждой копии');
  setSettingValueInSpreadsheet_(spreadsheet, 'instance_spreadsheet_id', spreadsheet.getId(),
    'ID экземпляра; меняется автоматически при копировании файла');
  setSettingValueInSpreadsheet_(spreadsheet, 'sheet_url', spreadsheet.getUrl(),
    'Прямая ссылка на этот журнал; обновляется автоматически');

  rebuildJournalFromRegistryInSpreadsheet_(spreadsheet);

  ui.alert(
    'Готово. Это чистый шаблон. Теперь его можно копировать для новой дисциплины.'
  );
}

