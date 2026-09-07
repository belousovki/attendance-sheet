/* -------------------- Занятие -------------------- */

function getActiveLesson_() {
  const raw = props_().getProperty(PROPS.ACTIVE_LESSON);
  return raw ? JSON.parse(raw) : null;
}

function saveActiveLesson_(lesson) {
  props_().setProperty(PROPS.ACTIVE_LESSON, JSON.stringify(lesson));
}

/**
 * mode нужен только как пометка "Очно / Дистанционно".
 * На логику посещаемости он не влияет.
 */
function startLesson_(meta) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    syncStudents_();

    const existing = getActiveLesson_();
    if (existing) {
      return { ok: false, message: 'Уже есть активное занятие.', state: getTeacherState_() };
    }

    if (typeof meta === 'string') meta = { mode: meta };
    meta = meta || {};

    const mode = meta.mode === 'Дистанционно' ? 'Дистанционно' : 'Очно';
    const lessonType = String(meta.lessonType || getDefaultLessonType_()).trim();
    const topic = String(meta.topic || '').trim();
    const notes = String(meta.notes || '').trim();
    const typeConfig = getLessonType_(lessonType);
    const typeNumber = nextLessonTypeNumber_(lessonType);

    const now = new Date();
    const tz = ss_().getSpreadsheetTimeZone();
    const dateIso = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const dateDisplay = Utilities.formatDate(now, tz, 'dd.MM');
    const lessonId = 'L-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 6);

    const journal = createLessonColumns_(
      dateDisplay,
      lessonId,
      lessonType,
      typeNumber,
      topic,
      typeConfig.color
    );

    const lessons = ss_().getSheetByName(SHEETS.LESSONS);
    lessons.appendRow([
      lessonId,
      now,
      mode,
      lessonType,
      typeNumber,
      topic,
      'active',
      now,
      '',
      journal.attendanceCol,
      journal.gradeCol,
      0,
      notes,
      now
    ]);

    /*
     * appendRow() may apply the sheet's automatic date-time format instead
     * of the format prepared on an otherwise empty row. Format the actual
     * registry row after it has been appended.
     */
    const lessonRow = lessons.getLastRow();
    lessons.getRange(lessonRow, LESSON_COL.DATE)
      .setNumberFormat('dd.MM.yyyy');
    [
      LESSON_COL.STARTED,
      LESSON_COL.ENDED,
      LESSON_COL.CREATED
    ].forEach(column => {
      lessons.getRange(lessonRow, column)
        .setNumberFormat('dd.MM.yyyy HH:mm:ss');
    });

    /*
     * alpha.8:
     * Шкала оценок строится только после записи занятия в реестр.
     * Иначе getLessonScoreRanges_() ещё не знает номера новых
     * колонок Пос./Оц. и условные правила не создаются.
     */
    applyScoreScaleFormatting_(ss_());

    const lesson = {
      lessonId,
      dateIso,
      dateDisplay,
      mode,
      lessonType,
      typeNumber,
      topic,
      notes,
      status: 'active',
      startedAt: now.getTime(),
      attendanceCol: journal.attendanceCol,
      gradeCol: journal.gradeCol
    };

    saveActiveLesson_(lesson);
    props_().deleteProperty(PROPS.ACTIVE_CHECK);

    return {
      ok: true,
      message: 'Занятие начато.',
      typeNumber,
      state: getTeacherState_()
    };
  } finally {
    lock.releaseLock();
  }
}

function createLessonColumns_(dateDisplay, lessonId, lessonType, typeNumber, topic, color) {
  const sh = ss_().getSheetByName(SHEETS.JOURNAL);
  const metaRow = JOURNAL_LAYOUT.META_ROW;
  const subRow = JOURNAL_LAYOUT.SUBHEADER_ROW;

  let lastCol = Math.max(sh.getLastColumn(), 5);
  let headers = sh.getRange(metaRow, 1, 1, lastCol).getDisplayValues()[0];
  let summaryCol = headers.indexOf('Баллы за посещение') + 1;

  if (!summaryCol) {
    summaryCol = lastCol + 1;
    sh.getRange(metaRow, summaryCol, 1, 3)
      .setValues([['Баллы за посещение', 'Баллы за работу', 'Итого']]);
  }

  sh.insertColumnsBefore(summaryCol, 2);

  const label = formatLessonMetaLabel_(dateDisplay, lessonType, typeNumber, topic);
  const bg = /^#[0-9A-Fa-f]{6}$/.test(String(color || '')) ? color : '#EDEDED';

  const metaRange = sh.getRange(metaRow, summaryCol, 1, 2);
  metaRange.merge();
  metaRange.setValue(label)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground(bg)
    .setNote('lesson_id=' + lessonId);

  sh.getRange(subRow, summaryCol, 1, 2)
    .setValues([['Пос.', 'Оц.']])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(bg)
    .setNotes([['lesson_id=' + lessonId, 'lesson_id=' + lessonId]]);

  sh.setColumnWidth(summaryCol, 78);
  sh.setColumnWidth(summaryCol + 1, 78);
  sh.setRowHeight(metaRow, Math.max(sh.getRowHeight(metaRow), 42));

  updateSummaryFormulas_();

  return {
    attendanceCol: summaryCol,
    gradeCol: summaryCol + 1,
    attendanceHeader: dateDisplay + ' Пос.',
    gradeHeader: dateDisplay + ' Оц.'
  };
}

function updateSummaryFormulas_() {
  recomputeAllTotals_();
}

function finishLesson_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const lesson = getActiveLesson_();
    if (!lesson) return { ok: false, message: 'Активного занятия нет.' };

    syncStudents_();

    const settings = getSettings_();
    const absentPoints = Number(settings.attendance_absent_points || 0);
    const autoAbsent = settingBool_(settings.auto_absent_on_finish, true);
    let autoFilled = 0;

    if (autoAbsent) {
      const journal = ss_().getSheetByName(SHEETS.JOURNAL);
      const marks = ss_().getSheetByName(SHEETS.MARKS);
      const attendanceCol =
        getAttendanceColumnForLesson_(lesson.lessonId) ||
        Number(lesson.attendanceCol);
      const absentRows = [];
      const now = new Date();

      getStudents_(true).forEach(student => {
        const row = getJournalRowByStudentNo_(student.no);
        if (!row) return;

        const cell = journal.getRange(row, attendanceCol);
        const value = cell.getValue();

        if (value === '' || value == null) {
          cell.setValue(absentPoints);
          autoFilled++;

          absentRows.push([
            'M-' + Utilities.getUuid().slice(0, 12),
            lesson.lessonId,
            student.no,
            student.name,
            'auto_finish',
            '',
            now,
            absentPoints,
            'absent',
            '',
            'auto_finish',
            ''
          ]);
        }

        recomputeRowTotals_(row);
      });

      if (absentRows.length) {
        const start = marks.getLastRow() + 1;
        marks.getRange(start, 1, absentRows.length, 12).setValues(absentRows);
      }

      SpreadsheetApp.flush();
    }

    const lessonRow = findLessonRow_(lesson.lessonId);
    if (lessonRow) {
      const lessons = ss_().getSheetByName(SHEETS.LESSONS);
      lessons.getRange(lessonRow, LESSON_COL.STATUS).setValue('closed');
      lessons.getRange(lessonRow, LESSON_COL.ENDED).setValue(new Date());
    }

    props_().deleteProperty(PROPS.ACTIVE_CHECK);
    props_().deleteProperty(PROPS.ACTIVE_LESSON);

    return {
      ok: true,
      autoAbsent,
      autoFilled,
      absentPoints,
      message: autoAbsent
        ? 'Занятие завершено. Без отметки: ' + autoFilled + '; записано ' + absentPoints + '.'
        : 'Занятие завершено. Пустые отметки оставлены без изменений.'
    };
  } finally {
    lock.releaseLock();
  }
}

function findLessonRow_(lessonId) {
  const sh = ss_().getSheetByName(SHEETS.LESSONS);
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return null;

  const ids = sh.getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues()
    .flat();

  const idx = ids.indexOf(String(lessonId));
  return idx >= 0 ? idx + 2 : null;
}
