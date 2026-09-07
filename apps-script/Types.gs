/* -------------------- Типы занятий -------------------- */

function getLessonTypes_(activeOnly) {
  const sh = ss_().getSheetByName(SHEETS.TYPES);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh.getRange(2, 1, lastRow - 1, 4).getValues()
    .map((r, idx) => ({
      row: idx + 2,
      name: String(r[0] || '').trim(),
      color: /^#[0-9A-Fa-f]{6}$/.test(String(r[1] || '').trim())
        ? String(r[1]).trim()
        : '#EDEDED',
      number: settingBool_(r[2], true),
      active: settingBool_(r[3], false)
    }))
    .filter(t => t.name && (!activeOnly || t.active));
}

function getLessonType_(name) {
  const all = getLessonTypes_(false);
  const found = all.find(t => t.name === String(name || '').trim());
  return found || { name: String(name || 'Другое'), color: '#EDEDED', number: false, active: true };
}

function getDefaultLessonType_() {
  const active = getLessonTypes_(true);
  return active.length ? active[0].name : 'Лекция';
}

function nextLessonTypeNumber_(typeName) {
  const config = getLessonType_(typeName);
  if (!config.number) return '';

  const sh = ss_().getSheetByName(SHEETS.LESSONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;

  const types = sh.getRange(2, LESSON_COL.TYPE, lastRow - 1, 1).getDisplayValues().flat();
  return types.filter(v => String(v) === String(typeName)).length + 1;
}

function formatLessonMetaLabel_(dateDisplay, typeName, typeNumber, topic) {
  let label = String(dateDisplay || '');
  if (typeName) {
    label += (label ? ' · ' : '') + typeName;
    if (typeNumber !== '' && typeNumber != null) label += ' №' + typeNumber;
  }
  const cleanTopic = String(topic || '').trim();
  if (cleanTopic) label += '\n' + cleanTopic;
  return label;
}

function recalculateLessonTypeNumbers_() {
  const sh = ss_().getSheetByName(SHEETS.LESSONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rows = sh.getRange(2, 1, lastRow - 1, LESSON_COL.CREATED).getValues();
  const counters = {};

  rows.forEach((r, idx) => {
    const typeName = String(r[LESSON_COL.TYPE - 1] || '').trim();
    if (!typeName) return;
    const config = getLessonType_(typeName);
    let num = '';
    if (config.number) {
      counters[typeName] = (counters[typeName] || 0) + 1;
      num = counters[typeName];
    }
    sh.getRange(idx + 2, LESSON_COL.TYPE_NUMBER).setValue(num);
  });
}

function refreshAllLessonJournalMetadata_() {
  recalculateLessonTypeNumbers_();
  const sh = ss_().getSheetByName(SHEETS.LESSONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  for (let row = 2; row <= lastRow; row++) refreshLessonJournalMetadataByRow_(row);
}

function refreshLessonJournalMetadataByRow_(row) {
  const lessons = ss_().getSheetByName(SHEETS.LESSONS);
  if (row < 2 || row > lessons.getLastRow()) return;

  const vals = lessons.getRange(row, 1, 1, LESSON_COL.CREATED).getValues()[0];
  const lessonId = String(vals[LESSON_COL.ID - 1] || '').trim();
  if (!lessonId) return;

  const date = vals[LESSON_COL.DATE - 1];
  const typeName = String(vals[LESSON_COL.TYPE - 1] || '').trim();
  const typeNumber = vals[LESSON_COL.TYPE_NUMBER - 1];
  const topic = String(vals[LESSON_COL.TOPIC - 1] || '').trim();
  const attendanceCol = Number(vals[LESSON_COL.ATTENDANCE_COL - 1] || 0);
  const gradeCol = Number(vals[LESSON_COL.GRADE_COL - 1] || 0);

  if (!attendanceCol || !gradeCol) return;

  const tz = ss_().getSpreadsheetTimeZone();
  const dateDisplay = date instanceof Date
    ? Utilities.formatDate(date, tz, 'dd.MM')
    : String(date || '');
  const config = getLessonType_(typeName);
  const journal = ss_().getSheetByName(SHEETS.JOURNAL);

  const metaRange = journal.getRange(JOURNAL_LAYOUT.META_ROW, attendanceCol, 1, 2);
  if (!metaRange.isPartOfMerge()) metaRange.merge();
  metaRange.setValue(formatLessonMetaLabel_(dateDisplay, typeName, typeNumber, topic))
    .setBackground(config.color)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setNote('lesson_id=' + lessonId);

  journal.getRange(JOURNAL_LAYOUT.SUBHEADER_ROW, attendanceCol, 1, 2)
    .setValues([['Пос.', 'Оц.']])
    .setBackground(config.color)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNotes([['lesson_id=' + lessonId, 'lesson_id=' + lessonId]]);

  const active = getActiveLesson_();
  if (active && String(active.lessonId) === lessonId) {
    active.lessonType = typeName;
    active.typeNumber = typeNumber;
    active.topic = topic;
    saveActiveLesson_(active);
  }
}


