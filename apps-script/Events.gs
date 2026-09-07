function onEdit(e) {
  if (!e || !e.range || !e.source) return;

  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const name = sh.getName();

  if (name === SHEETS.STUDENTS && row >= 2 && col <= 3) {
    if (col <= 2) {
      const no = sh.getRange(row, 1).getValue();
      const studentName = sh.getRange(row, 2).getValue();
      const activeCell = sh.getRange(row, 3);
      const active = activeCell.getValue();
      const editedWasEmpty = e.oldValue === undefined || e.oldValue === null || e.oldValue === '';
      if (no !== '' && studentName !== '' && editedWasEmpty && active !== true) {
        activeCell.setValue(true);
      }
    }
    syncStudentsInSpreadsheet_(e.source);
    return;
  }

  if (name === SHEETS.JOURNAL && row >= JOURNAL_LAYOUT.DATA_START_ROW && col >= 3) {
    recomputeRowTotalsInSpreadsheet_(e.source, row);
    return;
  }

  if (name === SHEETS.LESSONS && row >= 2) {
    const lessonId = String(sh.getRange(row, LESSON_COL.ID).getValue() || '').trim();

    // Если запись занятия очищена вручную, Журнал должен перестать хранить её колонки.
    if (!lessonId) {
      rebuildJournalFromRegistryInSpreadsheet_(e.source);
      return;
    }

    // Тип и тема меняют подпись/цвет журнала; заметка остаётся в реестре занятий.
    if ([LESSON_COL.TYPE, LESSON_COL.TOPIC].includes(col)) {
      recalculateLessonTypeNumbers_();
      refreshAllLessonJournalMetadata_();
    }

    if ([LESSON_COL.TYPE, LESSON_COL.TOPIC, LESSON_COL.NOTES].includes(col)) {
      const active = getActiveLesson_();
      if (active && String(active.lessonId) === lessonId) {
        active.lessonType = String(sh.getRange(row, LESSON_COL.TYPE).getValue() || '');
        active.typeNumber = sh.getRange(row, LESSON_COL.TYPE_NUMBER).getValue();
        active.topic = String(sh.getRange(row, LESSON_COL.TOPIC).getValue() || '');
        active.notes = String(sh.getRange(row, LESSON_COL.NOTES).getValue() || '');
        saveActiveLesson_(active);
      }
    }
    return;
  }

  if (name === SHEETS.TYPES && row >= 2) {
    const typeName = String(sh.getRange(row, 1).getValue() || '').trim();
    if (typeName) {
      const colorCell = sh.getRange(row, 2);
      const numberCell = sh.getRange(row, 3);
      const activeCell = sh.getRange(row, 4);
      if (!String(colorCell.getValue() || '').trim()) colorCell.setValue('#EDEDED');
      if (numberCell.getValue() === '') numberCell.setValue(true);
      if (activeCell.getValue() === '') activeCell.setValue(true);
      const color = String(colorCell.getValue() || '#EDEDED').trim();
      if (/^#[0-9A-Fa-f]{6}$/.test(color)) colorCell.setBackground(color);
    }
    refreshAllLessonJournalMetadata_();
    return;
  }

  if (name === SHEETS.SETTINGS && row >= 2) {
    const key = String(sh.getRange(row, 1).getValue() || '').trim();
    if (key === 'discipline') updateJournalTitleInSpreadsheet_(e.source);
  }
}

function isTeacherKeyValid_(key) {
  const expected = String(getSettings_().teacher_key || '').trim();
  return !!expected && String(key || '') === expected;
}

function assertTeacherKey_(key) {
  if (!isTeacherKeyValid_(key)) {
    throw new Error('Недействительный ключ преподавателя.');
  }
}

function teacherGetState(key) {
  assertTeacherKey_(key);

  /*
   * v2.4: НЕ вызываем syncStudents_() на каждом polling-запросе.
   * Раньше это могло перезаписать только что поставленную отметку,
   * если polling и submitAttendance выполнялись одновременно.
   * Состав синхронизируется через onEdit, явную команду,
   * а также при начале/завершении занятия.
   */
  return getTeacherState_();
}

function teacherStartLesson(key, meta) {
  assertTeacherKey_(key);
  return startLesson_(meta);
}

function teacherStartCheck(key, kind) {
  assertTeacherKey_(key);
  return startCheck_(kind);
}

function teacherStopCheck(key) {
  assertTeacherKey_(key);
  return stopCheck_();
}

function teacherFinishLesson(key) {
  assertTeacherKey_(key);
  const result = finishLesson_();
  result.state = getTeacherState_();
  return result;
}

function teacherSyncStudents(key) {
  assertTeacherKey_(key);
  const result = syncStudents_();

  return {
    ok: true,
    message:
      'Состав синхронизирован. Активных: ' +
      result.active +
      ', всего: ' +
      result.total +
      ', добавлено: ' +
      result.added +
      ', удалено: ' +
      result.removed +
      '.',
    state: getTeacherState_()
  };
}

