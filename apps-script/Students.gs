/* -------------------- Студенты -------------------- */

function getStudents_(activeOnly) {
  const sh = ss_().getSheetByName(SHEETS.STUDENTS);
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return [];

  return sh.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(r => r[0] !== '' && r[1] !== '')
    .map(r => ({
      no: String(r[0]),
      name: String(r[1]).trim(),
      active: r[2] === '' || r[2] == null ? true : Boolean(r[2])
    }))
    .filter(s => !activeOnly || s.active);
}

function searchStudents(query) {
  const settings = getSettings_();
  const minChars = Number(settings.student_search_min_chars || 3);
  const q = String(query || '').trim().toLocaleLowerCase('ru-RU');

  if (q.length < minChars) return [];

  return getStudents_(true)
    .filter(s => s.name.toLocaleLowerCase('ru-RU').includes(q))
    .slice(0, 10)
    .map(s => ({ no: s.no, name: s.name }));
}

function getStudentByNo(studentNo) {
  const no = String(studentNo || '');
  const student = getStudents_(true).find(s => s.no === no);

  return student
    ? { no: student.no, name: student.name }
    : null;
}


function syncStudents_() {
  return syncStudentsInSpreadsheet_(ss_());
}

function syncStudentsInSpreadsheet_(spreadsheet) {
  const source = spreadsheet.getSheetByName(SHEETS.STUDENTS);
  const journal = spreadsheet.getSheetByName(SHEETS.JOURNAL);

  const sourceLast = source.getLastRow();
  const students = [];

  if (sourceLast >= 2) {
    const sourceRows = source.getRange(2, 1, sourceLast - 1, 3).getValues();

    sourceRows.forEach((r, idx) => {
      if (r[0] === '' || r[1] === '') return;

      const row = idx + 2;
      const no = String(r[0]);
      const name = String(r[1]).trim();
      let active = r[2];

      // Пустая ячейка Active означает новый ввод: по умолчанию TRUE.
      // Осознанное FALSE больше не перезаписывается синхронизацией.
      if (active === '' || active == null) {
        source.getRange(row, 3).setValue(true);
        active = true;
      }

      students.push({
        row,
        no,
        name,
        active: Boolean(active)
      });
    });
  }

  const lastCol = Math.max(journal.getLastColumn(), 5);
  const oldLastRow = Math.max(journal.getLastRow(), JOURNAL_LAYOUT.SUBHEADER_ROW);
  const oldCount = Math.max(0, oldLastRow - JOURNAL_LAYOUT.SUBHEADER_ROW);

  const oldRows = oldCount
    ? journal.getRange(JOURNAL_LAYOUT.DATA_START_ROW, 1, oldCount, lastCol).getValues()
    : [];

  const oldByNo = new Map();
  oldRows.forEach(row => {
    if (row[0] !== '' && row[0] != null) {
      oldByNo.set(String(row[0]), row.slice());
    }
  });

  /*
   * В пустом шаблоне checkbox-ячейка может технически читаться как FALSE.
   * Если студент впервые появился в Журнале (например, список вставили сразу блоком),
   * считаем его активным по умолчанию и записываем TRUE.
   * Для уже существующего студента осознанное FALSE сохраняется.
   */
  students.forEach(student => {
    if (!oldByNo.has(student.no) && !student.active) {
      source.getRange(student.row, 3).setValue(true);
      student.active = true;
    }
  });

  const sourceNos = new Set(students.map(s => s.no));
  const removedNos = new Set();
  oldByNo.forEach((row, no) => {
    if (!sourceNos.has(no)) removedNos.add(no);
  });

  if (removedNos.size) {
    purgeMarksForStudentNosInSpreadsheet_(spreadsheet, removedNos);
  }

  const newRows = students.map(student => {
    const old = oldByNo.get(student.no);
    const row = old ? old.slice(0, lastCol) : Array(lastCol).fill('');
    while (row.length < lastCol) row.push('');

    const numericNo = Number(student.no);
    row[0] = Number.isFinite(numericNo) ? numericNo : student.no;
    row[1] = student.name;

    // У нового студента история прошлых занятий должна оставаться пустой.
    if (!old) {
      for (let c = 2; c < lastCol; c++) row[c] = '';
    }

    return row;
  });

  /*
   * Ключевое правило v2.3:
   * студенческий блок Журнала всегда начинается с DATA_START_ROW
   * и идёт без пустых промежутков. Остатки ручных удалений/старых нулей
   * больше не могут отправить студентов на 60-ю, 100-ю и т.п. строку.
   */
  const rowsToClear = Math.max(oldCount, students.length);
  if (rowsToClear > 0) {
    journal
      .getRange(JOURNAL_LAYOUT.DATA_START_ROW, 1, rowsToClear, lastCol)
      .clearContent();
  }

  if (newRows.length) {
    journal
      .getRange(JOURNAL_LAYOUT.DATA_START_ROW, 1, newRows.length, lastCol)
      .setValues(newRows);
  }

  // Итоговые клетки пустых строк должны быть пустыми, а не 0 0 0.
  recomputeAllTotalsInSpreadsheet_(spreadsheet);
  clearPhantomJournalTotals_(spreadsheet);

  let added = 0;
  let renamed = 0;
  students.forEach(student => {
    const old = oldByNo.get(student.no);
    if (!old) added++;
    else if (String(old[1] || '').trim() !== student.name) renamed++;
  });

  return {
    total: students.length,
    active: students.filter(s => s.active).length,
    added,
    removed: removedNos.size,
    renamed
  };
}



function purgeMarksForLessonInSpreadsheet_(spreadsheet, lessonId) {
  const marks =
    spreadsheet.getSheetByName(
      SHEETS.MARKS
    );

  const lastRow =
    marks.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  const ids =
    marks
      .getRange(
        2,
        2,
        lastRow - 1,
        1
      )
      .getDisplayValues()
      .flat();

  const rows = [];

  ids.forEach((id, idx) => {
    if (
      String(id) ===
      String(lessonId)
    ) {
      rows.push(idx + 2);
    }
  });

  rows
    .sort((a, b) => b - a)
    .forEach(row => {
      marks.deleteRow(row);
    });

  return rows.length;
}

function purgeMarksForStudentNosInSpreadsheet_(
  spreadsheet,
  studentNos
) {
  if (
    !studentNos ||
    !studentNos.size
  ) {
    return 0;
  }

  const marks =
    spreadsheet.getSheetByName(
      SHEETS.MARKS
    );

  const lastRow =
    marks.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  /*
   * В Отметки studentNo находится
   * в колонке C.
   */
  const ids =
    marks
      .getRange(
        2,
        3,
        lastRow - 1,
        1
      )
      .getDisplayValues()
      .flat();

  const rows = [];

  ids.forEach((id, idx) => {
    if (
      studentNos.has(
        String(id)
      )
    ) {
      rows.push(idx + 2);
    }
  });

  rows
    .sort((a, b) => b - a)
    .forEach(row => {
      marks.deleteRow(row);
    });

  return rows.length;
}

function getJournalRowByStudentNo_(studentNo) {
  const sh = ss_().getSheetByName(SHEETS.JOURNAL);
  const lastRow = sh.getLastRow();
  if (lastRow < JOURNAL_LAYOUT.DATA_START_ROW) return null;

  const ids = sh.getRange(
    JOURNAL_LAYOUT.DATA_START_ROW,
    1,
    lastRow - JOURNAL_LAYOUT.DATA_START_ROW + 1,
    1
  ).getDisplayValues().flat();

  const idx = ids.indexOf(String(studentNo));
  return idx >= 0 ? idx + JOURNAL_LAYOUT.DATA_START_ROW : null;
}

function recomputeAllTotals_() {
  recomputeAllTotalsInSpreadsheet_(ss_());
}


function clearPhantomJournalTotals_(spreadsheet) {
  const sh =
    spreadsheet.getSheetByName(
      SHEETS.JOURNAL
    );

  const lastCol =
    sh.getLastColumn();

  if (lastCol < 3) {
    return;
  }

  const meta =
    sh
      .getRange(
        JOURNAL_LAYOUT.META_ROW,
        1,
        1,
        lastCol
      )
      .getDisplayValues()[0];

  const attCol =
    meta.indexOf(
      'Баллы за посещение'
    ) + 1;

  if (!attCol) {
    return;
  }

  /*
   * В строке подзаголовков итоговые клетки всегда должны быть пустыми.
   */
  sh
    .getRange(
      JOURNAL_LAYOUT.SUBHEADER_ROW,
      attCol,
      1,
      3
    )
    .clearContent();

  const lastRow =
    sh.getLastRow();

  if (
    lastRow <
    JOURNAL_LAYOUT.DATA_START_ROW
  ) {
    return;
  }

  const ids =
    sh
      .getRange(
        JOURNAL_LAYOUT.DATA_START_ROW,
        1,
        lastRow -
          JOURNAL_LAYOUT.DATA_START_ROW +
          1,
        2
      )
      .getDisplayValues();

  ids.forEach((r, idx) => {
    const no =
      String(r[0] || '').trim();

    const name =
      String(r[1] || '').trim();

    if (!no || !name) {
      sh
        .getRange(
          JOURNAL_LAYOUT.DATA_START_ROW +
            idx,
          attCol,
          1,
          3
        )
        .clearContent();
    }
  });
}

function recomputeAllTotalsInSpreadsheet_(spreadsheet) {
  const sh = spreadsheet.getSheetByName(SHEETS.JOURNAL);
  const lastRow = sh.getLastRow();
  if (lastRow < JOURNAL_LAYOUT.DATA_START_ROW) return;

  const ids = sh.getRange(
    JOURNAL_LAYOUT.DATA_START_ROW,
    1,
    lastRow - JOURNAL_LAYOUT.DATA_START_ROW + 1,
    1
  ).getDisplayValues().flat();

  ids.forEach((id, idx) => {
    const row = JOURNAL_LAYOUT.DATA_START_ROW + idx;
    recomputeRowTotalsInSpreadsheet_(spreadsheet, row);
  });
}

function recomputeRowTotals_(row) {
  recomputeRowTotalsInSpreadsheet_(ss_(), row);
}

function recomputeRowTotalsInSpreadsheet_(spreadsheet, row) {
  const sh =
    spreadsheet.getSheetByName(
      SHEETS.JOURNAL
    );

  const lastCol =
    sh.getLastColumn();

  if (
    lastCol < 5 ||
    row <
      JOURNAL_LAYOUT.DATA_START_ROW
  ) {
    return;
  }

  const metaHeaders =
    sh
      .getRange(
        JOURNAL_LAYOUT.META_ROW,
        1,
        1,
        lastCol
      )
      .getDisplayValues()[0];

  const subHeaders =
    sh
      .getRange(
        JOURNAL_LAYOUT.SUBHEADER_ROW,
        1,
        1,
        lastCol
      )
      .getDisplayValues()[0];

  const attCol =
    metaHeaders.indexOf(
      'Баллы за посещение'
    ) + 1;

  const gradeCol =
    metaHeaders.indexOf(
      'Баллы за работу'
    ) + 1;

  const totalCol =
    metaHeaders.indexOf(
      'Итого'
    ) + 1;

  if (
    !attCol ||
    !gradeCol ||
    !totalCol
  ) {
    return;
  }

  const studentNo =
    sh.getRange(row, 1).getValue();

  const studentName =
    sh.getRange(row, 2).getValue();

  if (
    studentNo === '' ||
    studentNo == null ||
    studentName === '' ||
    studentName == null
  ) {
    sh
      .getRange(
        row,
        attCol,
        1,
        3
      )
      .clearContent();

    return;
  }

  let attendance = 0;
  let grades = 0;
  let hasAttendanceData = false;
  let hasGradeData = false;

  if (attCol > 3) {
    const values =
      sh
        .getRange(
          row,
          3,
          1,
          attCol - 3
        )
        .getValues()[0];

    values.forEach((value, idx) => {
      if (
        value === '' ||
        value == null
      ) {
        return;
      }

      const n =
        Number(value);

      if (
        !Number.isFinite(n)
      ) {
        return;
      }

      const header =
        String(
          subHeaders[idx + 2] || ''
        );

      if (
        header === 'Пос.'
      ) {
        attendance += n;
        hasAttendanceData = true;
      }

      if (
        header === 'Оц.'
      ) {
        grades += n;
        hasGradeData = true;
      }
    });
  }

  /*
   * Ключевое различие:
   * пусто = данных ещё нет;
   * 0 = есть явная нулевая запись (например отсутствие).
   */
  const attendanceOut =
    hasAttendanceData
      ? attendance
      : '';

  const gradesOut =
    hasGradeData
      ? grades
      : '';

  const totalOut =
    (
      hasAttendanceData ||
      hasGradeData
    )
      ? attendance + grades
      : '';

  sh
    .getRange(
      row,
      attCol,
      1,
      3
    )
    .setValues([[
      attendanceOut,
      gradesOut,
      totalOut
    ]]);
}


