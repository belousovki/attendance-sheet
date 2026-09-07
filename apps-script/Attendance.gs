/* -------------------- Отметка студента -------------------- */

function submitAttendance(studentNo, code, clientToken) {
  const lesson = getActiveLesson_();

  if (!lesson) {
    return {
      ok: false,
      message: 'Сейчас нет активного занятия.'
    };
  }

  const state = getCheckState_();

  if (!state.active) {
    return {
      ok: false,
      message: 'Проверка присутствия сейчас не проводится.'
    };
  }

  const inputCode =
    String(code || '').replace(/\D/g, '');

  let codeOk = inputCode === state.code;

  if (!codeOk) {
    const rawCheck = getActiveCheckRaw_();

    if (rawCheck) {
      const cycleMs =
        Number(rawCheck.seconds || 60) * 1000;

      const elapsed = Math.max(
        0,
        Date.now() - Number(rawCheck.startMs)
      );

      const cycle =
        Math.floor(elapsed / cycleMs);

      const intoCycle =
        elapsed - cycle * cycleMs;

      if (cycle > 0 && intoCycle <= 5000) {
        const digits =
          Number(getSettings_().code_digits || 4);

        const previousCode = stableCode_(
          rawCheck.seed,
          cycle - 1,
          digits
        );

        codeOk =
          inputCode === previousCode;
      }
    }
  }

  if (!codeOk) {
    return {
      ok: false,
      message: 'Код неверный или уже сменился.'
    };
  }

  const student =
    getStudentByNo(studentNo);

  if (!student) {
    return {
      ok: false,
      message:
        'Студент не найден в активном составе группы.'
    };
  }

  if (
    alreadyMarked_(
      lesson.lessonId,
      student.no
    )
  ) {
    ensureJournalAttendanceFromLatestMark_(
      lesson.lessonId,
      student.no
    );

    return {
      ok: true,
      already: true,
      message:
        '✓ Присутствие уже отмечено.',
      student: student.name
    };
  }

  let journalRow =
    getJournalRowByStudentNo_(student.no);

  if (!journalRow) {
    syncStudents_();
    journalRow =
      getJournalRowByStudentNo_(student.no);
  }

  if (!journalRow) {
    return {
      ok: false,
      message:
        'Не удалось создать строку студента в журнале.'
    };
  }

  const settings = getSettings_();
  const isLate =
    state.kind === 'late';

  const points = Number(
    isLate
      ? settings.late_points
      : settings.attendance_present_points
  ) || 0;

  const status =
    isLate ? 'late' : 'present';

  const now = new Date();

  const marks =
    ss_().getSheetByName(SHEETS.MARKS);

  marks.appendRow([
    'M-' +
      Utilities.getUuid().slice(0, 12),
    lesson.lessonId,
    student.no,
    student.name,
    state.kind,
    state.cycle,
    now,
    points,
    status,
    String(clientToken || '')
      .slice(0, 100),
    'web',
    ''
  ]);

  writeJournalAttendance_(
    student.no,
    points,
    isLate,
    now
  );

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: isLate
      ? '✓ Отметка принята: опоздание зафиксировано.'
      : '✓ Присутствие отмечено.',
    student: student.name,
    points
  };
}


function getAttendanceStatus(studentNo) {
  const lesson = getActiveLesson_();
  const student = getStudentByNo(studentNo);

  if (!lesson || !student) {
    return { marked: false };
  }

  const marked = alreadyMarked_(lesson.lessonId, student.no);

  if (marked) {
    ensureJournalAttendanceFromLatestMark_(lesson.lessonId, student.no);
  }

  return {
    marked,
    student: student.name,
    message: marked
      ? '✓ Присутствие отмечено.'
      : 'Отметка пока не найдена.'
  };
}


function getLatestAttendanceMark_(lessonId, studentNo) {
  const sh = ss_().getSheetByName(SHEETS.MARKS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const rows = sh.getRange(2, 1, lastRow - 1, 12).getValues();

  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (
      String(r[1]) === String(lessonId) &&
      String(r[2]) === String(studentNo)
    ) {
      return {
        markId: r[0],
        lessonId: r[1],
        studentNo: r[2],
        studentName: r[3],
        checkType: r[4],
        submittedAt: r[6],
        points: r[7],
        status: String(r[8] || ''),
        source: String(r[10] || '')
      };
    }
  }

  return null;
}

function getLatestAttendanceMarksMap_() {
  const sh = ss_().getSheetByName(SHEETS.MARKS);
  const lastRow = sh.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const rows = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  rows.forEach(r => {
    const lessonId = String(r[1] || '').trim();
    const studentNo = String(r[2] || '').trim();
    if (!lessonId || !studentNo) return;

    map.set(studentNo + '\u0001' + lessonId, {
      points: r[7],
      status: String(r[8] || ''),
      submittedAt: r[6],
      source: String(r[10] || '')
    });
  });

  return map;
}

function getAttendanceColumnForLesson_(lessonId) {
  const row = findLessonRow_(lessonId);
  if (!row) return null;

  const lessons = ss_().getSheetByName(SHEETS.LESSONS);
  const col = Number(
    lessons.getRange(row, LESSON_COL.ATTENDANCE_COL).getValue()
  );

  return Number.isFinite(col) && col >= 3 ? col : null;
}

function ensureJournalAttendanceFromLatestMark_(lessonId, studentNo) {
  const mark = getLatestAttendanceMark_(lessonId, studentNo);
  if (!mark) return false;

  const row = getJournalRowByStudentNo_(studentNo);
  const col = getAttendanceColumnForLesson_(lessonId);
  if (!row || !col) return false;

  const points = mark.points;
  if (points === '' || points == null || !Number.isFinite(Number(points))) {
    return false;
  }

  const journal = ss_().getSheetByName(SHEETS.JOURNAL);
  const cell = journal.getRange(row, col);
  cell.setValue(Number(points));

  if (mark.status === 'late' && mark.submittedAt instanceof Date) {
    const tz = ss_().getSpreadsheetTimeZone();
    const time = Utilities.formatDate(mark.submittedAt, tz, 'HH:mm');
    cell.setNote('Опоздание. Отметка: ' + time);
  } else {
    cell.clearNote();
  }

  recomputeRowTotals_(row);
  SpreadsheetApp.flush();
  return true;
}

function alreadyMarked_(lessonId, studentNo) {
  const mark = getLatestAttendanceMark_(lessonId, studentNo);
  if (!mark) return false;

  return !['cleared', 'cancelled'].includes(mark.status);
}

function writeJournalAttendance_(studentNo, points, isLate, when) {
  const lesson = getActiveLesson_();
  if (!lesson) {
    throw new Error('Нет активного занятия.');
  }

  const sh = ss_().getSheetByName(SHEETS.JOURNAL);
  const row = getJournalRowByStudentNo_(studentNo);
  if (!row) {
    throw new Error('Студент не найден в журнале.');
  }

  const attendanceCol =
    getAttendanceColumnForLesson_(lesson.lessonId) ||
    Number(lesson.attendanceCol);

  if (!attendanceCol || attendanceCol < 3) {
    throw new Error('Не найдена колонка посещаемости занятия.');
  }

  const cell = sh.getRange(row, attendanceCol);
  cell.setValue(points);

  if (isLate) {
    const tz = ss_().getSpreadsheetTimeZone();
    const time = Utilities.formatDate(when, tz, 'HH:mm');
    cell.setNote('Опоздание. Отметка: ' + time);
  } else {
    cell.clearNote();
  }

  recomputeRowTotals_(row);
}

