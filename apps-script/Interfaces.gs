/* -------------------- Состояние интерфейсов -------------------- */

function countMarked_(lessonId) {
  if (!lessonId) return 0;

  const sh = ss_().getSheetByName(SHEETS.MARKS);
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return 0;

  const rows = sh.getRange(
    2, 2, lastRow - 1, 2
  ).getDisplayValues();

  const unique = new Set();

  rows.forEach(([lid, studentNo]) => {
    if (String(lid) === String(lessonId)) {
      unique.add(String(studentNo));
    }
  });

  return unique.size;
}

function getTeacherState_() {
  const lesson = getActiveLesson_();
  const check = getCheckState_();
  const settings = getSettings_();

  return {
    discipline: String(settings.discipline || 'Дисциплина'),
    lesson,
    check,
    marked: lesson ? countMarked_(lesson.lessonId) : 0,
    total: getStudents_(true).length,
    lessonTypes: getLessonTypes_(true).map(t => ({ name: t.name, color: t.color, number: t.number })),
    autoAbsentOnFinish: settingBool_(settings.auto_absent_on_finish, true),
    absentPoints: Number(settings.attendance_absent_points || 0),
    links: getAppLinks_()
  };
}

function getDisplayState() {
  const lesson = getActiveLesson_();
  const current = getCheckState_();
  const check = current && current.active && current.kind === 'regular'
    ? current
    : { active: false };
  const links = getAppLinks_();

  return {
    discipline: String(getSettings_().discipline || 'Дисциплина'),
    lesson,
    check,
    marked: lesson ? countMarked_(lesson.lessonId) : 0,
    total: getStudents_(true).length,
    links: { studentUrl: links.studentUrl }
  };
}

