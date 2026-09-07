/* -------------------- Динамические коды -------------------- */

function stableCode_(seed, cycle, digits) {
  const n = Math.max(1, Number(digits || 4));
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(seed) + ':' + String(cycle)
  );

  let value = 0;
  for (let i = 0; i < 6; i++) {
    value = (value * 256 + (digest[i] & 0xff)) % Math.pow(10, n);
  }

  return String(value).padStart(n, '0');
}

function getActiveCheckRaw_() {
  const raw = props_().getProperty(PROPS.ACTIVE_CHECK);
  return raw ? JSON.parse(raw) : null;
}

function startCheck_(kind) {
  const lesson = getActiveLesson_();

  if (!lesson) {
    return {
      ok: false,
      message: 'Сначала начните занятие.'
    };
  }

  kind = kind === 'late' ? 'late' : 'regular';
  const settings = getSettings_();

  const seconds = Number(
    kind === 'late'
      ? settings.late_code_seconds
      : settings.regular_code_seconds
  ) || 60;

  const check = {
    checkId: 'C-' + Utilities.getUuid().slice(0, 10),
    lessonId: lesson.lessonId,
    kind,
    seconds,
    startMs: Date.now(),
    seed: Utilities.getUuid()
  };

  props_().setProperty(
    PROPS.ACTIVE_CHECK,
    JSON.stringify(check)
  );

  const lessonRow = findLessonRow_(lesson.lessonId);

  if (lessonRow) {
    const lessons = ss_().getSheetByName(SHEETS.LESSONS);
    const oldCount = Number(
      lessons.getRange(lessonRow, LESSON_COL.CHECK_COUNT).getValue() || 0
    );
    lessons.getRange(lessonRow, LESSON_COL.CHECK_COUNT).setValue(oldCount + 1);
  }

  return {
    ok: true,
    message:
      kind === 'late'
        ? 'Генерация кодов для опоздавшего запущена.'
        : 'Генерация кодов присутствия запущена.',
    state: getTeacherState_()
  };
}

function getCheckState_() {
  const check = getActiveCheckRaw_();

  if (!check) {
    return { active: false };
  }

  const settings = getSettings_();
  const digits = Number(settings.code_digits || 4);
  const cycleMs = Number(check.seconds || 60) * 1000;
  const elapsed = Math.max(0, Date.now() - Number(check.startMs));
  const cycle = Math.floor(elapsed / cycleMs);

  const expiresAt =
    Number(check.startMs) + (cycle + 1) * cycleMs;

  return {
    active: true,
    checkId: check.checkId,
    lessonId: check.lessonId,
    kind: check.kind,
    code: stableCode_(check.seed, cycle, digits),
    nextCode: stableCode_(check.seed, cycle + 1, digits),
    cycle: cycle + 1,
    seconds: Number(check.seconds || 60),
    expiresAt
  };
}

function stopCheck_() {
  props_().deleteProperty(PROPS.ACTIVE_CHECK);

  return {
    ok: true,
    state: getTeacherState_()
  };
}

