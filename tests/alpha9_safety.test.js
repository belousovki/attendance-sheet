const fs = require('fs');
const path = require('path');
const assert = require('assert');

const root = path.resolve(__dirname, '..');
const appsDir = path.join(root, 'apps-script');

function read(name) {
  return fs.readFileSync(path.join(appsDir, name), 'utf8');
}

function functionBody(source, name) {
  const re = new RegExp(`function\\s+${name}\\s*\\(`);
  const match = re.exec(source);
  assert(match, `Function ${name} not found`);
  const start = match.index;
  const brace = source.indexOf('{', match.index);
  let depth = 0;
  let quote = null;
  let escape = false;

  for (let i = brace; i < source.length; i++) {
    const ch = source[i];
    if (quote) {
      if (escape) escape = false;
      else if (ch === '\\') escape = true;
      else if (ch === quote) quote = null;
      continue;
    }
    if (ch === "'" || ch === '"' || ch === '`') {
      quote = ch;
      continue;
    }
    if (ch === '{') depth++;
    if (ch === '}') {
      depth--;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  throw new Error(`Unbalanced function ${name}`);
}

function checkSyntax() {
  for (const file of fs.readdirSync(appsDir).filter(f => f.endsWith('.gs'))) {
    new Function(read(file));
  }
}

function checkNoDuplicateFunctions() {
  const seen = new Map();
  for (const file of fs.readdirSync(appsDir).filter(f => f.endsWith('.gs'))) {
    const source = read(file);
    const re = /function\s+([A-Za-z0-9_]+)\s*\(/g;
    let m;
    while ((m = re.exec(source))) {
      const name = m[1];
      if (seen.has(name)) {
        throw new Error(`Duplicate function ${name}: ${seen.get(name)} and ${file}`);
      }
      seen.set(name, file);
    }
  }
}

function checkInstallerSafety() {
  const installer = read('Installer.gs');
  const core = read('Core.gs');

  assert(installer.includes('3.0.0-alpha.9'), 'alpha.9 version marker missing');

  const installUi = functionBody(installer, 'installAttendanceWorkbook');
  assert(
    installUi.includes('recoverAttendanceWorkbookFromMenu'),
    'ready/manual install must redirect to safe recovery'
  );
  assert(
    !installUi.includes('resetAttendanceWorkbookDestructive_'),
    'normal install path must not call destructive reset'
  );

  const recover = functionBody(installer, 'recoverAttendanceWorkbookFromMenu');
  assert(
    recover.includes('repairAttendanceWorkbook_'),
    'recovery must use non-destructive repair'
  );

  const reset = functionBody(installer, 'resetAttendanceWorkbookFromMenu');
  const backupPos = reset.indexOf('createAttendanceBackup_');
  const destructivePos = reset.indexOf('resetAttendanceWorkbookDestructive_');
  assert(backupPos >= 0, 'full reset must create a backup');
  assert(destructivePos > backupPos, 'backup must happen before destructive reset');
  assert(reset.includes("'СБРОС'"), 'full reset must require typed confirmation');

  const cleanup = functionBody(installer, 'cleanupInstallerBlankSheets_');
  assert(
    cleanup.includes("getSheetByName('Установщик')"),
    'cleanup must target only the known bootstrap sheet'
  );
  assert(
    !cleanup.includes('spreadsheet.getSheets().slice().forEach'),
    'cleanup must not delete arbitrary blank sheets'
  );

  assert(
    core.includes("'recoverAttendanceWorkbookFromMenu'"),
    'normal menu must expose safe recovery'
  );
  assert(
    core.includes("'resetAttendanceWorkbookFromMenu'"),
    'normal menu must expose explicit full reset'
  );
}

function checkAttendanceConcurrency() {
  const attendance = read('Attendance.gs');
  const codes = read('Codes.gs');
  const submit = functionBody(attendance, 'submitAttendance');

  assert(
    submit.includes('LockService.getScriptLock'),
    'submitAttendance must use ScriptLock'
  );
  assert(
    submit.includes('tryLock(10000)'),
    'submitAttendance must wait for the lock with a bounded timeout'
  );

  const lockPos = submit.indexOf('tryLock(10000)');
  const checksAfterLock = submit.slice(lockPos).match(/alreadyMarked_\s*\(/g) || [];
  assert(
    checksAfterLock.length >= 1,
    'submitAttendance must re-check alreadyMarked_ after acquiring the lock'
  );

  const lockedPart = submit.slice(lockPos);
  const lessonPos = lockedPart.indexOf('getActiveLesson_');
  const validationPos = lockedPart.indexOf('validateAttendanceCode_');
  const duplicatePos = lockedPart.indexOf('alreadyMarked_');
  const appendPos = lockedPart.indexOf('marks.appendRow');
  assert(
    lessonPos >= 0 && validationPos > lessonPos && duplicatePos > validationPos && appendPos > duplicatePos,
    'lesson, check and duplicate state must be revalidated under the lock before append'
  );
  assert(
    lockedPart.includes("String(state.checkId || '') !== expectedCheckId"),
    'a submission must not cross into a newer check'
  );

  const validator = functionBody(attendance, 'validateAttendanceCode_');
  assert(
    validator.includes('rawCheck.lessonId') && validator.includes('rawCheck.checkId'),
    'previous-code grace must belong to the same lesson and check'
  );

  for (const name of ['startCheck_', 'stopCheck_']) {
    const body = functionBody(codes, name);
    assert(body.includes('LockService.getScriptLock'), `${name} must use ScriptLock`);
    assert(body.includes('waitLock(5000)'), `${name} must wait for ScriptLock`);
    assert(body.includes('finally') && body.includes('releaseLock'), `${name} must release ScriptLock`);
  }
}

function checkManualAttendanceHistory() {
  const events = read('Events.gs');
  const onEdit = functionBody(events, 'onEdit');
  const record = functionBody(events, 'recordManualAttendanceEdit_');
  assert(onEdit.includes('recordManualAttendanceEdit_(e)'), 'Journal edits must record attendance corrections');
  assert(record.includes("subheader !== 'Пос.'"), 'only attendance columns may create correction marks');
  assert(record.includes("cleared ? 'cleared' : 'manual'"), 'manual clears and values need explicit history statuses');
  assert(record.includes('LockService.getScriptLock'), 'manual history append must use ScriptLock');
}

function checkRecoverySafety() {
  const installer = read('Installer.gs');
  const recover = functionBody(installer, 'recoverAttendanceWorkbookFromMenu');
  const repair = functionBody(installer, 'repairAttendanceWorkbook_');
  const preflight = functionBody(installer, 'getRecoverySourceIssues_');
  const reset = functionBody(installer, 'resetAttendanceWorkbookDestructive_');

  assert(recover.indexOf('createAttendanceBackup_') < recover.indexOf('repairAttendanceWorkbook_'), 'recovery backup must precede repair');
  assert(repair.includes('LockService.getScriptLock') && repair.includes('releaseLock'), 'repair must use ScriptLock');
  assert(reset.includes('LockService.getScriptLock') && reset.includes('releaseLock'), 'destructive reset must use ScriptLock');
  assert(preflight.includes('hasJournalStudents') && preflight.includes('hasStudentSource'), 'recovery must protect student data when its source is missing');
  assert(preflight.includes('hasJournalLessons') && preflight.includes('hasLessonSource'), 'recovery must protect lesson data when its source is missing');
}

function checkLessonDateFormats() {
  const start = read('Lesson.gs');
  const functionPos = start.indexOf('function startLesson_');
  const appendPos = start.indexOf('appendRow', functionPos);
  const dateFormatPos = start.indexOf("setNumberFormat('dd.MM.yyyy')", functionPos);
  assert(dateFormatPos > appendPos, 'the appended lesson row must receive an explicit date format');
  assert(start.includes("setNumberFormat('dd.MM.yyyy HH:mm:ss')"), 'lesson timestamps must receive an explicit date-time format');
  assert(start.includes('typeNumber'), 'startLesson_ must return the assigned type number');
}

function checkScoreScaleDirection() {
  const installer = read('Installer.gs');
  assert(
    installer.includes("[0, 0, '#D9EAF7', '0 — синий', true]"),
    '0 must remain blue'
  );
  assert(
    installer.includes("[10, '', '#F4CCCC', '10+ — красный', true]"),
    '10+ must remain red'
  );
  assert(
    installer.includes('холодные цвета — низкие значения, тёплые — высокие'),
    'documentation note must explain intensity semantics'
  );
}

checkSyntax();
checkNoDuplicateFunctions();
checkInstallerSafety();
checkAttendanceConcurrency();
checkManualAttendanceHistory();
checkRecoverySafety();
checkLessonDateFormats();
checkScoreScaleDirection();

console.log('alpha.9 safety checks: OK');
