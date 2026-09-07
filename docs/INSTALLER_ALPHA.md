# Installer Edition — v3 alpha.9

This branch contains the tested installer distribution model for Attendance Sheet while keeping the stable v2.7 template workflow on `main`.

## End-user model

The teacher does **not** need GitHub, `clasp`, a terminal, or manual Apps Script file creation.

The canonical Installer Template is a Google Sheet with the complete bound Apps Script project already attached. The visible workbook itself is intentionally minimal.

Canonical copy link: [https://docs.google.com/spreadsheets/d/10nzyqzAO_Hnc20MFKZKDVmikuQ4OOFCQNdk3JwpGnXQ/copy](https://docs.google.com/spreadsheets/d/10nzyqzAO_Hnc20MFKZKDVmikuQ4OOFCQNdk3JwpGnXQ/copy)

The source template must be shared as **Anyone with the link → Viewer** before external distribution.

User flow:

1. Make a copy of the Installer Template.
2. Open the copy and reload it if necessary.
3. Choose `Посещаемость → Установить / восстановить журнал…`.
4. Enter the discipline name.
5. Grant Google permissions when requested.
6. Fill `Список группы`.
7. Start lessons from `Посещаемость → Начать занятие…` or from the teacher control page.
8. Deploy the bound Apps Script project as a Web App.
9. Copy the exact production URL ending in `/exec`.
10. Save it through `Посещаемость → Настроить URL веб-приложения`.

The installer can be rerun after an interrupted installation and reconstructs a partially created workbook **without clearing existing course data**.

For an already installed workbook the normal menu exposes `Восстановить структуру журнала…`. Recovery preserves students, lesson registry, grades, attendance history and existing editable configuration while recreating missing schema elements and rebuilding the generated journal.

A destructive reset is separated into `Полный сброс журнала…`. It requires a second typed confirmation (`СБРОС`) and automatically creates a Google Drive backup before clearing data.

## Workbook created by the installer

- `Список группы`;
- `Журнал`;
- `Занятия`;
- `Типы занятий`;
- `Настройки`;
- hidden technical sheet `Отметки`;
- `Шкала оценок` — editable universal score-to-color mapping;
- headers, widths, frozen rows/columns, checkboxes and dropdown validation;
- default lesson types and colors;
- default settings;
- `teacher_key` and instance identifiers;
- version markers.

## Bootstrap behavior

`onOpen()` checks whether the workbook is installed.

If the schema is missing or incomplete, the menu contains only installer-oriented commands:

- `Установить / восстановить журнал…`;
- `Диагностика установки`.

After successful installation the normal attendance menu is shown.

## Tested workflow

A fresh copy made from the canonical alpha.9 template has been installed and deployed as a Web App end-to-end. The live smoke test verified:

- installation from the minimal `Установщик` sheet;
- `app_version = 3.0.0-alpha.9`, `code_version = 3.0.0-alpha.9`, `schema_version = 3`, `install_status = ready`;
- student, display and teacher Web App routes from the configured `/exec` deployment;
- lesson creation and closing;
- regular attendance codes;
- late-attendance codes with separate late points;
- student submissions written to the technical `Отметки` history;
- automatic absence records on lesson finish;
- manual `Оц.` values remaining independent from attendance;
- immediate fixed conditional formatting for both `Пос.` and `Оц.`;
- cold-to-hot score scale (`0` blue → `10+` red);
- late-attendance notes in the journal;
- summary columns remaining outside the per-lesson color scale.

The previously verified scenarios remain part of the functional baseline:

- recovery after a partially failed installation;
- automatic student synchronization;
- creating a lesson from the spreadsheet menu and teacher control page;
- safe lesson deletion;
- diagnostics;
- editable universal score scale.

alpha.9 additionally adds automated **source-level safety tests** for:

- non-destructive recovery path;
- explicit backup-before-reset ordering;
- bootstrap-only cleanup;
- attendance duplicate protection under `ScriptLock`;
- JavaScript syntax and duplicate-function checks;
- fixed cold-to-hot score-scale direction.

These tests do not replace live Google Apps Script integration tests; they are a regression guard for the safety properties introduced in alpha.9.

## Universal score scale

Installer alpha.9 adds the editable `Шкала оценок` sheet. It defines one fixed visual language for all `Пос.` and `Оц.` cells, independent of the lesson column:

| From | To | Default meaning |
| ---: | ---: | --- |
| 0 | 0 | blue |
| 1 | 2 | green |
| 3 | 4 | light green |
| 5 | 5 | yellow |
| 6 | 7 | yellow-orange |
| 8 | 9 | orange |
| 10 | blank | red (`10+`) |

Each row also contains a HEX color, human-readable description and `Активен` checkbox.

The same score therefore has the same color everywhere in the journal. Summary columns (`Баллы за посещение`, `Баллы за работу`, `Итого`) are intentionally excluded because they are cumulative values and are not directly comparable with a single lesson score.

The scale is intentionally **cold → hot by score intensity**: low values are blue and high values are red. Red is not an error/warning color here; it represents a high score.

Changing the scale rebuilds the journal conditional-formatting rules. Starting a lesson, rebuilding the journal and reopening the workbook also resynchronize the score formatting.

alpha.8 fixed the alpha.7 ordering bug where formatting rules were requested before the newly created lesson had been written to the `Занятия` registry.

## alpha.9 hardening

alpha.9 implements the safety fixes identified during review:

- normal installation/recovery is non-destructive;
- an already installed workbook can no longer be silently “reinstalled” through the ordinary installer path;
- `Восстановить структуру журнала…` repairs missing schema elements and rebuilds the generated journal while preserving course data;
- `Полный сброс журнала…` is a separate operation with an explicit warning, typed `СБРОС` confirmation and automatic Drive backup;
- `cleanupInstallerBlankSheets_()` deletes only the known blank bootstrap sheet `Установщик` with the expected 100×10 shape; arbitrary empty user sheets are never removed;
- `submitAttendance()` uses `LockService.getScriptLock()` and repeats the duplicate check under the lock before appending to `Отметки`;
- repository tests include syntax, duplicate-function, installer-safety, concurrency-lock and score-scale-direction checks.

## Status model

Lesson status values are standardized as:

- `active`;
- `closed`.

Legacy `finished` values are migrated to `closed`.

## Date/time formats

The installer applies:

- `Дата` → `dd.MM.yyyy`;
- `Начало` → `dd.MM.yyyy HH:mm:ss`;
- `Завершение` → `dd.MM.yyyy HH:mm:ss`;
- `Создано` → `dd.MM.yyyy HH:mm:ss`.

## Version markers

Current alpha writes and synchronizes:

- `app_version = 3.0.0-alpha.9`;
- `code_version = 3.0.0-alpha.9`;
- `schema_version = 3`;
- `install_status = ready`.

`code_version` identifies the installed application code. `schema_version` is intentionally separate so future schema migrations can be introduced without conflating them with ordinary code releases.

## Diagnostics

Use:

`Посещаемость → Диагностика установки`

The diagnostic reports:

- required sheet structure;
- current code version;
- stored application version;
- expected and stored schema version;
- installation state;
- whether Web App URL is configured;
- student count and active-student count;
- lesson count and active-lesson count;
- final health status.

## Maintainer workflow

The modular source lives in `apps-script/`.

`clasp` may be used by maintainers to synchronize the repository with the canonical Apps Script project. This is a development workflow only; it must not be exposed as an installation requirement to teachers.

## Alpha status

v3.0.0-alpha.9 is the current Installer Edition candidate. A fresh copy of the canonical template has passed the full installation/Web App smoke test. The branch remains alpha while the same build is used in a small number of real teaching sessions and the new safe-recovery/reset paths receive additional practical testing.
