# Installer Edition — v3 alpha.6

This branch contains the tested installer distribution model for Attendance Sheet while keeping the stable v2.7 template workflow on `main`.

## End-user model

The teacher does **not** need GitHub, `clasp`, a terminal, or manual Apps Script file creation.

The canonical Installer Template is a Google Sheet with the complete bound Apps Script project already attached. The visible workbook itself is intentionally minimal.

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

The installer can be rerun after an interrupted installation and reconstructs a partially created workbook.

## Workbook created by the installer

- `Список группы`;
- `Журнал`;
- `Занятия`;
- `Типы занятий`;
- `Настройки`;
- hidden technical sheet `Отметки`;
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

## Tested alpha.6 workflow

The alpha.6 integration test covered:

- installation from a blank spreadsheet;
- recovery after a partially failed installation;
- automatic student synchronization;
- creating a lesson from the spreadsheet menu;
- creating a lesson from the teacher control page;
- creation of `Пос.` / `Оц.` columns in the generated journal;
- regular attendance codes;
- late-attendance codes;
- regular and late codes on the classroom display;
- student attendance submission;
- configurable late points;
- automatic absence points on lesson finish;
- technical attendance history in `Отметки`;
- safe lesson deletion;
- version/schema diagnostics.

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

- `app_version = 3.0.0-alpha.6`;
- `code_version = 3.0.0-alpha.6`;
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

v3.0.0-alpha.6 is the first installer build that completed the full live integration workflow. The branch remains alpha until the canonical Installer Template has been copied and tested as a fresh end-user instance.
