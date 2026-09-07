# Installer variant — v3 alpha

This branch experiments with a second distribution model while keeping the stable template workflow intact.

## Two variants

### Variant A — ready Google Sheets template

This is the current stable workflow on `main`.

The user makes a copy of a fully prepared spreadsheet. The copy already contains all sheets, formatting and the bound Apps Script project.

Advantages:

- simplest for most teachers;
- almost no setup;
- easy to explain and support.

### Variant B — installer / loader

The spreadsheet starts essentially blank, but the bound Apps Script project contains the application code plus `Installer.gs`.

The installer creates the workbook structure automatically:

- `Список группы`;
- `Журнал`;
- `Занятия`;
- `Типы занятий`;
- `Настройки`;
- hidden technical sheet `Отметки`;
- headers, widths, checkboxes and dropdown validation;
- default lesson types and colors;
- default settings;
- `teacher_key` and instance identifiers;
- `app_version` and `schema_version`.

## Alpha test procedure

At this first stage the installer is intentionally isolated from the stable `onOpen()` logic.

1. Create a blank Google Sheet.
2. Attach the Apps Script project from this branch (for development, `clasp` is the intended route).
3. In Apps Script run `installAttendanceWorkbook()` once.
4. Grant permissions when Google asks.
5. Return to the spreadsheet and reload it.
6. The normal `Посещаемость` menu should now work.
7. Fill `Список группы` and synchronize students.
8. Deploy the Apps Script project as a Web App.
9. Paste the exact `/exec` URL through `Посещаемость → Настроить URL веб-приложения`.

## Why the first run is manual in alpha

The stable v2.7 `onOpen()` assumes that the workbook structure already exists. Calling it in a completely blank spreadsheet would fail before the installer can run.

The next installer iteration should change the bootstrap behavior so that `onOpen()` detects a missing schema and shows only an `Установить журнал…` command. That change is intentionally not being made on `main` until the installer is tested independently.

## Diagnostics

Run `diagnoseAttendanceInstallation()` to check the required sheets and version markers.

## Version markers

The installer writes:

- `app_version = 3.0.0-alpha.1`
- `schema_version = 3`

These markers are the basis for future migration functions such as `migrateSchema3To4_()`.
