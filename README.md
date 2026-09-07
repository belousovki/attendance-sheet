# Attendance Sheet

Google Sheets + Google Apps Script attendance journal for university courses.

The project provides a reusable attendance workflow with:

- a student list and generated journal;
- dynamic attendance codes;
- separate regular / late attendance modes;
- configurable attendance points;
- manual grades for class work;
- automatic absence points when a lesson is finished;
- configurable lesson types, numbering and colors;
- student web page;
- classroom display page;
- teacher control page;
- safe rebuilding of the journal from source data;
- safe deletion of lessons without manually deleting journal columns.

## Two distribution variants

### 1. Stable ready-made template

The `main` branch contains the tested **v2.7** workflow. A teacher makes a copy of a fully prepared Google Sheet and deploys its bound Apps Script project as a Web App.

This remains the recommended stable option.

### 2. Installer / loader prototype

This branch (`installer-v3-alpha`) experiments with **v3.0.0-alpha.1**.

Instead of depending on a prebuilt spreadsheet layout, the Apps Script project contains `Installer.gs`. The installer can create the required sheets, headers, validation rules, lesson types, settings and technical structures in a blank Google Sheet.

See [docs/INSTALLER_ALPHA.md](docs/INSTALLER_ALPHA.md).

The current alpha intentionally requires running `installAttendanceWorkbook()` once from the Apps Script editor. After independent testing, the next step is a bootstrap `onOpen()` that detects a missing schema and offers an `Установить журнал…` menu automatically.

## Data model

The journal is treated as a generated view. Source data lives in:

- `Список группы` — students;
- `Занятия` — lesson registry;
- `Типы занятий` — lesson types and colors;
- `Настройки` — configuration;
- `Отметки` — technical attendance history.

The `Журнал` sheet should not be structurally edited by hand. Attendance (`Пос.`) and grades (`Оц.`) may be corrected manually when needed.

## Apps Script source

The tested v2.7 project originally used one large `Code.gs`. In this repository that file is split at top-level boundaries into smaller Apps Script modules for easier maintenance. Google Apps Script loads all `.gs` files into the same project namespace, so no imports are required.

```text
apps-script/
├── Core.gs
├── Settings.gs
├── Types.gs
├── Students.gs
├── Journal.gs
├── Events.gs
├── Lesson.gs
├── Codes.gs
├── Attendance.gs
├── Interfaces.gs
├── Utils.gs
├── Installer.gs          # installer-v3-alpha only
├── Sidebar.html
├── Student.html
├── Display.html
└── Teacher.html
```

## Installation

Stable template: [docs/INSTALL.md](docs/INSTALL.md)

Installer alpha: [docs/INSTALLER_ALPHA.md](docs/INSTALLER_ALPHA.md)

## Usage

See [docs/USAGE.md](docs/USAGE.md).

## Security notes

Do not publish real:

- student lists;
- grades;
- `teacher_key`;
- deployment URLs from working courses;
- short URLs tied to a private deployment.

Each course copy should have its own web-app deployment and its own generated teacher key.

## License

MIT License. See [LICENSE](LICENSE).
