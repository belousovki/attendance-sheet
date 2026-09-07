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

This remains the recommended stable option until v3 leaves alpha.

### 2. Installer Edition — v3 alpha

The `installer-v3-alpha` branch now contains the tested **v3.0.0-alpha.8** installer workflow.

The canonical Installer Template is intentionally almost blank. Its bound Apps Script project contains the whole application. After making a copy, the teacher opens the spreadsheet and chooses:

`Посещаемость → Установить / восстановить журнал…`

The installer creates the workbook structure, settings, validation rules and technical sheets automatically. No `clasp`, terminal, or manual addition of Apps Script files is required for end users.

After installation the teacher fills the student list, deploys the bound Apps Script project as a Web App, and saves the exact `/exec` URL through the spreadsheet menu.

See [docs/INSTALLER_ALPHA.md](docs/INSTALLER_ALPHA.md).

## Data model

The journal is treated as a generated view. Source data lives in:

- `Список группы` — students;
- `Занятия` — lesson registry;
- `Типы занятий` — lesson types and colors;
- `Настройки` — configuration;
- `Шкала оценок` — universal fixed color scale for lesson scores;
- `Отметки` — technical attendance history.

The `Журнал` sheet should not be structurally edited by hand. Attendance (`Пос.`) and grades (`Оц.`) may be corrected manually when needed.

## Apps Script source

The tested project originally used one large `Code.gs`. In this repository the code is split at top-level boundaries into smaller Apps Script modules for easier maintenance. Google Apps Script loads all `.gs` files into the same project namespace, so no imports are required.

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
├── Installer.gs          # installer + score-scale configuration
├── Sidebar.html
├── Student.html
├── Display.html
└── Teacher.html
```

`clasp` is a maintainer/developer convenience for syncing these source files with the canonical Apps Script project. It is not part of the teacher workflow.

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