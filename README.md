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

## Current version

**v2.7**

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
├── Sidebar.html
├── Student.html
├── Display.html
└── Teacher.html
```

The `.gs` modules are an exact split of the tested v2.7 `Code.gs`; their concatenation in the order above reproduces the original source.

## Installation

See [docs/INSTALL.md](docs/INSTALL.md).

## Usage

See [docs/USAGE.md](docs/USAGE.md).

## Data model

See [docs/DATA_MODEL.md](docs/DATA_MODEL.md).

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
