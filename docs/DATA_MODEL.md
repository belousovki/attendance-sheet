# Data model

## Sheets

### `Список группы`

| Column | Meaning |
| --- | --- |
| № | Student number / stable local identifier |
| ФИО | Full name |
| Активен | Whether the student participates in future attendance checks |

### `Журнал`

Generated view.

Fixed columns:

- `№`
- `ФИО`

Each lesson creates two columns:

- `Пос.` — attendance
- `Оц.` — grade

Summary columns:

- `Баллы за посещение`
- `Баллы за работу`
- `Итого`

### `Занятия`

Lesson registry:

- ID занятия
- Дата
- Формат
- Тип занятия
- № по типу
- Тема
- Статус
- Начало
- Завершение
- Колонка Пос.
- Колонка Оц.
- Проверок
- Заметка
- Создано

### `Типы занятий`

Configurable lesson types:

- type name;
- HEX color;
- automatic numbering;
- active flag.

### `Настройки`

Key/value configuration.

Important keys include:

- `regular_code_seconds`
- `late_code_seconds`
- `attendance_present_points`
- `attendance_absent_points`
- `auto_absent_on_finish`
- `late_points`
- `code_digits`
- `student_search_min_chars`
- `remember_student_in_browser`
- `web_app_url`
- `student_url`
- `display_url`
- `teacher_url`
- `teacher_key`
- `instance_spreadsheet_id`

### `Отметки`

Technical attendance history. Usually hidden.

Attendance history is used as the primary recovery source when the journal is rebuilt.

## Design principle

`Журнал` is a projection, not the primary database.

This prevents structural manual edits from becoming the authoritative state and allows the journal to be rebuilt from the student list, lesson registry and technical attendance history.
