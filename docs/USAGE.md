# Usage

## Starting a lesson

Use the teacher interface or the spreadsheet attendance menu to start a lesson.

A lesson stores:

- date;
- mode (`Очно` / `Дистанционно`);
- lesson type;
- automatic number within type;
- topic;
- notes;
- status;
- start / finish timestamps.

## Attendance codes

Default behavior:

- 4-digit code;
- code changes every 60 seconds;
- regular and late attendance can use different point values;
- a short grace period is accepted after a code changes.

The current code is intentionally not shown on the student page.

## Finishing a lesson

When a lesson is finished:

- students who have already checked in retain their attendance points;
- if `auto_absent_on_finish = TRUE`, active students without a mark receive `attendance_absent_points`;
- a real `0` means confirmed absence;
- a blank cell means no applicable / no recorded data.

## Manual grades

`Оц.` is a manual grade for work during the lesson.

Checking in for attendance does **not** automatically create a grade.

## Editing the journal

Safe manual edits:

- attendance values in `Пос.`;
- grade values in `Оц.`.

Do not manually:

- delete student rows;
- insert or delete lesson columns;
- change service headers;
- move summary columns.

Instead use:

- `Посещаемость → Синхронизировать состав группы`;
- `Посещаемость → Перестроить журнал из исходных данных`;
- `Посещаемость → Удалить занятие по выбранному столбцу…`.

## Deleting a lesson from the journal

1. In `Журнал`, select any cell in the `Пос.` or `Оц.` column of the lesson.
2. Run:
   `Посещаемость → Удалить занятие по выбранному столбцу…`
3. Confirm.

The script removes:

- the lesson from `Занятия`;
- technical attendance marks;
- the corresponding `Пос.` / `Оц.` pair;
- active lesson state if necessary.

The journal is then rebuilt automatically.

## Preparing a clean template

Use:

`Посещаемость → Подготовить как пустой шаблон…`

This safely clears course-specific data without manually deleting journal structure.
