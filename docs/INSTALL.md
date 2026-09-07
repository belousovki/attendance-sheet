# Installation

## Recommended workflow

The recommended workflow is based on a prepared Google Sheets template with the Apps Script project bound to it.

1. Make a copy of the clean spreadsheet template.
2. Rename the copy for the course.
3. Open the spreadsheet once so the bound Apps Script can initialize the copy.
4. Set the course name in `Настройки`.
5. Fill `Список группы`.
6. Run `Посещаемость → Синхронизировать состав группы`.
7. Open `Расширения → Apps Script`.
8. Deploy as a **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
9. Copy the exact deployment URL ending with `/exec`.
10. In the spreadsheet run:
    `Посещаемость → Настроить URL веб-приложения`
11. Paste the `/exec` URL.

The application will generate the student, display and teacher URLs.

## Apps Script source files

The GitHub version is modularized. Add all `.gs` files from `apps-script/` to the same Apps Script project:

- `Core.gs`
- `Settings.gs`
- `Types.gs`
- `Students.gs`
- `Journal.gs`
- `Events.gs`
- `Lesson.gs`
- `Codes.gs`
- `Attendance.gs`
- `Interfaces.gs`
- `Utils.gs`

Also add the HTML files:

- `Sidebar.html`
- `Student.html`
- `Display.html`
- `Teacher.html`

Google Apps Script shares one global namespace across `.gs` files, so these modules do not need imports. They are an exact split of the tested v2.7 monolithic `Code.gs`.

If you already have the tested single-file `Code.gs` in a working template, you do **not** need to replace it with the modular layout merely to use the application. The split is primarily for repository maintenance and future development.

## Updating an existing deployment

When only the code changes:

1. Update the Apps Script source files.
2. Open `Развернуть → Управление развертываниями`.
3. Edit the existing deployment.
4. Select **Новая версия**.
5. Deploy.

This normally preserves the existing `/exec` URL.

## Important

Do not use `ScriptApp.getService().getUrl()` as the authoritative production deployment URL.
The exact `/exec` URL should be copied from the active Apps Script deployment and stored explicitly in `web_app_url`.

Never commit working-course secrets or data to the repository: real student lists, grades, `teacher_key`, deployment URLs or course-specific short links.
