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

## Updating an existing deployment

When only the code changes:

1. Replace the Apps Script source files.
2. Open `Развернуть → Управление развертываниями`.
3. Edit the existing deployment.
4. Select **Новая версия**.
5. Deploy.

This normally preserves the existing `/exec` URL.

## Important

Do not use `ScriptApp.getService().getUrl()` as the authoritative production deployment URL.
The exact `/exec` URL should be copied from the active Apps Script deployment and stored explicitly in `web_app_url`.

## Files

Copy these files into the bound Apps Script project:

- `Code.gs`
- `Sidebar.html`
- `Student.html`
- `Display.html`
- `Teacher.html`
