# Changelog

## v3.0.0-alpha.6 — Installer Edition

- completed live installer workflow from a blank Google Sheet;
- installer can recover from a partially created workbook;
- bootstrap `onOpen()` exposes installer commands before the schema exists;
- added menu-based lesson creation without requiring the Web App;
- classroom display now shows both regular and late-attendance codes;
- standardized lesson statuses as `active / closed`;
- added code/application/schema version markers;
- added date/time formatting for the lesson registry;
- expanded installation diagnostics;
- verified student page, display page, teacher control page, regular/late attendance, automatic absences and safe lesson deletion end-to-end.

## v2.7

- delete a lesson directly from the selected `Пос.` / `Оц.` journal column;
- safe lesson deletion through the lesson registry remains available.

## v2.6

- fixed Google Sheets error when rebuilding a journal with frozen columns;
- journal width is resized safely instead of deleting all unfrozen columns.

## v2.5

- safe lesson deletion command;
- cleanup of phantom `0 / 0 / 0`;
- distinction between blank data and an explicit zero.

## v2.4

- fixed attendance write / polling race condition;
- technical attendance history is used to restore journal attendance;
- automatic absences are also written to technical history.

## v2.3

- compact student block in the journal;
- journal rebuild command;
- safe command for preparing an empty template.

## v2.2

- explicit web-app URL management;
- cleanup of stale deployment-derived URLs.
