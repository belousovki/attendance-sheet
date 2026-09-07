# Changelog

## v3.0.0-alpha.8 — Installer Edition

- added editable `Шкала оценок` with a fixed universal score-to-color mapping;
- default visual scale: `0` blue, `1–2` green, `3–4` light green, `5` yellow, `6–7` yellow-orange, `8–9` orange, `10+` red;
- applies the same color semantics to all lesson `Пос.` / `Оц.` cells;
- summary columns are excluded from the score scale;
- score-scale edits automatically rebuild conditional formatting;
- journal rebuilds and new lessons resynchronize score formatting;
- fixed alpha.7 ordering bug: formatting is now applied only after the new lesson is written to the lesson registry;
- canonical Installer Template and a fresh copied instance were smoke-tested successfully.

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