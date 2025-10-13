# SQL Sheets – AI Coding Guide

## Big Picture
- This is a VS Code extension that parses SQL files for special comment parameters and exports query results to Google Sheets; queries run against Snowflake.
- Entry: `src/extension.ts` registers commands and the webview panel; core flow is: read SQL + parameters → run in Snowflake → transform → write to Google Sheets.

## Key Modules
- `src/models/SqlSheetConfiguration.ts` defines supported parameters, descriptions, and boolean handling (true/1/yes recognized).
- `src/models/SqlFile.ts` splits a document into queries by semicolon, extracts `--key: value` header comments, and can edit/insert params in-place.
- `src/services/snowflakeService.ts` loads connection JSON via `sql-sheets.connection.credentialsFile` and executes queries using `snowflake-sdk`.
- `src/services/googleSheetsService.ts` loads a Google Service Account JSON via `sql-sheets.google.serviceAccountFile`, writes values with `USER_ENTERED`, clears target ranges, applies formatting, and can create/update Sheet “tables”.
- `src/services/googleSheetsExportService.ts` orchestrates: validates config, runs Snowflake query, adapts rows/headers, and calls Sheets upload.
- `src/viewmodels/SqlSheetViewModel.ts` reflects config keys/descriptions for the UI and updates parameters in the document.
- `src/views/SqlSheetEditorProvider.ts` hosts the parameter editor webview using `media/main.js` and `@vscode/webview-ui-toolkit`.
- `src/services/loggingService.ts` centralizes logging; `logger.info/warn/error` go to the Output channel (support-facing) while `logger.debug` targets the developer console.

## Parameters (comment headers)
- Add before each query block as SQL comments: `--spreadsheet_id: ...`, `--sheet_name: ...`, `--start_cell: A1`, `--transpose: true|false`, `--data_only: true|false`, `--skip: true|false`, optional `--name:` and `--table_name:`.
- Queries are delimited by semicolons; the editor and exporter operate on the query under the cursor or all queries in file.
- Booleans accept `true/false`, `1/0`, `yes/no` and are normalized when written back.

## Developer Workflow
- Build/watch: `npm install` then `npm run watch` (or use VS Code “Run Extension” which depends on the default watch task).
- One‑off build: `npm run compile`; lint: `npm run lint`; package VSIX: `npm run package`.
- VS Code tasks: see `.vscode/tasks.json` for “Clean and Compile”, “Package and Install” (macOS Code path).
- Debug configs: `.vscode/launch.json` (“Run Extension” uses the watch task).

## User Workflow (what commands expect)
- Configure credentials via command palette: `sql-sheets: Show Settings` to set:
  - Snowflake credentials JSON path at `sql-sheets.connection.credentialsFile` (must include `user`, `password`, `account`; optional `warehouse`, `database`, `schema`).
  - Google Service Account JSON path at `sql-sheets.google.serviceAccountFile`.
- Execute: `sql-sheets: Execute SQL Query` shows JSON results in a side editor.
- Export current query: `sql-sheets: Export SQL Query to Google Sheets`.
- Export entire file: `sql-sheets: Export SQL File to Google Sheets` loops all query blocks.

## Patterns and Conventions
- Parameter UI is schema‑driven from `SqlSheetConfiguration`; adding a new parameter there exposes it in the view automatically (types come from `isBooleanParameter`).
- Parameter insertion preserves leading comment blocks: new params are inserted before the first non-comment line in the query block.
- Export writes with `USER_ENTERED` and applies formatting (header bolding, column sizing, optional title/timestamp row) unless `data_only` is true.
- If `table_name` is set, the exporter attempts to create/update a Sheets “table” for that range.
- Logs: use the “SQL Sheets” Output channel for detailed Snowflake/Sheets request info and troubleshooting.
- Logging pattern: call `getLogger()` and prefer `logger.info/warn/error` for user/support visibility, and `logger.debug` for verbose payloads (e.g., Google API request bodies) so they only appear in the developer console.

## Integration Notes
- Snowflake: single connection is created on demand; duplicate columns are auto‑renamed by the driver (`rowMode: 'object_with_renamed_duplicated_columns'`).
- Google Sheets: uses `googleapis` + `google-auth-library` JWT; ranges are calculated from `start_cell` and data width; existing content is cleared in the target rectangle before upload.

## Gotchas
- `start_cell` must be A1‑style like `A1`; named ranges and “offset N” are mentioned in descriptions but not implemented by the uploader.
- `pre_file` and `start_named_range` are parsed into the model but currently unused by the exporter.
- When Snowflake isn’t configured, the Execute flow opens `sql-sheets: Show Settings` so users can set credential file paths.

## Example
```sql
-- spreadsheet_id: 1AbCDEF...sheetId...
-- sheet_name: Sheet1
-- start_cell: A1
-- name: Customer Orders
-- table_name: orders_table
-- transpose: false
-- data_only: false
SELECT * FROM analytics.orders WHERE order_date >= CURRENT_DATE - INTERVAL '7 DAY';
```
