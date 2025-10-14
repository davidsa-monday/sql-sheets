# SQL Sheets

SQL Sheets is a VS Code extension that helps you manage SQL query parameters for exporting data to Google Sheets. The extension provides a convenient editor pane to view and edit parameters within SQL files, and execute SQL queries directly against Snowflake.

## Features

- Parameter editor for SQL files with Google Sheets integration parameters
- Auto-detection of parameters in SQL files
- Edit parameters directly in the editor pane
- Parameter tooltips showing descriptions
- Specialized field types (dropdowns for boolean values)
- Execute SQL queries directly against Snowflake
- Load Snowflake credentials from a JSON file

## Requirements

This extension works with SQL files that include special comment headers for Google Sheets integration.

## Usage

1. Open a SQL file in VS Code
2. Add parameters as comments in the format `--parameter: value`
3. Configure Snowflake connection using one of these methods:
   - Command: "sql-sheets: Configure Snowflake Connection"
   - Command: "sql-sheets: Load Snowflake Credentials from File"
4. Execute SQL queries using the command "sql-sheets: Execute SQL Query"

## Configuring Snowflake Connection

### Method 1: Using the Configuration UI

1. Open the command palette (`Ctrl+Shift+P` or `Cmd+Shift+P`)
2. Run the command "sql-sheets: Configure Snowflake Connection"
3. Enter your Snowflake credentials when prompted

### Method 2: Using a Credentials File

1. Create a JSON file with your Snowflake credentials using the following format:
   ```json
   {
       "user": "your_username",
       "password": "your_password",
       "account": "your_account",
       "warehouse": "your_warehouse",
       "database": "your_database",
       "schema": "your_schema"
   }
   ```
2. Open the command palette (`Ctrl+Shift+P` or `Cmd+Shift+P`)
3. Run the command "sql-sheets: Load Snowflake Credentials from File"
4. Select your credentials JSON file when prompted

A template credentials file is provided in the `samples` directory of this extension.
3. Click on the SQL Parameters icon in the Activity Bar (looks like a database icon)
4. The SQL Sheet Editor pane will appear, showing the parameters for your SQL file
5. If the view is not visible, you can also run the "Show SQL Sheet Editor" command from the Command Palette (Ctrl+Shift+P or Cmd+Shift+P)

### Enabling the View

The SQL Parameters view should automatically appear in your Activity Bar after installing the extension. If you don't see it:

1. Open the Command Palette with `Ctrl+Shift+P` (Windows/Linux) or `Cmd+Shift+P` (macOS)
2. Type "View: Show SQL Parameters" and select it when it appears
3. The SQL Parameters icon should now appear in your Activity Bar

## Parameters

The following parameters are supported:

- `spreadsheet_id`: The Google Sheet ID where results will be sent
- `sheet_name`: The sheet identifier. Supply an ID, a name, or combine both as `315704920 | Summary` to keep them in sync automatically.
- `start_cell`: The cell where output should begin (e.g., "A1"). You can also combine a named range by writing `MyNamedRange | A1`, or provide just the named range name to anchor the export dynamically. When a named range is supplied, the extension keeps the range anchored to the first cell and will create or update the named range automatically after each export.
- `start_named_range`: (Legacy) Alternative to `start_cell` for referencing a named range. Prefer the combined `start_cell` syntax above.
- `transpose`: Set to "true" to transpose the output data
- `name`: A title for the table
- `data_only`: Set to "true" to output only data without headers
- `skip`: Set to "true" to skip processing this query
- `pre_file`: Path to a SQL file that should be executed before this query

### Commands

- `sql-sheets: Export SQL Query to Google Sheets`
- `sql-sheets: Export SQL Query to Google Sheets (with pre_file)`
- `sql-sheets: Export SQL File to Google Sheets`
- `sql-sheets: Export SQL File to Google Sheets (with pre_file)`

When you use the “with pre_file” commands, each query’s `--pre_file` is resolved relative to the SQL file and executed once per export run before the main query. The results are cached so shared pre-files only run a single time per export.

Include if your extension adds any VS Code settings through the `contributes.configuration` extension point.

For example:

This extension contributes the following settings:

* `myExtension.enable`: Enable/disable this extension.
* `myExtension.thing`: Set to `blah` to do something.

## Known Issues

Calling out known issues can help limit users opening duplicate issues against your extension.

## Release Notes

Users appreciate release notes as you update your extension.

### 1.0.0

Initial release of ...

### 1.0.1

Fixed issue #.

### 1.1.0

Added features X, Y, and Z.

---

## Following extension guidelines

Ensure that you've read through the extensions guidelines and follow the best practices for creating your extension.

* [Extension Guidelines](https://code.visualstudio.com/api/references/extension-guidelines)

## Working with Markdown

You can author your README using Visual Studio Code. Here are some useful editor keyboard shortcuts:

* Split the editor (`Cmd+\` on macOS or `Ctrl+\` on Windows and Linux).
* Toggle preview (`Shift+Cmd+V` on macOS or `Shift+Ctrl+V` on Windows and Linux).
* Press `Ctrl+Space` (Windows, Linux, macOS) to see a list of Markdown snippets.

## For more information

* [Visual Studio Code's Markdown Support](http://code.visualstudio.com/docs/languages/markdown)
* [Markdown Syntax Reference](https://help.github.com/articles/markdown-basics/)

**Enjoy!**
