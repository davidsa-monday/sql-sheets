// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { SqlSheetViewModel } from './viewmodels/SqlSheetViewModel';
import { SqlSheetEditorProvider } from './views/SqlSheetEditorProvider';
import { getSnowflakeService } from './services/snowflakeService';
import { getGoogleSheetsService } from './services/googleSheetsService';
import { SettingsWebviewProvider } from './views/SettingsWebviewProvider';
import { getSqlSheetsExportViewModel } from './viewmodels/SqlSheetsExportViewModel';
import { SqlFile } from './models/SqlFile';
import { SqlQuery } from './models/SqlQuery';

// Suppress specific Node.js warnings to keep the extension host logs readable while testing.
process.on('warning', (warning) => {
        if (
                warning.name === 'DeprecationWarning' &&
                warning.message.includes('The `punycode` module is deprecated')
        ) {
                return;
        }

        if (
                warning.name === 'ExperimentalWarning' &&
                warning.message.includes('SQLite is an experimental feature')
        ) {
                return;
        }

        console.warn(warning.stack ?? warning.message);
});

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	// Use the console to output diagnostic information (console.log) and errors (console.error)
	// This line of code will only be executed once when your extension is activated
	console.log('Congratulations, your extension "sql-sheets" is now active!');

	// Register command to show the SQL Sheet editor
	const showEditorCommand = vscode.commands.registerCommand('sql-sheets.showEditor', () => {
		// Show the SQL Sheet Editor view
		vscode.commands.executeCommand('sql-sheets.editor.focus');
	});

	// Register command to execute SQL query
	const executeQueryCommand = vscode.commands.registerCommand('sql-sheets.executeQuery', async () => {
		const snowflakeService = getSnowflakeService();

		// Check if connection is configured
		if (!snowflakeService.isConfigured()) {
			const loadCredentials = 'Load Credentials';
			const response = await vscode.window.showErrorMessage(
				'Snowflake connection is not configured yet.',
				loadCredentials
			);

			if (response === loadCredentials) {
				vscode.commands.executeCommand('sql-sheets.loadCredentials');
			}
			return;
		}

		// Get the active text editor
		const editor = vscode.window.activeTextEditor;
		if (!editor) {
			vscode.window.showErrorMessage('No active editor found');
			return;
		}

		// Get the selected text or the query at cursor position if no selection
		let query: string;
		const selection = editor.selection;

		if (selection.isEmpty) {
			// Get the query at the current cursor position
			query = getQueryAtCursorPosition(editor);
		} else {
			query = editor.document.getText(selection);
		}

		if (!query.trim()) {
			vscode.window.showErrorMessage('No query to execute');
			return;
		}

		// Execute the query
		try {
			// Show progress indicator
			await vscode.window.withProgress({
				location: vscode.ProgressLocation.Notification,
				title: 'Executing SQL query...',
				cancellable: false
			}, async (progress) => {
				try {
					const results = await snowflakeService.executeQuery(query);

					// Show results in a new editor
					const document = await vscode.workspace.openTextDocument({
						language: 'json',
						content: JSON.stringify(results, null, 2)
					});

					await vscode.window.showTextDocument(document, vscode.ViewColumn.Beside);
					vscode.window.showInformationMessage(`Query executed successfully. Returned ${results.length} rows.`);
				} catch (err) {
					vscode.window.showErrorMessage(`Query execution failed: ${err instanceof Error ? err.message : String(err)}`);
				}
			});
		} catch (err) {
			vscode.window.showErrorMessage(`Error executing query: ${err instanceof Error ? err.message : String(err)}`);
		}
	});


	// Register command to show settings page
	const showSettingsCommand = vscode.commands.registerCommand('sql-sheets.showSettings', () => {
		SettingsWebviewProvider.createOrShow(context.extensionUri);
	});

	// Register command for exporting query to Google Sheets
	const exportQueryToSheetsCommand = vscode.commands.registerCommand('sql-sheets.exportQueryToSheets', async (queryArg?: SqlQuery) => {
		const exportViewModel = getSqlSheetsExportViewModel();

		if (queryArg) {
			await exportViewModel.exportQueryToSheets(queryArg);
			return;
		}

		const editor = vscode.window.activeTextEditor;
		if (!editor) {
			vscode.window.showErrorMessage('No active editor found');
			return;
		}

		const sqlFile = new SqlFile(editor.document);
		const query = sqlFile.getQueryAt(editor.selection.active);

		if (!query) {
			vscode.window.showErrorMessage('No SQL query found at the current cursor position.');
			return;
		}

		await exportViewModel.exportQueryToSheets(query);
	});

	// Register command for exporting SQL file to Google Sheets
	const exportFileToSheetsCommand = vscode.commands.registerCommand('sql-sheets.exportFileToSheets', async () => {
		const editor = vscode.window.activeTextEditor;
		if (!editor) {
			vscode.window.showErrorMessage('No active editor found');
			return;
		}

		if (editor.document.languageId !== 'sql' && editor.document.languageId !== 'snowflake-sql') {
			vscode.window.showErrorMessage('Active file is not an SQL file.');
			return;
		}

		const sqlFile = new SqlFile(editor.document);
		if (sqlFile.queries.length === 0) {
			vscode.window.showInformationMessage('No SQL queries found in the active file.');
			return;
		}

		await vscode.window.withProgress({
			location: vscode.ProgressLocation.Notification,
			title: 'Exporting SQL queries to Google Sheets...',
			cancellable: false
		}, async (progress) => {
			for (let index = 0; index < sqlFile.queries.length; index++) {
				const query = sqlFile.queries[index];
				progress.report({ message: `Exporting query ${index + 1} of ${sqlFile.queries.length}...` });
				try {
					await vscode.commands.executeCommand('sql-sheets.exportQueryToSheets', query);
				} catch (err) {
					console.error(`Failed to export query ${index + 1}:`, err);
				}
			}
		});

		vscode.window.showInformationMessage('Finished exporting queries to Google Sheets.');
	});



	// Register all commands
	context.subscriptions.push(showEditorCommand);
	context.subscriptions.push(executeQueryCommand);
	context.subscriptions.push(showSettingsCommand);
	context.subscriptions.push(exportQueryToSheetsCommand);
	context.subscriptions.push(exportFileToSheetsCommand);


	// Register view provider
	const viewModel = new SqlSheetViewModel();
	const editorProvider = new SqlSheetEditorProvider(context.extensionUri, viewModel);
	context.subscriptions.push(
		vscode.window.registerWebviewViewProvider(SqlSheetEditorProvider.viewType, editorProvider));
}

/**
 * Extracts the SQL query at the current cursor position
 * It identifies a query as text between two semicolons or from the beginning/end of the file
 * @param editor The active text editor
 * @returns The SQL query at the cursor position
 */
function getQueryAtCursorPosition(editor: vscode.TextEditor): string {
	const document = editor.document;
	const cursorPosition = editor.selection.active;
	const text = document.getText();

	// Find the start of the current query (after the previous semicolon or beginning of file)
	let startPos = 0;
	let endPos = text.length;

	// Find the last semicolon before the cursor
	const textBeforeCursor = text.substring(0, document.offsetAt(cursorPosition));
	const lastSemicolonBeforeCursor = textBeforeCursor.lastIndexOf(';');

	if (lastSemicolonBeforeCursor !== -1) {
		startPos = lastSemicolonBeforeCursor + 1; // Start after the semicolon
	}

	// Find the next semicolon after the cursor
	const textAfterCursor = text.substring(document.offsetAt(cursorPosition));
	const nextSemicolonAfterCursor = textAfterCursor.indexOf(';');

	if (nextSemicolonAfterCursor !== -1) {
		endPos = document.offsetAt(cursorPosition) + nextSemicolonAfterCursor + 1; // Include the semicolon
	}

	return text.substring(startPos, endPos).trim();
}

// This method is called when your extension is deactivated
export function deactivate() {
	// Close any open Snowflake connections
	const snowflakeService = getSnowflakeService();
	return snowflakeService.closeConnection();

	// No cleanup needed for Google service
}
