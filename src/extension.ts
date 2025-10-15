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
import { ExportResult, ExportOptions, getSqlSheetsExportViewModel, SqlSheetsExportViewModel } from './viewmodels/SqlSheetsExportViewModel';
import { SqlFile } from './models/SqlFile';
import { SqlQuery } from './models/SqlQuery';
import { getLogger } from './services/loggingService';

const logger = getLogger();

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	logger.info('SQL Sheets extension activated');

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
			const loadCredentials = 'Open Settings';
			const response = await vscode.window.showErrorMessage(
				'Snowflake connection is not configured yet.',
				loadCredentials
			);

			if (response === loadCredentials) {
				vscode.commands.executeCommand('sql-sheets.showSettings');
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

	const runSingleQueryExport = async (
		exportViewModel: SqlSheetsExportViewModel,
		query: SqlQuery,
		options?: ExportOptions
	): Promise<ExportResult> => {
		const result = await exportViewModel.exportQueryToSheets(query, options);
		switch (result) {
			case ExportResult.SkippedMissingConfig:
				vscode.window.showWarningMessage('Please add spreadsheet_id, a sheet ID or sheet name, and a start cell or named range before exporting.');
				break;
			case ExportResult.ExecutedCreate:
				vscode.window.showInformationMessage('Executed CREATE statement in Snowflake.');
				break;
			default:
				break;
		}
		return result;
	};

	// Register command for exporting query to Google Sheets
	const exportQueryToSheetsCommand = vscode.commands.registerCommand('sql-sheets.exportQueryToSheets', async (queryArg?: SqlQuery) => {
		const exportViewModel = getSqlSheetsExportViewModel();

		if (queryArg) {
			return runSingleQueryExport(exportViewModel, queryArg);
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

		return runSingleQueryExport(exportViewModel, query);
	});

	const exportQueryToSheetsWithDepsCommand = vscode.commands.registerCommand('sql-sheets.exportQueryToSheetsWithDeps', async (queryArg?: SqlQuery) => {
		const exportViewModel = getSqlSheetsExportViewModel();
		const executedPreFiles = new Set<string>();

		if (queryArg) {
			return runSingleQueryExport(exportViewModel, queryArg, { executeDependencies: true, executedPreFiles });
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

		return runSingleQueryExport(exportViewModel, query, { executeDependencies: true, executedPreFiles });
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

		const exportViewModel = getSqlSheetsExportViewModel();
		let skippedMissingConfigCount = 0;
		let cancelledCount = 0;
		let executedCreateCount = 0;

		await vscode.window.withProgress({
			location: vscode.ProgressLocation.Notification,
			title: 'Exporting SQL queries to Google Sheets...',
			cancellable: false
		}, async (progress) => {
			for (let index = 0; index < sqlFile.queries.length; index++) {
				const query = sqlFile.queries[index];
				progress.report({ message: `Exporting query ${index + 1} of ${sqlFile.queries.length}...` });
				try {
					const result = await exportViewModel.exportQueryToSheets(query);
					switch (result) {
						case ExportResult.SkippedMissingConfig:
							skippedMissingConfigCount++;
							break;
						case ExportResult.UserCancelled:
							cancelledCount++;
							break;
						case ExportResult.ExecutedCreate:
							executedCreateCount++;
							break;
						default:
							break;
					}
				} catch (err) {
					logger.error(`Failed to export query ${index + 1}`, { data: err });
				}
			}
		});

		const summaryMessages: string[] = [];
		if (skippedMissingConfigCount > 0) {
			summaryMessages.push(`${skippedMissingConfigCount} skipped (missing configuration)`);
		}
		if (cancelledCount > 0) {
			summaryMessages.push(`${cancelledCount} cancelled by user`);
		}
		if (executedCreateCount > 0) {
			summaryMessages.push(`${executedCreateCount} CREATE statements executed`);
		}

		const summarySuffix = summaryMessages.length > 0 ? ` (${summaryMessages.join(', ')})` : '';
		vscode.window.showInformationMessage(`Finished exporting queries to Google Sheets${summarySuffix}.`);
	});

	const exportFileToSheetsWithDepsCommand = vscode.commands.registerCommand('sql-sheets.exportFileToSheetsWithDeps', async () => {
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

		const exportViewModel = getSqlSheetsExportViewModel();
		const executedPreFiles = new Set<string>();
		const globalPreFiles = sqlFile.getGlobalPreFiles();
		if (globalPreFiles.length > 0) {
			const ranGlobalPreFiles = await exportViewModel.runPreFiles(editor.document.uri, globalPreFiles, executedPreFiles);
			if (!ranGlobalPreFiles) {
				return;
			}
		}
		let skippedMissingConfigCount = 0;
		let cancelledCount = 0;
		let executedCreateCount = 0;

		await vscode.window.withProgress({
			location: vscode.ProgressLocation.Notification,
			title: 'Exporting SQL queries with dependencies...',
			cancellable: false
		}, async (progress) => {
			for (let index = 0; index < sqlFile.queries.length; index++) {
				const query = sqlFile.queries[index];
				progress.report({ message: `Exporting query ${index + 1} of ${sqlFile.queries.length}...` });
				try {
					const result = await exportViewModel.exportQueryToSheets(query, {
						executeDependencies: true,
						executedPreFiles
					});
					switch (result) {
						case ExportResult.SkippedMissingConfig:
							skippedMissingConfigCount++;
							break;
						case ExportResult.UserCancelled:
							cancelledCount++;
							break;
						case ExportResult.ExecutedCreate:
							executedCreateCount++;
							break;
						default:
							break;
					}
				} catch (err) {
					logger.error(`Failed to export query ${index + 1}`, { data: err });
				}
			}
		});

		const summaryMessages: string[] = [];
		if (skippedMissingConfigCount > 0) {
			summaryMessages.push(`${skippedMissingConfigCount} skipped (missing configuration)`);
		}
		if (cancelledCount > 0) {
			summaryMessages.push(`${cancelledCount} cancelled by user`);
		}
		if (executedCreateCount > 0) {
			summaryMessages.push(`${executedCreateCount} CREATE statements executed`);
		}
		if (executedPreFiles.size > 0) {
			summaryMessages.push(`${executedPreFiles.size} pre-files executed`);
		}

		const summarySuffix = summaryMessages.length > 0 ? ` (${summaryMessages.join(', ')})` : '';
		vscode.window.showInformationMessage(`Finished exporting queries to Google Sheets${summarySuffix}.`);
	});



	// Register all commands
	context.subscriptions.push(showEditorCommand);
	context.subscriptions.push(executeQueryCommand);
	context.subscriptions.push(showSettingsCommand);
	context.subscriptions.push(exportQueryToSheetsCommand);
	context.subscriptions.push(exportQueryToSheetsWithDepsCommand);
	context.subscriptions.push(exportFileToSheetsCommand);
	context.subscriptions.push(exportFileToSheetsWithDepsCommand);


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
