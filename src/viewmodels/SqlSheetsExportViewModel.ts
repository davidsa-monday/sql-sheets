import * as vscode from 'vscode';
import { getGoogleSheetsExportService } from '../services/googleSheetsExportService';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';
import { SqlFile } from '../models/SqlFile';

/**
 * ViewModel for SQL to Sheets export functionality
 */
export class SqlSheetsExportViewModel {
    /**
     * Export the current SQL query to Google Sheets
     * @param sqlQuery The SQL query to export
     */
    public async exportQueryToSheets(
        sqlQuery: string,
    ): Promise<void> {
        try {
            // Get active text editor
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                throw new Error('No active editor found');
            }

            // Use SqlFile to parse the configuration
            const sqlFile = new SqlFile(editor.document);
            const query = sqlFile.getQueryAt(editor.selection.active);
            let config = query?.config ?? new SqlSheetConfiguration();

            // Prompt for any missing required configuration
            const completeConfig = await this._promptForMissingConfig(config);
            if (!completeConfig) {
                return; // User cancelled
            }

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(sqlQuery, completeConfig);
        } catch (err) {
            vscode.window.showErrorMessage(
                `Failed to export SQL query to Google Sheets: ${err instanceof Error ? err.message : String(err)}`
            );
            console.error(err);
        }
    }

    /**
     * Export the active SQL file to Google Sheets
     */
    public async exportActiveFileToSheets(): Promise<void> {
        try {
            // Get active text editor
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                throw new Error('No active editor found');
            }

            // Check if the file is an SQL file
            if (editor.document.languageId !== 'sql') {
                throw new Error('Active file is not an SQL file');
            }

            // Get the file path and text content
            const sqlContent = editor.document.getText();
            const sqlFile = new SqlFile(editor.document);
            // For the whole file, we can take the config from the first query
            let config = sqlFile.queries[0]?.config ?? new SqlSheetConfiguration();

            // Prompt for any missing required configuration
            const completeConfig = await this._promptForMissingConfig(config, editor.document.fileName);
            if (!completeConfig) {
                return; // User cancelled
            }

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(
                sqlContent,
                completeConfig
            );
        } catch (err) {
            vscode.window.showErrorMessage(
                `Failed to export SQL file to Google Sheets: ${err instanceof Error ? err.message : String(err)}`
            );
            console.error(err);
        }
    }

    /**
     * Export a selected SQL query to Google Sheets
     */
    public async exportSelectionToSheets(): Promise<void> {
        try {
            // Get active text editor
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                throw new Error('No active editor found');
            }

            // Get the selected text or the entire document if no selection
            const selection = editor.selection;
            let sqlQuery: string;

            if (selection.isEmpty) {
                sqlQuery = editor.document.getText();
            } else {
                sqlQuery = editor.document.getText(selection);
            }

            if (!sqlQuery.trim()) {
                throw new Error('No SQL query to export');
            }

            // Use SqlFile to parse the configuration from the active query
            const sqlFile = new SqlFile(editor.document);
            const query = sqlFile.getQueryAt(editor.selection.active);
            let config = query?.config ?? new SqlSheetConfiguration();

            // Prompt for any missing required configuration
            const completeConfig = await this._promptForMissingConfig(config, 'Selected Query');
            if (!completeConfig) {
                return; // User cancelled
            }

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(
                sqlQuery,
                completeConfig
            );
        } catch (err) {
            vscode.window.showErrorMessage(
                `Failed to export SQL selection to Google Sheets: ${err instanceof Error ? err.message : String(err)}`
            );
            console.error(err);
        }
    }

    /**
     * Prompts the user for any missing required configuration values.
     * @param config The partially filled configuration.
     * @param defaultTableTitle Default title for the table.
     * @returns The completed configuration or undefined if the user cancels.
     */
    private async _promptForMissingConfig(
        config: SqlSheetConfiguration,
        defaultTableTitle: string = 'SQL Query Result'
    ): Promise<SqlSheetConfiguration | undefined> {
        let {
            spreadsheet_id,
            sheet_name,
            start_cell,
            name,
            transpose,
            data_only,
        } = config;

        if (!spreadsheet_id) {
            spreadsheet_id = await vscode.window.showInputBox({
                prompt: 'Enter the Google Sheets spreadsheet ID',
                validateInput: value => value ? null : 'Spreadsheet ID is required'
            });
            if (spreadsheet_id === undefined) { return undefined; }
        }

        if (!sheet_name) {
            sheet_name = await vscode.window.showInputBox({
                prompt: 'Enter the sheet name or ID',
                validateInput: value => value ? null : 'Sheet name or ID is required'
            });
            if (sheet_name === undefined) { return undefined; }
        }

        if (!start_cell) {
            start_cell = await vscode.window.showInputBox({
                prompt: 'Enter the start cell',
                value: 'A1',
                validateInput: value => value && /^[A-Za-z]+[0-9]+$/.test(value) ? null : 'Invalid cell reference'
            });
            if (start_cell === undefined) { return undefined; }
        }

        return new SqlSheetConfiguration(
            spreadsheet_id,
            sheet_name,
            start_cell,
            config.start_named_range,
            name || defaultTableTitle,
            config.table_name,
            config.pre_file,
            transpose,
            data_only,
            config.skip
        );
    }
}

// Singleton instance
let sqlSheetsExportViewModel: SqlSheetsExportViewModel | undefined;

/**
 * Get the SQL Sheets Export ViewModel instance
 */
export function getSqlSheetsExportViewModel(): SqlSheetsExportViewModel {
    if (!sqlSheetsExportViewModel) {
        sqlSheetsExportViewModel = new SqlSheetsExportViewModel();
    }
    return sqlSheetsExportViewModel;
}