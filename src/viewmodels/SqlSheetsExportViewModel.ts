import * as vscode from 'vscode';
import { getGoogleSheetsExportService } from '../services/googleSheetsExportService';

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
            // Prompt for the export configuration
            const exportConfig = await this._promptForExportConfig();
            if (!exportConfig) {
                return; // User cancelled
            }

            const { spreadsheetId, sheetName, startCell, transpose, tableTitle } = exportConfig;

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(
                sqlQuery,
                spreadsheetId,
                sheetName,
                startCell,
                {
                    transpose,
                    tableTitle
                }
            );
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
            const filePath = editor.document.uri.fsPath;
            const sqlContent = editor.document.getText();

            // Prompt for the export configuration
            const exportConfig = await this._promptForExportConfig(editor.document.fileName);
            if (!exportConfig) {
                return; // User cancelled
            }

            const { spreadsheetId, sheetName, startCell, transpose, tableTitle } = exportConfig;

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(
                sqlContent,
                spreadsheetId,
                sheetName,
                startCell,
                {
                    transpose,
                    tableTitle
                }
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

            // Prompt for the export configuration
            const exportConfig = await this._promptForExportConfig('Selected Query');
            if (!exportConfig) {
                return; // User cancelled
            }

            const { spreadsheetId, sheetName, startCell, transpose, tableTitle } = exportConfig;

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(
                sqlQuery,
                spreadsheetId,
                sheetName,
                startCell,
                {
                    transpose,
                    tableTitle
                }
            );
        } catch (err) {
            vscode.window.showErrorMessage(
                `Failed to export SQL selection to Google Sheets: ${err instanceof Error ? err.message : String(err)}`
            );
            console.error(err);
        }
    }

    /**
     * Prompt the user for the export configuration
     * @param defaultTableTitle Default title for the table
     * @returns The export configuration or undefined if cancelled
     */
    private async _promptForExportConfig(defaultTableTitle: string = 'SQL Query Result'): Promise<{
        spreadsheetId: string;
        sheetName: string;
        startCell: string;
        transpose: boolean;
        tableTitle: string;
    } | undefined> {
        // Get the previous configuration if available
        const config = vscode.workspace.getConfiguration('sql-sheets.export');
        const previousSpreadsheetId = config.get<string>('lastSpreadsheetId') || '';
        const previousSheetName = config.get<string>('lastSheetName') || '';
        const previousStartCell = config.get<string>('lastStartCell') || 'A1';

        // Prompt for spreadsheet ID
        const spreadsheetId = await vscode.window.showInputBox({
            prompt: 'Enter the Google Sheets spreadsheet ID',
            value: previousSpreadsheetId,
            placeHolder: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
            validateInput: (value) => {
                if (!value || !value.trim()) {
                    return 'Spreadsheet ID is required';
                }
                return null;
            }
        });

        if (!spreadsheetId) {
            return undefined;
        }

        // Prompt for sheet name
        const sheetName = await vscode.window.showInputBox({
            prompt: 'Enter the sheet name or ID',
            value: previousSheetName,
            placeHolder: 'Sheet1',
            validateInput: (value) => {
                if (!value || !value.trim()) {
                    return 'Sheet name or ID is required';
                }
                return null;
            }
        });

        if (!sheetName) {
            return undefined;
        }

        // Prompt for start cell
        const startCell = await vscode.window.showInputBox({
            prompt: 'Enter the start cell',
            value: previousStartCell,
            placeHolder: 'A1',
            validateInput: (value) => {
                if (!value || !value.trim()) {
                    return 'Start cell is required';
                }
                // Check if the cell reference is valid (e.g. A1, B10)
                if (!value.match(/^[A-Za-z]+[0-9]+$/)) {
                    return 'Invalid cell reference. Please use format like A1, B10, etc.';
                }
                return null;
            }
        });

        if (!startCell) {
            return undefined;
        }

        // Prompt for table title
        const tableTitle = await vscode.window.showInputBox({
            prompt: 'Enter the table title',
            value: defaultTableTitle,
            placeHolder: 'SQL Query Result'
        });

        if (tableTitle === undefined) {
            return undefined;
        }

        // Prompt for transpose option
        const transposeOptions = ['No', 'Yes'];
        const transposeSelection = await vscode.window.showQuickPick(transposeOptions, {
            placeHolder: 'Transpose the data?'
        });

        if (!transposeSelection) {
            return undefined;
        }

        const transpose = transposeSelection === 'Yes';

        // Save the configuration for next time
        await config.update('lastSpreadsheetId', spreadsheetId, vscode.ConfigurationTarget.Global);
        await config.update('lastSheetName', sheetName, vscode.ConfigurationTarget.Global);
        await config.update('lastStartCell', startCell, vscode.ConfigurationTarget.Global);

        return {
            spreadsheetId,
            sheetName,
            startCell,
            transpose,
            tableTitle: tableTitle || defaultTableTitle
        };
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