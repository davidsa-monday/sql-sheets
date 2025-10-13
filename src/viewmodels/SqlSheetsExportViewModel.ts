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

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(sqlQuery, config);
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

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(
                sqlContent,
                config
            );
        } catch (err) {
            vscode.window.showErrorMessage(
                `Failed to export SQL file to Google Sheets: ${err instanceof Error ? err.message : String(err)}`
            );
            console.error(err);
        }
    }

    // ...existing code...
    // ...existing code...
}

/**
 * Prompts the user for any missing required configuration values.
 * @param config The partially filled configuration.
 * @param defaultTableTitle Default title for the table.
 * @returns The completed configuration or undefined if the user cancels.
 */
// ...existing code...
// ...existing code...

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