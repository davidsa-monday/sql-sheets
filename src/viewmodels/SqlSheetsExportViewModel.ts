import * as vscode from 'vscode';
import { getGoogleSheetsExportService } from '../services/googleSheetsExportService';
import { SqlQuery } from '../models/SqlQuery';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';

/**
 * ViewModel for SQL to Sheets export functionality
 */
export class SqlSheetsExportViewModel {
    /**
     * Export the current SQL query to Google Sheets
     * @param sqlQuery The SQL query to export
     */
    public async exportQueryToSheets(
        sqlQuery: SqlQuery,
    ): Promise<void> {
        try {
            if (!sqlQuery) {
                throw new Error('No SQL query provided for export');
            }

            if (sqlQuery.config.skip) {
                console.log('Skipping query export because the configuration is marked to skip.');
                return;
            }

            const completedConfig = await this._promptForMissingConfig(
                sqlQuery.config,
                sqlQuery.config.name ?? 'SQL Query Result'
            );
            if (!completedConfig) {
                return;
            }

            // Get the export service and perform the export
            const exportService = getGoogleSheetsExportService();
            await exportService.exportQueryToSheet(sqlQuery.queryText, completedConfig);
        } catch (err) {
            vscode.window.showErrorMessage(
                `Failed to export SQL query to Google Sheets: ${err instanceof Error ? err.message : String(err)}`
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
