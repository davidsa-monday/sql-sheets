import * as vscode from 'vscode';
import { getGoogleSheetsExportService } from '../services/googleSheetsExportService';
import { getSnowflakeService } from '../services/snowflakeService';
import { SqlQuery } from '../models/SqlQuery';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';
import { getLogger } from '../services/loggingService';

const logger = getLogger();

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
                logger.info('Skipping query export because the configuration is marked to skip.');
                return;
            }

            const hasRequiredConfig = this._hasRequiredConfig(sqlQuery.config);
            const isCreateStatement = this._isCreateStatement(sqlQuery.queryText);

            if (!hasRequiredConfig) {
                if (isCreateStatement) {
                    await this._executeCreateStatement(sqlQuery.queryText);
                } else {
                    logger.info(
                        'Skipping query export because spreadsheet_id, sheet_name, and start_cell are required.', { audience: ['support'] }
                    );
                }
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
            const errorMessage = err instanceof Error ? err.message : String(err);
            logger.error('Failed to export SQL query via view model', { data: err });
            vscode.window.showErrorMessage(
                `Failed to export SQL query to Google Sheets: ${errorMessage}`
            );
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

    private _hasRequiredConfig(config: SqlSheetConfiguration): boolean {
        const hasSpreadsheetId = typeof config.spreadsheet_id === 'string' && config.spreadsheet_id.trim().length > 0;
        const hasSheetName = typeof config.sheet_name === 'string' && config.sheet_name.trim().length > 0;
        const hasStartCell = typeof config.start_cell === 'string' && config.start_cell.trim().length > 0;

        return hasSpreadsheetId && hasSheetName && hasStartCell;
    }

    private _isCreateStatement(queryText: string): boolean {
        const strippedQuery = queryText
            .replace(/\/\*[\s\S]*?\*\//g, '') // Remove block comments
            .split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0 && !line.startsWith('--'))
            .join(' ');

        return /^create\b/i.test(strippedQuery);
    }

    private async _executeCreateStatement(queryText: string): Promise<void> {
        const snowflakeService = getSnowflakeService();

        if (!snowflakeService.isConfigured()) {
            vscode.window.showErrorMessage('Snowflake connection is not configured yet.');
            return;
        }

        await vscode.window.withProgress({
            location: vscode.ProgressLocation.Notification,
            title: 'Executing CREATE statement in Snowflake...',
            cancellable: false
        }, async () => {
            try {
                await snowflakeService.executeQuery(queryText);
                vscode.window.showInformationMessage('CREATE statement executed successfully.');
            } catch (err) {
                const errorMessage = err instanceof Error ? err.message : String(err);
                logger.error('Failed to execute CREATE statement', { data: err });
                vscode.window.showErrorMessage(`Failed to execute CREATE statement: ${errorMessage}`);
            }
        });
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
