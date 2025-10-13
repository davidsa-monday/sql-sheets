import * as vscode from 'vscode';
import { createHash } from 'crypto';
import { getGoogleSheetsExportService } from '../services/googleSheetsExportService';
import { getSnowflakeService } from '../services/snowflakeService';
import { SqlQuery } from '../models/SqlQuery';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';
import type { SheetUploadResult } from '../services/googleSheetsService';
import { SqlFile } from '../models/SqlFile';
import { getLogger } from '../services/loggingService';

export enum ExportResult {
    Exported = 'exported',
    SkippedMissingConfig = 'skipped_missing_config',
    UserCancelled = 'user_cancelled',
    ExecutedCreate = 'executed_create'
}

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
    ): Promise<ExportResult> {
        try {
            if (!sqlQuery) {
                throw new Error('No SQL query provided for export');
            }

            if (sqlQuery.config.skip) {
                logger.info('Skipping query export because the configuration is marked to skip.');
                return ExportResult.SkippedMissingConfig;
            }

            const hasRequiredConfig = this._hasRequiredConfig(sqlQuery.config);
            const isCreateStatement = this._isCreateStatement(sqlQuery.queryText);

            if (!hasRequiredConfig) {
                if (isCreateStatement) {
                    await this._executeCreateStatement(sqlQuery.queryText);
                    return ExportResult.ExecutedCreate;
                }

                logger.info(
                    'Skipping query export because spreadsheet_id, sheet_name, and a start cell or named range are required.',
                    { audience: ['support'] }
                );
                return ExportResult.SkippedMissingConfig;
            }

            const completedConfig = await this._promptForMissingConfig(
                sqlQuery.config,
                sqlQuery.config.name ?? 'SQL Query Result'
            );
            if (!completedConfig) {
                return ExportResult.UserCancelled;
            }

            // Get the export service and perform the export
            const effectiveConfig = this._ensureNamedRange(sqlQuery, completedConfig);

            const exportService = getGoogleSheetsExportService();
            const uploadResult = await exportService.exportQueryToSheet(sqlQuery.queryText, effectiveConfig);

            if (uploadResult) {
                await this._updateStartCellParameter(sqlQuery, uploadResult);
            }

            return ExportResult.Exported;
        } catch (err) {
            const errorMessage = err instanceof Error ? err.message : String(err);
            logger.error('Failed to export SQL query via view model', { data: err });
            vscode.window.showErrorMessage(
                `Failed to export SQL query to Google Sheets: ${errorMessage}`
            );
            throw err;
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
        let startNamedRange = config.start_named_range;

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

        if (!start_cell && !startNamedRange) {
            const startInput = await vscode.window.showInputBox({
                prompt: 'Enter the start location (cell or "NamedRange | Cell")',
                value: 'A1',
                validateInput: value => {
                    if (!value || value.trim().length === 0) {
                        return 'A start location is required';
                    }
                    const parsed = SqlSheetConfiguration.parseStartCellParameter(value);
                    if (!parsed.startCell && !parsed.startNamedRange) {
                        return 'Provide a cell like "A1", a named range, or combine them as "MyRange | A1".';
                    }
                    return null;
                }
            });
            if (startInput === undefined) { return undefined; }

            const parsed = SqlSheetConfiguration.parseStartCellParameter(startInput);
            start_cell = parsed.startCell ?? (!parsed.startNamedRange ? startInput : undefined);
            startNamedRange = parsed.startNamedRange;
        }

        return new SqlSheetConfiguration(
            spreadsheet_id,
            sheet_name,
            start_cell,
            startNamedRange,
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
        const hasStartLocation = config.hasStartLocation();

        return hasSpreadsheetId && hasSheetName && hasStartLocation;
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

    private _ensureNamedRange(
        query: SqlQuery,
        config: SqlSheetConfiguration
    ): SqlSheetConfiguration {
        const hasNamedRange = typeof config.start_named_range === 'string' && config.start_named_range.trim().length > 0;
        if (hasNamedRange) {
            return config;
        }

        const generatedName = this._generateDeterministicNamedRange(query, config);
        logger.info(`Generated named range "${generatedName}" for export run.`, { audience: ['developer'] });

        return new SqlSheetConfiguration(
            config.spreadsheet_id,
            config.sheet_name,
            config.start_cell,
            generatedName,
            config.name,
            config.table_name,
            config.pre_file,
            config.transpose,
            config.data_only,
            config.skip
        );
    }

    private _generateDeterministicNamedRange(
        query: SqlQuery,
        config: SqlSheetConfiguration
    ): string {
        const editor = vscode.window.activeTextEditor;
        const sqlFilePath = editor?.document?.uri.fsPath ?? 'unknown-sql-file';
        const tableTitle = config.name ?? config.table_name ?? 'SQL Query Result';
        const sheetIdentifier = config.sheet_name ?? '';
        const startCell = config.start_cell ?? '';

        const seed = `${sqlFilePath}|${tableTitle}|${sheetIdentifier}|${startCell}|${query.startOffset}|${query.endOffset}`;
        const hash = createHash('md5').update(seed).digest('hex').slice(0, 10);
        return `SQLS_${hash}`;
    }

    private async _updateStartCellParameter(
        query: SqlQuery,
        uploadResult: SheetUploadResult
    ): Promise<void> {
        const namedRangeFromResult = uploadResult.startNamedRange ?? query.config.start_named_range;
        if (!namedRangeFromResult) {
            return; // Only update files when a named range is involved
        }

        const newCombinedValue = SqlSheetConfiguration.formatStartCellParameter(
            namedRangeFromResult,
            uploadResult.startCell
        );

        const existingCombinedValue = SqlSheetConfiguration.formatStartCellParameter(
            query.config.start_named_range,
            query.config.start_cell
        );

        if (newCombinedValue.length === 0 || newCombinedValue === existingCombinedValue) {
            return;
        }

        const editor = vscode.window.activeTextEditor;
        if (!editor || (editor.document.languageId !== 'sql' && editor.document.languageId !== 'snowflake-sql')) {
            return;
        }

        const sqlFile = new SqlFile(editor.document);
        await sqlFile.updateParameter(query, 'start_cell', newCombinedValue);
        logger.info(`Updated start_cell parameter to "${newCombinedValue}" in SQL file.`, { audience: ['developer'] });
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
