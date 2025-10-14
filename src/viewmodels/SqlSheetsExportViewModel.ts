import * as vscode from 'vscode';
import { createHash } from 'crypto';
import * as path from 'path';
import { promises as fs } from 'fs';
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

export interface ExportOptions {
    executeDependencies?: boolean;
    executedPreFiles?: Set<string>;
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
        options: ExportOptions = {}
    ): Promise<ExportResult> {
        try {
            if (!sqlQuery) {
                throw new Error('No SQL query provided for export');
            }

            if (sqlQuery.config.skip) {
                logger.info('Skipping query export because the configuration is marked to skip.');
                return ExportResult.SkippedMissingConfig;
            }

            const executeDependencies = options.executeDependencies ?? false;
            const executedPreFiles = options.executedPreFiles ?? new Set<string>();

            const hasRequiredConfig = this._hasRequiredConfig(sqlQuery.config);
            const isCreateStatement = this._isCreateStatement(sqlQuery.queryText);

            if (!hasRequiredConfig) {
                if (isCreateStatement) {
                    await this._executeCreateStatement(sqlQuery.queryText);
                    return ExportResult.ExecutedCreate;
                }

                logger.info(
                    'Skipping query export because spreadsheet_id, sheet identifier, and a start cell or named range are required.',
                    { audience: ['support'] }
                );
                return ExportResult.SkippedMissingConfig;
            }

            if (executeDependencies && sqlQuery.config.pre_file) {
                const resolvedPreFile = this._resolvePreFilePath(sqlQuery, sqlQuery.config.pre_file);
                if (!resolvedPreFile) {
                    vscode.window.showErrorMessage(`Unable to resolve pre_file path: ${sqlQuery.config.pre_file}`);
                    return ExportResult.UserCancelled;
                }

                if (!executedPreFiles.has(resolvedPreFile)) {
                    await this._executePreFile(resolvedPreFile);
                    executedPreFiles.add(resolvedPreFile);
                }
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
            return ExportResult.UserCancelled;
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
            sheet_id,
            start_cell,
            name,
            transpose,
            data_only,
        } = config;
        let startNamedRange = config.start_named_range;

        if (sheet_name) {
            const parsedExistingSheet = SqlSheetConfiguration.parseSheetNameParameter(sheet_name);
            if (parsedExistingSheet.sheetId !== undefined && sheet_id === undefined) {
                sheet_id = parsedExistingSheet.sheetId;
            }
            if (parsedExistingSheet.sheetName) {
                sheet_name = parsedExistingSheet.sheetName;
            }
        }

        if (!spreadsheet_id) {
            spreadsheet_id = await vscode.window.showInputBox({
                prompt: 'Enter the Google Sheets spreadsheet ID',
                validateInput: value => value ? null : 'Spreadsheet ID is required'
            });
            if (spreadsheet_id === undefined) { return undefined; }
        }

        if (!sheet_name && sheet_id === undefined) {
            const sheetInput = await vscode.window.showInputBox({
                prompt: 'Enter the sheet identifier (ID or "ID | Name")',
                value: SqlSheetConfiguration.formatSheetNameParameter(sheet_id, sheet_name),
                validateInput: value => {
                    if (!value || value.trim().length === 0) {
                        return 'Sheet ID or name is required';
                    }
                    return null;
                }
            });
            if (sheetInput === undefined) { return undefined; }

            const parsedSheet = SqlSheetConfiguration.parseSheetNameParameter(sheetInput);
            sheet_id = parsedSheet.sheetId;
            sheet_name = parsedSheet.sheetName ?? (!parsedSheet.sheetId ? sheetInput : undefined);
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
            sheet_id,
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
        const hasSheetIdentifier = config.hasSheetIdentifier();
        const hasStartLocation = config.hasStartLocation();

        return hasSpreadsheetId && hasSheetIdentifier && hasStartLocation;
    }

    private _isCreateStatement(queryText: string): boolean {
        const sanitizedLines = queryText
            .replace(/\/\*[\s\S]*?\*\//g, '')
            .split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0 && !line.startsWith('--'));

        for (const line of sanitizedLines) {
            if (this._isIgnorableLeadingStatement(line)) {
                continue;
            }

            if (/^create\b/i.test(line)) {
                return true;
            }

            break;
        }

        const flattened = sanitizedLines.join(' ');
        const createMatch = /\bcreate\b/i.exec(flattened);
        if (!createMatch) {
            return false;
        }

        const textBeforeCreate = flattened.slice(0, createMatch.index).toLowerCase();
        if (textBeforeCreate.trim().length === 0) {
            return true;
        }

        const disallowedLeadingKeywords = /\b(select|insert|update|delete|merge|copy|call|with|begin|explain)\b/;
        if (disallowedLeadingKeywords.test(textBeforeCreate)) {
            return false;
        }

        return true;
    }

    private _isIgnorableLeadingStatement(line: string): boolean {
        return /^(?:use|set|alter\s+(?:session|warehouse|database|schema|user)|show|describe)\b/i.test(line);
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
            config.sheet_id,
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
        const sheetIdentifier = config.sheet_id !== undefined
            ? config.sheet_id.toString()
            : (config.sheet_name ?? '');
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

        const document = await vscode.workspace.openTextDocument(query.documentUri);
        const sqlFile = new SqlFile(document);
        const matchingQuery = sqlFile.queries.find(q => q.startOffset === query.startOffset && q.endOffset === query.endOffset);
        if (!matchingQuery) {
            logger.warn('Failed to locate matching query for parameter update.', { audience: ['developer'] });
            return;
        }

        if (newCombinedValue !== existingCombinedValue) {
            await sqlFile.updateParameter(matchingQuery, 'start_cell', newCombinedValue);
            logger.info(`Updated start_cell parameter to "${newCombinedValue}" in SQL file.`, { audience: ['developer'] });
        }

        const newSheetParameter = SqlSheetConfiguration.formatSheetNameParameter(
            uploadResult.sheetId ?? query.config.sheet_id,
            uploadResult.sheetName ?? query.config.sheet_name
        );
        const existingSheetParameter = SqlSheetConfiguration.formatSheetNameParameter(
            query.config.sheet_id,
            query.config.sheet_name
        );

        if (newSheetParameter !== existingSheetParameter && newSheetParameter.length > 0) {
            await sqlFile.updateParameter(matchingQuery, 'sheet_name', newSheetParameter);
            logger.info(`Updated sheet_name parameter to "${newSheetParameter}" in SQL file.`, { audience: ['developer'] });
        }
    }

    private _resolvePreFilePath(query: SqlQuery, preFile: string): string | undefined {
        try {
            if (!preFile || preFile.trim().length === 0) {
                return undefined;
            }

            const queryFilePath = query.documentUri.fsPath;
            if (!queryFilePath) {
                return undefined;
            }

            if (path.isAbsolute(preFile)) {
                return path.normalize(preFile);
            }

            return path.normalize(path.resolve(path.dirname(queryFilePath), preFile));
        } catch (err) {
            logger.warn('Failed to resolve pre_file path', { data: err });
            return undefined;
        }
    }

    private async _executePreFile(filePath: string): Promise<void> {
        try {
            await fs.access(filePath);
        } catch {
            throw new Error(`pre_file not found: ${filePath}`);
        }

        let fileContents: string;
        try {
            fileContents = await fs.readFile(filePath, 'utf-8');
        } catch (err) {
            throw new Error(`Failed to read pre_file ${filePath}: ${err instanceof Error ? err.message : String(err)}`);
        }

        if (fileContents.trim().length === 0) {
            logger.info(`Pre-file ${filePath} is empty. Skipping execution.`, { audience: ['developer'] });
            return;
        }

        const snowflakeService = getSnowflakeService();
        if (!snowflakeService.isConfigured()) {
            throw new Error('Snowflake connection is not configured. Cannot execute pre_file.');
        }

        await vscode.window.withProgress({
            location: vscode.ProgressLocation.Notification,
            title: `Executing pre-file: ${path.basename(filePath)}`,
            cancellable: false
        }, async () => {
            logger.info(`Executing pre_file ${filePath}`, { audience: ['developer'] });
            await snowflakeService.executeQuery(fileContents);
            logger.info(`Completed pre_file ${filePath}`, { audience: ['developer'] });
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
