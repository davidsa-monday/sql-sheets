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
            const executeDependencies = options.executeDependencies ?? false;
            const executedPreFiles = options.executedPreFiles ?? new Set<string>();
            const destinationEntries = sqlQuery.destinations.map((destination, index) => ({ destination, index }));
            const nonSkippedDestinations = destinationEntries.filter(entry => !entry.destination.config.skip);

            if (nonSkippedDestinations.length === 0) {
                if (this._isCreateStatement(sqlQuery.queryText)) {
                    await this._executeCreateStatement(sqlQuery.queryText);
                    return ExportResult.ExecutedCreate;
                }

                logger.info('Skipping query export because all destinations are marked to skip.', { audience: ['support'] });
                return ExportResult.SkippedMissingConfig;
            }

            const completedDestinations: Array<{ index: number; config: SqlSheetConfiguration }> = [];
            for (const entry of nonSkippedDestinations) {
                const baseConfig = entry.destination.config;
                const defaultTitle = baseConfig.name
                    ?? sqlQuery.config.name
                    ?? `SQL Query Result (${entry.index + 1})`;

                let configured: SqlSheetConfiguration = baseConfig;
                if (!this._hasRequiredConfig(baseConfig)) {
                    const prompted = await this._promptForMissingConfig(baseConfig, defaultTitle);
                    if (!prompted) {
                        return ExportResult.UserCancelled;
                    }
                    configured = prompted;
                }

                const effectiveConfig = this._ensureNamedRange(sqlQuery, entry.index, configured);
                completedDestinations.push({ index: entry.index, config: effectiveConfig });
            }

            if (completedDestinations.length === 0) {
                if (this._isCreateStatement(sqlQuery.queryText)) {
                    await this._executeCreateStatement(sqlQuery.queryText);
                    return ExportResult.ExecutedCreate;
                }

                logger.info(
                    'Skipping query export because spreadsheet_id, sheet identifier, and a start cell or named range are required.',
                    { audience: ['support'] }
                );
                return ExportResult.SkippedMissingConfig;
            }

            if (executeDependencies) {
                const preFileSet = new Set<string>();
                for (const destination of completedDestinations) {
                    for (const preFile of destination.config.pre_files) {
                        preFileSet.add(preFile);
                    }
                }

                if (preFileSet.size > 0) {
                    const ranDependencies = await this.runPreFiles(sqlQuery.documentUri, Array.from(preFileSet), executedPreFiles);
                    if (!ranDependencies) {
                        return ExportResult.UserCancelled;
                    }
                }
            }

            const snowflakeService = getSnowflakeService();
            if (!snowflakeService.isConfigured()) {
                throw new Error('Snowflake connection is not configured');
            }
            const exportService = getGoogleSheetsExportService();

            const results = await snowflakeService.executeQuery(sqlQuery.queryText);

            for (const destination of completedDestinations) {
                const uploadResult = await exportService.exportResultsToSheet(results, sqlQuery.queryText, destination.config);
                if (uploadResult) {
                    await this._updateStartCellParameterForDestination(sqlQuery, destination.index, uploadResult);
                }
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
            config.pre_files,
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
        destinationIndex: number,
        config: SqlSheetConfiguration
    ): SqlSheetConfiguration {
        const hasNamedRange = typeof config.start_named_range === 'string' && config.start_named_range.trim().length > 0;
        if (hasNamedRange) {
            return config;
        }

        const generatedName = this._generateDeterministicNamedRange(query, destinationIndex, config);
        logger.info(`Generated named range "${generatedName}" for export run.`, { audience: ['developer'] });

        return new SqlSheetConfiguration(
            config.spreadsheet_id,
            config.sheet_name,
            config.sheet_id,
            config.start_cell,
            generatedName,
            config.name,
            config.table_name,
            config.pre_files,
            config.transpose,
            config.data_only,
            config.skip
        );
    }

    private _generateDeterministicNamedRange(
        query: SqlQuery,
        destinationIndex: number,
        config: SqlSheetConfiguration
    ): string {
        const editor = vscode.window.activeTextEditor;
        const sqlFilePath = editor?.document?.uri.fsPath ?? 'unknown-sql-file';
        const tableTitle = config.name ?? config.table_name ?? 'SQL Query Result';
        const sheetIdentifier = config.sheet_id !== undefined
            ? config.sheet_id.toString()
            : (config.sheet_name ?? '');
        const startCell = config.start_cell ?? '';

        const seed = `${sqlFilePath}|${tableTitle}|${sheetIdentifier}|${startCell}|${query.startOffset}|${query.endOffset}|dest:${destinationIndex}`;
        const hash = createHash('md5').update(seed).digest('hex').slice(0, 10);
        return `SQLS_${hash}`;
    }

    private async _updateStartCellParameterForDestination(
        query: SqlQuery,
        destinationIndex: number,
        uploadResult: SheetUploadResult
    ): Promise<void> {
        const destination = query.destinations[destinationIndex];
        if (!destination) {
            return;
        }

        const destinationConfig = destination.config;
        const namedRangeFromResult = uploadResult.startNamedRange ?? destinationConfig.start_named_range;
        if (!namedRangeFromResult) {
            return;
        }

        const newCombinedValue = SqlSheetConfiguration.formatStartCellParameter(
            namedRangeFromResult,
            uploadResult.startCell
        );

        const existingCombinedValue = SqlSheetConfiguration.formatStartCellParameter(
            destinationConfig.start_named_range,
            destinationConfig.start_cell
        );

        const document = await vscode.workspace.openTextDocument(query.documentUri);
        const sqlFile = new SqlFile(document);
        const matchingQuery = sqlFile.queries.find(q => q.startOffset === query.startOffset && q.endOffset === query.endOffset);
        if (!matchingQuery) {
            logger.warn('Failed to locate matching query for parameter update.', { audience: ['developer'] });
            return;
        }

        const targetDestination = matchingQuery.destinations[destinationIndex];
        if (!targetDestination) {
            logger.warn('Destination metadata not found during parameter update.', { audience: ['developer'] });
            return;
        }

        if (newCombinedValue !== existingCombinedValue) {
            if (destinationIndex === 0) {
                await sqlFile.updateParameter(matchingQuery, 'start_cell', newCombinedValue);
                logger.info(`Updated start_cell parameter to "${newCombinedValue}" in SQL file.`, { audience: ['developer'] });
            } else {
                const rangeInfo = targetDestination.parameterRanges.get('start_cell');
                if (rangeInfo) {
                    const edit = new vscode.WorkspaceEdit();
                    const range = new vscode.Range(document.positionAt(rangeInfo.start), document.positionAt(rangeInfo.end));
                    edit.replace(document.uri, range, `--start_cell: ${newCombinedValue}`);
                    await vscode.workspace.applyEdit(edit);
                    logger.info(`Updated start_cell parameter to "${newCombinedValue}" in SQL file (destination ${destinationIndex + 1}).`, { audience: ['developer'] });
                } else {
                    logger.warn('Unable to update start_cell for non-primary destination because no parameter range was recorded.', { audience: ['developer'] });
                }
            }
        }

        const newSheetParameter = SqlSheetConfiguration.formatSheetNameParameter(
            uploadResult.sheetId ?? destinationConfig.sheet_id,
            uploadResult.sheetName ?? destinationConfig.sheet_name
        );
        const existingSheetParameter = SqlSheetConfiguration.formatSheetNameParameter(
            destinationConfig.sheet_id,
            destinationConfig.sheet_name
        );

        if (newSheetParameter !== existingSheetParameter && newSheetParameter.length > 0) {
            if (destinationIndex === 0) {
                await sqlFile.updateParameter(matchingQuery, 'sheet_name', newSheetParameter);
                logger.info(`Updated sheet_name parameter to "${newSheetParameter}" in SQL file.`, { audience: ['developer'] });
            } else {
                const rangeInfo = targetDestination.parameterRanges.get('sheet_name');
                if (rangeInfo) {
                    const edit = new vscode.WorkspaceEdit();
                    const range = new vscode.Range(document.positionAt(rangeInfo.start), document.positionAt(rangeInfo.end));
                    edit.replace(document.uri, range, `--sheet_name: ${newSheetParameter}`);
                    await vscode.workspace.applyEdit(edit);
                    logger.info(`Updated sheet_name parameter to "${newSheetParameter}" in SQL file (destination ${destinationIndex + 1}).`, { audience: ['developer'] });
                } else {
                    logger.warn('Unable to update sheet_name for non-primary destination because no parameter range was recorded.', { audience: ['developer'] });
                }
            }
        }
    }

    public async runPreFiles(
        documentUri: vscode.Uri,
        preFiles: readonly string[],
        executedPreFiles: Set<string>
    ): Promise<boolean> {
        if (!preFiles || preFiles.length === 0) {
            return true;
        }

        const processing = new Set<string>();

        try {
            await this._runPreFilesRecursive(documentUri, preFiles, executedPreFiles, processing);
            return true;
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            vscode.window.showErrorMessage(message);
            logger.error('Failed to execute pre_file dependency chain', { data: err });
            return false;
        }
    }

    private async _runPreFilesRecursive(
        documentUri: vscode.Uri,
        preFiles: readonly string[],
        executedPreFiles: Set<string>,
        processing: Set<string>
    ): Promise<void> {
        for (const rawPreFile of preFiles) {
            const trimmedPreFile = rawPreFile.trim();
            if (trimmedPreFile.length === 0) {
                continue;
            }

            const resolvedPreFile = this._resolvePreFilePath(documentUri, trimmedPreFile);
            if (!resolvedPreFile) {
                throw new Error(`Unable to resolve pre_file path: ${trimmedPreFile}`);
            }

            await this._executePreFile(resolvedPreFile, executedPreFiles, processing);
        }
    }

    private _resolvePreFilePath(documentUri: vscode.Uri, preFile: string): string | undefined {
        try {
            if (!preFile || preFile.trim().length === 0) {
                return undefined;
            }

            const trimmedPreFile = preFile.trim();

            if (path.isAbsolute(trimmedPreFile)) {
                return path.normalize(trimmedPreFile);
            }

            const workspaceFolder = vscode.workspace.getWorkspaceFolder(documentUri)
                ?? vscode.workspace.workspaceFolders?.[0];

            if (workspaceFolder) {
                return path.normalize(path.resolve(workspaceFolder.uri.fsPath, trimmedPreFile));
            }

            const queryFilePath = documentUri.fsPath;
            if (!queryFilePath) {
                return undefined;
            }

            return path.normalize(path.resolve(path.dirname(queryFilePath), trimmedPreFile));
        } catch (err) {
            logger.warn('Failed to resolve pre_file path', { data: err });
            return undefined;
        }
    }

    private async _executePreFile(
        filePath: string,
        executedPreFiles: Set<string>,
        processing: Set<string>
    ): Promise<void> {
        const normalizedPath = path.normalize(filePath);
        if (executedPreFiles.has(normalizedPath)) {
            return;
        }

        if (processing.has(normalizedPath)) {
            throw new Error(`Circular pre_file dependency detected involving ${normalizedPath}`);
        }

        processing.add(normalizedPath);

        try {
            await fs.access(normalizedPath);
        } catch {
            processing.delete(normalizedPath);
            throw new Error(`pre_file not found: ${normalizedPath}`);
        }

        let fileContents: string;
        try {
            fileContents = await fs.readFile(normalizedPath, 'utf-8');
        } catch (err) {
            processing.delete(normalizedPath);
            throw new Error(`Failed to read pre_file ${normalizedPath}: ${err instanceof Error ? err.message : String(err)}`);
        }

        const nestedPreFiles = this._extractPreFileDirectives(fileContents);
        if (nestedPreFiles.length > 0) {
            const preFileUri = vscode.Uri.file(normalizedPath);
            await this._runPreFilesRecursive(preFileUri, nestedPreFiles, executedPreFiles, processing);
        }

        if (fileContents.trim().length === 0) {
            logger.info(`Pre-file ${normalizedPath} is empty. Skipping execution.`, { audience: ['developer'] });
            processing.delete(normalizedPath);
            executedPreFiles.add(normalizedPath);
            return;
        }

        if (this._shouldSkipSelectStatement(fileContents)) {
            logger.info(`Skipping pre-file ${normalizedPath} because it only contains a SELECT statement.`, { audience: ['developer'] });
            processing.delete(normalizedPath);
            executedPreFiles.add(normalizedPath);
            return;
        }

        const snowflakeService = getSnowflakeService();
        if (!snowflakeService.isConfigured()) {
            processing.delete(normalizedPath);
            throw new Error('Snowflake connection is not configured. Cannot execute pre_file.');
        }

        try {
            await vscode.window.withProgress({
                location: vscode.ProgressLocation.Notification,
                title: `Executing pre-file: ${path.basename(normalizedPath)}`,
                cancellable: false
            }, async () => {
                logger.info(`Executing pre_file ${normalizedPath}`, { audience: ['developer'] });
                await snowflakeService.executeQuery(fileContents);
                logger.info(`Completed pre_file ${normalizedPath}`, { audience: ['developer'] });
            });
            executedPreFiles.add(normalizedPath);
        } finally {
            processing.delete(normalizedPath);
        }
    }

    private _extractPreFileDirectives(fileContents: string): string[] {
        const results: string[] = [];
        const directiveRegex = /^--\s*pre_file:[\t ]*(.*)$/gim;
        let match: RegExpExecArray | null;

        while ((match = directiveRegex.exec(fileContents)) !== null) {
            const entry = match[1]?.trim() ?? '';
            if (entry.length > 0) {
                results.push(entry);
            }
        }

        return results;
    }

    private _shouldSkipSelectStatement(sqlText: string): boolean {
        const withoutBlockComments = sqlText.replace(/\/\*[\s\S]*?\*\//g, ' ');
        const statements = withoutBlockComments.split(';');

        for (const rawStatement of statements) {
            const trimmedStatement = rawStatement.trim();
            if (trimmedStatement.length === 0) {
                continue;
            }

            const sanitizedLines = trimmedStatement
                .split('\n')
                .map(line => line.trim())
                .filter(line => line.length > 0 && !line.startsWith('--'));

            if (sanitizedLines.length === 0) {
                continue;
            }

            const firstLine = sanitizedLines[0];
            if (this._isIgnorableLeadingStatement(firstLine)) {
                // Statements like USE or SET are treated as ignorable and don't trigger skipping.
                continue;
            }

            return /^(select|with)\b/i.test(firstLine);
        }

        return false;
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
