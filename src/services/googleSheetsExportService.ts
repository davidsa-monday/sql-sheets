import * as vscode from 'vscode';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';
import { getGoogleSheetsService } from './googleSheetsService';
import type { SheetUploadResult } from './googleSheetsService';
import { getSnowflakeService } from './snowflakeService';
import { getLogger } from './loggingService';

const logger = getLogger();

type PreparedUploadContext = {
    spreadsheetId: string;
    sheetIdentifier: string | number;
    startCell?: string;
    startNamedRange?: string;
    tableTitle?: string;
    tableName?: string;
    transpose: boolean;
    dataOnly: boolean;
};

/**
 * Service for exporting SQL query results to Google Sheets
 */
export class GoogleSheetsExportService {
    /**
     * Export SQL query results to a Google Sheet
     * @param sqlQuery The SQL query to execute
     * @param config The configuration for the SQL sheet
     * @returns A promise that resolves when the export is complete
     */
    public async exportQueryToSheet(
        sqlQuery: string,
        config: SqlSheetConfiguration
    ): Promise<SheetUploadResult | void> {
        if (config.skip) {
            logger.info('Skipping query execution because configuration is marked skip', { audience: ['developer'] });
            return;
        }

        try {
            const uploadContext = this._prepareUploadContext(sqlQuery, config);

            const snowflakeService = getSnowflakeService();
            if (!snowflakeService.isConfigured()) {
                throw new Error('Snowflake connection is not configured');
            }

            logger.info('Executing query in Snowflake...', { audience: ['developer'] });
            const results = await snowflakeService.executeQuery(sqlQuery);
            logger.info('Query executed successfully.', { audience: ['developer'] });

            return await this._uploadResultsToSheet(results, sqlQuery, uploadContext);
        } catch (err) {
            const errorMessage = err instanceof Error ? err.message : String(err);
            logger.error('Failed to export query to Google Sheet', { data: err });
            vscode.window.showErrorMessage(`Failed to export query to Google Sheet: ${errorMessage}`);
        }
    }

    public async exportResultsToSheet(
        rawResults: any[],
        sqlQuery: string,
        config: SqlSheetConfiguration
    ): Promise<SheetUploadResult | void> {
        if (config.skip) {
            logger.info('Skipping query export because configuration is marked skip', { audience: ['developer'] });
            return;
        }

        try {
            const uploadContext = this._prepareUploadContext(sqlQuery, config);
            return await this._uploadResultsToSheet(rawResults, sqlQuery, uploadContext);
        } catch (err) {
            const errorMessage = err instanceof Error ? err.message : String(err);
            logger.error('Failed to export prepared results to Google Sheet', { data: err });
            vscode.window.showErrorMessage(`Failed to export query to Google Sheet: ${errorMessage}`);
        }
    }

    private _prepareUploadContext(
        sqlQuery: string,
        config: SqlSheetConfiguration
    ): PreparedUploadContext {
        if (!sqlQuery) {
            throw new Error('SQL query is required');
        }

        const spreadsheetId = typeof config.spreadsheet_id === 'string'
            ? config.spreadsheet_id.trim()
            : '';
        if (!spreadsheetId) {
            throw new Error('Spreadsheet ID is required');
        }

        const sheetName = typeof config.sheet_name === 'string' ? config.sheet_name.trim() : undefined;
        const sheetIdentifier = config.sheet_id !== undefined ? config.sheet_id : sheetName;
        if (sheetIdentifier === undefined || (typeof sheetIdentifier === 'string' && sheetIdentifier.length === 0)) {
            throw new Error('Sheet name or ID is required');
        }

        if (typeof sheetIdentifier === 'string') {
            logger.info(`Using sheet name: "${sheetIdentifier}"`, { audience: ['developer'] });
        } else {
            logger.info(`Using sheet ID: ${sheetIdentifier}`, { audience: ['developer'] });
        }

        const cleanedStartNamedRange = typeof config.start_named_range === 'string'
            ? config.start_named_range.trim()
            : undefined;
        const cleanedStartCell = typeof config.start_cell === 'string'
            ? config.start_cell.trim()
            : undefined;

        if (!cleanedStartCell && !cleanedStartNamedRange) {
            throw new Error('A start cell or named range is required');
        }

        if (cleanedStartCell) {
            const validCellFormat = /^[A-Za-z]+[0-9]+$/;
            if (!validCellFormat.test(cleanedStartCell)) {
                throw new Error(`Start cell must be in the format like 'A1', 'B2', etc.`);
            }

            const columnPart = cleanedStartCell.match(/^[A-Za-z]+/)?.[0] || '';
            const rowPart = cleanedStartCell.match(/[0-9]+$/)?.[0] || '';

            if (!columnPart || columnPart.length === 0) {
                throw new Error(`Invalid column reference in start cell: ${cleanedStartCell}`);
            }

            if (!rowPart || parseInt(rowPart, 10) <= 0) {
                throw new Error(`Invalid row reference in start cell: ${cleanedStartCell}`);
            }

            logger.info(`Using start cell: ${cleanedStartCell} (column: ${columnPart}, row: ${rowPart})`, { audience: ['developer'] });
        }

        if (cleanedStartNamedRange) {
            logger.info(`Using start named range: ${cleanedStartNamedRange}`, { audience: ['developer'] });
        }

        return {
            spreadsheetId,
            sheetIdentifier,
            startCell: cleanedStartCell,
            startNamedRange: cleanedStartNamedRange,
            tableTitle: config.name,
            tableName: config.table_name,
            transpose: Boolean(config.transpose),
            dataOnly: Boolean(config.data_only)
        };
    }

    private async _uploadResultsToSheet(
        rawResults: any[],
        sqlQuery: string,
        context: PreparedUploadContext
    ): Promise<SheetUploadResult | void> {
        if (rawResults.length === 0) {
            logger.info('Query returned no results. Nothing to upload.', { audience: ['developer'] });
            return;
        }

        const headers = Object.keys(rawResults[0]);
        const data = rawResults.map(row => Object.values(row));
        const dataToUpload = context.dataOnly ? data : [headers, ...data];

        const googleService = getGoogleSheetsService();
        if (!googleService.isConfigured()) {
            throw new Error('Google Sheets API is not configured');
        }

        logger.info('Uploading data to Google Sheets...', { audience: ['developer'] });
        const uploadResult: SheetUploadResult = await googleService.uploadDataToSheet(
            dataToUpload,
            context.spreadsheetId,
            context.sheetIdentifier,
            context.startCell,
            {
                startNamedRange: context.startNamedRange,
                transpose: context.transpose,
                tableTitle: context.tableTitle,
                dataOnly: context.dataOnly,
                sqlQuery,
                tableName: context.tableName
            }
        );
        logger.info('Data uploaded successfully.', { audience: ['developer'] });

        logger.info(`Spreadsheet ID: ${context.spreadsheetId}`);
        const displaySheetName = uploadResult.sheetName
            ?? (typeof context.sheetIdentifier === 'string' ? context.sheetIdentifier : undefined)
            ?? (typeof context.sheetIdentifier === 'number' ? context.sheetIdentifier.toString() : '');
        if (displaySheetName) {
            logger.info(`Sheet: ${displaySheetName}`);
        }
        logger.info(`Range: ${uploadResult.range}`);

        const rangeParts = uploadResult.range.split('!');
        const cellRange = rangeParts.length > 1 ? rangeParts[1] : '';
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${context.spreadsheetId}/edit#gid=${uploadResult.sheetId}&range=${cellRange}`;
        const openSheetButton = 'Open Sheet';
        const successName = displaySheetName || (typeof context.sheetIdentifier === 'number'
            ? `Sheet ID ${context.sheetIdentifier}`
            : 'the target sheet');
        vscode.window.showInformationMessage(`Successfully exported query results to Google Sheet: ${successName}`, openSheetButton)
            .then(selection => {
                if (selection === openSheetButton) {
                    vscode.env.openExternal(vscode.Uri.parse(sheetUrl));
                }
            });

        return uploadResult;
    }
}

/**
 * Get the singleton instance of the GoogleSheetsExportService
 * @returns The singleton instance of the GoogleSheetsExportService
 */
export function getGoogleSheetsExportService(): GoogleSheetsExportService {
    const googleSheetsService = getGoogleSheetsService();
    if (!googleSheetsService.isConfigured()) {
        throw new Error('Google Sheets service is not configured. Please set the service account file in the settings.');
    }
    return new GoogleSheetsExportService();
}

/**
 * Command to export a SQL query to Google Sheets.
 * This function will prompt the user for the necessary information.
 * @param context The extension context
 */
export function registerExportQueryToSheetCommand(context: vscode.ExtensionContext) {
    const command = 'sql-sheets.exportQueryToSheet';

    const commandHandler = async () => {
        // Get the active text editor
        const editor = vscode.window.activeTextEditor;
        if (!editor) {
            vscode.window.showErrorMessage('No active SQL file found.');
            return;
        }

        // Get the SQL query from the editor
        const sqlQuery = editor.document.getText();
        if (!sqlQuery) {
            vscode.window.showErrorMessage('No SQL query found in the active editor.');
            return;
        }

        try {
            // Get the export service
            const exportService = getGoogleSheetsExportService();

            // Prompt for spreadsheet ID
            const spreadsheetId = await vscode.window.showInputBox({
                prompt: 'Enter the Google Sheets spreadsheet ID',
                placeHolder: 'Spreadsheet ID'
            });
            if (!spreadsheetId) {
                return; // User cancelled
            }

            // Prompt for sheet name
            const sheetNameInput = await vscode.window.showInputBox({
                prompt: 'Enter the sheet identifier (ID or "ID | Name")',
                placeHolder: 'Sheet1'
            });
            if (!sheetNameInput) {
                return; // User cancelled
            }

            const parsedSheetIdentifier = SqlSheetConfiguration.parseSheetNameParameter(sheetNameInput);
            const resolvedSheetId = parsedSheetIdentifier.sheetId;
            const resolvedSheetName = parsedSheetIdentifier.sheetName ?? (resolvedSheetId === undefined ? sheetNameInput : undefined);

            // Prompt for start location
            const startLocationInput = await vscode.window.showInputBox({
                prompt: 'Enter the start location (cell or "NamedRange | Cell")',
                placeHolder: 'A1 or MyNamedRange | B2'
            });
            if (startLocationInput === undefined) {
                return; // User cancelled
            }

            const trimmedStartInput = startLocationInput.trim();
            if (trimmedStartInput.length === 0) {
                vscode.window.showErrorMessage('A start location is required.');
                return;
            }

            const parsedStart = SqlSheetConfiguration.parseStartCellParameter(trimmedStartInput);
            const resolvedStartCell = parsedStart.startCell ?? (!parsedStart.startNamedRange ? trimmedStartInput : undefined);
            const resolvedStartNamedRange = parsedStart.startNamedRange;

            if (!resolvedStartCell && !resolvedStartNamedRange) {
                vscode.window.showErrorMessage('Enter a valid start cell (e.g., "A1") or named range (e.g., "MyRange"), or combine them as "MyRange | A1".');
                return;
            }

            // Prompt for table title
            const tableTitle = await vscode.window.showInputBox({
                prompt: 'Enter the table title (optional)',
                placeHolder: 'SQL Query Result'
            });

            // Prompt for transpose
            const transposeSelection = await vscode.window.showQuickPick(['No', 'Yes'], {
                placeHolder: 'Transpose the data?'
            });
            const transpose = transposeSelection === 'Yes';

            // Prompt for data only
            const dataOnlySelection = await vscode.window.showQuickPick(['No', 'Yes'], {
                placeHolder: 'Export data only (no headers)?'
            });
            const dataOnly = dataOnlySelection === 'Yes';

            // Create a configuration object
            const config = new SqlSheetConfiguration(
                spreadsheetId,
                resolvedSheetName,
                resolvedSheetId,
                resolvedStartCell,
                resolvedStartNamedRange,
                tableTitle,
                undefined, // table_name
                undefined, // pre_files
                transpose,
                dataOnly,
                false // skip
            );

            // Export the query to the sheet
            await exportService.exportQueryToSheet(sqlQuery, config);

        } catch (err) {
            const errorMessage = err instanceof Error ? err.message : String(err);
            vscode.window.showErrorMessage(`Failed to export query: ${errorMessage}`);
        }
    };

    context.subscriptions.push(vscode.commands.registerCommand(command, commandHandler));
}
