import * as vscode from 'vscode';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';
import { getGoogleSheetsService } from './googleSheetsService';
import type { SheetUploadResult } from './googleSheetsService';
import { getSnowflakeService } from './snowflakeService';
import { getLogger } from './loggingService';

const logger = getLogger();

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
        const {
            spreadsheet_id: spreadsheetId,
            sheet_name: sheetName,
            sheet_id: sheetId,
            start_cell: startCell,
            start_named_range: startNamedRange,
            name: tableTitle,
            transpose,
            data_only: dataOnly,
            skip
        } = config;

        const cleanedStartNamedRange = typeof startNamedRange === 'string' ? startNamedRange.trim() : undefined;
        const cleanedStartCell = typeof startCell === 'string' ? startCell.trim() : undefined;

        if (skip) {
            logger.info('Skipping query execution because configuration is marked skip', { audience: ['developer'] });
            return;
        }

        // Validate inputs
        if (!sqlQuery) {
            throw new Error('SQL query is required');
        }
        if (!spreadsheetId) {
            throw new Error('Spreadsheet ID is required');
        }
        const sheetIdentifier = sheetId !== undefined ? sheetId : sheetName;
        if (sheetIdentifier === undefined || (typeof sheetIdentifier === 'string' && sheetIdentifier.trim() === '')) {
            throw new Error('Sheet name or ID is required');
        }
        if (typeof sheetIdentifier === 'string') {
            logger.info(`Using sheet name: "${sheetIdentifier}"`, { audience: ['developer'] });
            logger.revealOutput();
        } else {
            logger.info(`Using sheet ID: ${sheetIdentifier}`, { audience: ['developer'] });
            logger.revealOutput();
        }

        if (!cleanedStartCell && !cleanedStartNamedRange) {
            throw new Error('A start cell or named range is required');
        }

        if (cleanedStartCell) {
            // Validate start cell format
            try {
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
                logger.revealOutput();
            } catch (err) {
                throw new Error(`Invalid start cell format: ${err instanceof Error ? err.message : String(err)}`);
            }
        }

        if (cleanedStartNamedRange) {
            logger.info(`Using start named range: ${cleanedStartNamedRange}`, { audience: ['developer'] });
            logger.revealOutput();
        }

        try {
            // Get the services
            const snowflakeService = getSnowflakeService();
            const googleService = getGoogleSheetsService();

            // Check if the services are configured
            if (!snowflakeService.isConfigured()) {
                throw new Error('Snowflake connection is not configured');
            }

            if (!googleService.isConfigured()) {
                throw new Error('Google Sheets API is not configured');
            }

            // Execute the query using the Snowflake service
            logger.info('Executing query in Snowflake...', { audience: ['developer'] });
            logger.revealOutput();
            const results = await snowflakeService.executeQuery(sqlQuery);
            logger.info('Query executed successfully.', { audience: ['developer'] });

            // Check if there are results to upload
            if (results.length === 0) {
                logger.info('Query returned no results. Nothing to upload.', { audience: ['developer'] });
                return;
            }

            // Prepare the data for Google Sheets
            const headers = Object.keys(results[0]);
            const data = results.map(row => Object.values(row));
            const dataToUpload = dataOnly ? data : [headers, ...data];

            // Upload the data to Google Sheets
            logger.info('Uploading data to Google Sheets...', { audience: ['developer'] });
            const uploadResult: SheetUploadResult = await googleService.uploadDataToSheet(
                dataToUpload,
                spreadsheetId,
                sheetIdentifier,
                cleanedStartCell,
                {
                    startNamedRange: cleanedStartNamedRange,
                    transpose: transpose,
                    tableTitle: tableTitle,
                    dataOnly: dataOnly,
                    sqlQuery: sqlQuery,
                    tableName: config.table_name
                }
            );
            logger.info('Data uploaded successfully.', { audience: ['developer'] });

            // Log the result details
            logger.info(`Spreadsheet ID: ${spreadsheetId}`);
            const displaySheetName = uploadResult.sheetName
                ?? (typeof sheetIdentifier === 'string' ? sheetIdentifier : undefined)
                ?? (typeof sheetIdentifier === 'number' ? sheetIdentifier.toString() : '');
            if (displaySheetName) {
                logger.info(`Sheet: ${displaySheetName}`);
            }
            logger.info(`Range: ${uploadResult.range}`);

            // Show a success message to the user with a link to the sheet
            const rangeParts = uploadResult.range.split('!');
            const cellRange = rangeParts.length > 1 ? rangeParts[1] : '';
            const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${uploadResult.sheetId}&range=${cellRange}`;
            const openSheetButton = 'Open Sheet';
            const successName = displaySheetName || (typeof sheetIdentifier === 'number' ? `Sheet ID ${sheetIdentifier}` : 'the target sheet');
            vscode.window.showInformationMessage(`Successfully exported query results to Google Sheet: ${successName}`, openSheetButton)
                .then(selection => {
                    if (selection === openSheetButton) {
                        vscode.env.openExternal(vscode.Uri.parse(sheetUrl));
                    }
                });

            return uploadResult;
        } catch (err) {
            const errorMessage = err instanceof Error ? err.message : String(err);
            logger.error('Failed to export query to Google Sheet', { data: err });
            vscode.window.showErrorMessage(`Failed to export query to Google Sheet: ${errorMessage}`);
        }
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
                undefined, // name_t
                undefined, // pre_file
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
