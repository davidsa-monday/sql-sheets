import * as vscode from 'vscode';
import * as fs from 'fs';
import { getGoogleSheetsService } from './googleSheetsService';
import type { SheetUploadResult } from './googleSheetsService';
import { getSnowflakeService } from './snowflakeService';

// Create a dedicated output channel for SQL Sheets logs
const outputChannel = vscode.window.createOutputChannel('SQL Sheets');

/**
 * Service for exporting SQL query results to Google Sheets
 */
export class GoogleSheetsExportService {
    /**
     * Export SQL query results to a Google Sheet
     * @param sqlQuery The SQL query to execute
     * @param spreadsheetId The Google Sheets spreadsheet ID
     * @param sheetName The name or ID of the sheet
     * @param startCell The cell where the data should start (e.g. "A1")
     * @param options Additional export options
     * @returns A promise that resolves when the export is complete
     */
    public async exportQueryToSheet(
        sqlQuery: string,
        spreadsheetId: string,
        sheetName: string | number,
        startCell: string,
        options: {
            transpose?: boolean,
            tableTitle?: string,
            dataOnly?: boolean
        } = {}
    ): Promise<void> {
        // Validate inputs
        if (!sqlQuery) {
            throw new Error('SQL query is required');
        }
        if (!spreadsheetId) {
            throw new Error('Spreadsheet ID is required');
        }
        if (!sheetName) {
            throw new Error('Sheet name or ID is required');
        }
        if (typeof sheetName === 'string') {
            if (sheetName.trim() === '') {
                throw new Error('Sheet name cannot be empty');
            }
            // Log the sheet name for debugging
            outputChannel.appendLine(`Using sheet name: "${sheetName}"`);
            outputChannel.show(true);
        } else if (typeof sheetName === 'number') {
            // Log the sheet ID for debugging
            outputChannel.appendLine(`Using sheet ID: ${sheetName}`);
            outputChannel.show(true);
        }

        if (!startCell) {
            throw new Error('Start cell is required');
        }

        // Validate start cell format
        try {
            // More comprehensive validation of cell format (e.g. A1, B12, AC345)
            const validCellFormat = /^[A-Za-z]+[0-9]+$/;
            if (!validCellFormat.test(startCell)) {
                throw new Error(`Start cell must be in the format like 'A1', 'B2', etc.`);
            }

            // Extract and validate column and row parts
            const columnPart = startCell.match(/^[A-Za-z]+/)?.[0] || '';
            const rowPart = startCell.match(/[0-9]+$/)?.[0] || '';

            if (!columnPart || columnPart.length === 0) {
                throw new Error(`Invalid column reference in start cell: ${startCell}`);
            }

            if (!rowPart || parseInt(rowPart, 10) <= 0) {
                throw new Error(`Invalid row reference in start cell: ${startCell}`);
            }

            // Log the validated cell for debugging
            outputChannel.appendLine(`Using start cell: ${startCell} (column: ${columnPart}, row: ${rowPart})`);
            outputChannel.show(true);
        } catch (err) {
            throw new Error(`Invalid start cell format: ${err instanceof Error ? err.message : String(err)}`);
        }

        const { transpose = false, tableTitle = 'SQL Query Result', dataOnly = false } = options;

        try {
            // Get the services
            const snowflakeService = getSnowflakeService();
            const googleService = getGoogleSheetsService();

            // Check if the services are configured
            if (!snowflakeService.isConfigured()) {
                throw new Error('Snowflake connection is not configured');
            }

            if (!googleService.isConfigured()) {
                throw new Error('Google Sheets service is not configured. Please configure a Google Service Account JSON file in the settings.');
            }

            // Show progress notification
            const uploadSummary = await vscode.window.withProgress<SheetUploadResult>({
                location: vscode.ProgressLocation.Notification,
                title: 'Exporting SQL query to Google Sheets',
                cancellable: false
            }, async (progress) => {
                // Execute the SQL query
                progress.report({ message: 'Executing SQL query...' });
                outputChannel.appendLine('Executing SQL query...');
                const results = await snowflakeService.executeQuery(sqlQuery);

                // Log result info
                outputChannel.appendLine(`Query returned ${results.length} row(s)`);
                if (results.length > 0) {
                    outputChannel.appendLine(`First row sample: ${JSON.stringify(results[0], null, 2)}`);
                }

                // Convert results to an appropriate format for Google Sheets
                progress.report({ message: 'Processing results...' });
                outputChannel.appendLine('Processing results for Google Sheets format...');
                const processedResults = this._processQueryResults(results);

                // Log processed data info
                outputChannel.appendLine(`Processed data has ${processedResults.length} row(s)`);
                if (processedResults.length > 0) {
                    outputChannel.appendLine(`Headers: ${JSON.stringify(processedResults[0])}`);
                    outputChannel.appendLine(`Data dimensions: ${processedResults.length} rows x ${processedResults[0].length} columns`);

                    // Debug the full structure to see what's happening with the data
                    outputChannel.appendLine(`First two rows of processed data (for debugging):`);
                    for (let i = 0; i < Math.min(2, processedResults.length); i++) {
                        outputChannel.appendLine(`Row ${i}: ${JSON.stringify(processedResults[i])}`);
                    }
                }
                outputChannel.show(true);

                // Export to Google Sheets
                progress.report({ message: 'Uploading to Google Sheets...' });

                // Process sheet name - if it's a string that looks like a number, convert it to actual number
                let processedSheetName = sheetName;
                if (typeof sheetName === 'string' && /^\d+$/.test(sheetName)) {
                    // Convert to number if it's a numeric string
                    processedSheetName = parseInt(sheetName, 10);
                    outputChannel.appendLine(`Converting sheet name "${sheetName}" to numeric ID: ${processedSheetName}`);
                    outputChannel.show(true);
                }

                const uploadResult = await googleService.uploadDataToSheet(
                    processedResults,
                    spreadsheetId,
                    processedSheetName,
                    startCell,
                    {
                        transpose,
                        tableTitle,
                        dataOnly,
                        sqlQuery,
                        autoCreateSheet: true // Explicitly enable sheet auto-creation
                    }
                );
                progress.report({ message: 'Export completed successfully' });
                return uploadResult;
            });

            // Show success message
            const sheetNameDisplay = typeof sheetName === 'number' ? `Sheet ID ${sheetName}` : sheetName;
            vscode.window.showInformationMessage(
                `Successfully exported SQL results to "${sheetNameDisplay}" at ${startCell}`
            );

            const fallbackSheetId = typeof sheetName === 'number'
                ? sheetName
                : (/^\d+$/.test(sheetName) ? parseInt(sheetName, 10) : undefined);
            const sheetIdForUrl = uploadSummary?.sheetId ?? fallbackSheetId;
            const rangeForUrl = this.extractRangeForUrl(uploadSummary?.range);
            const spreadsheetUrl = this.buildSpreadsheetUrl(spreadsheetId, sheetIdForUrl, rangeForUrl);

            // Open the spreadsheet directly at the exported range
            vscode.env.openExternal(vscode.Uri.parse(spreadsheetUrl));


        } catch (err) {
            vscode.window.showErrorMessage(`Failed to export SQL query to Google Sheets: ${err instanceof Error ? err.message : String(err)}`);
            throw err;
        }
    }

    private extractRangeForUrl(range: string | undefined): string | undefined {
        if (!range) {
            return undefined;
        }

        const separatorIndex = range.indexOf('!');
        const rangePortion = separatorIndex >= 0 ? range.slice(separatorIndex + 1) : range;
        const trimmed = rangePortion.trim();

        return trimmed.length > 0 ? trimmed : undefined;
    }

    private buildSpreadsheetUrl(
        spreadsheetId: string,
        sheetId: number | undefined,
        range: string | undefined
    ): string {
        const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
        const hashParts: string[] = [];

        if (typeof sheetId === 'number' && !Number.isNaN(sheetId)) {
            hashParts.push(`gid=${sheetId}`);
        }

        if (range) {
            hashParts.push(`range=${encodeURIComponent(range)}`);
        }

        return hashParts.length > 0 ? `${baseUrl}#${hashParts.join('&')}` : baseUrl;
    }

    /**
     * Export SQL file to a Google Sheet
     * @param filePath The SQL file path
     * @param spreadsheetId The Google Sheets spreadsheet ID
     * @param sheetName The name or ID of the sheet
     * @param startCell The cell where the data should start (e.g. "A1")
     * @param options Additional export options
     * @returns A promise that resolves when the export is complete
     */
    public async exportSqlFileToSheet(
        filePath: string,
        spreadsheetId: string,
        sheetName: string | number,
        startCell: string,
        options: {
            transpose?: boolean,
            tableTitle?: string,
            dataOnly?: boolean
        } = {}
    ): Promise<void> {
        try {
            // Read the SQL file
            const fileContent = fs.readFileSync(filePath, 'utf-8');

            // Set the table title to the file name if not provided
            const fileName = filePath.split('/').pop() || filePath;
            const tableTitle = options.tableTitle || fileName;

            // Export the query
            await this.exportQueryToSheet(
                fileContent,
                spreadsheetId,
                sheetName,
                startCell,
                {
                    ...options,
                    tableTitle
                }
            );
        } catch (err) {
            vscode.window.showErrorMessage(`Failed to export SQL file to Google Sheets: ${err instanceof Error ? err.message : String(err)}`);
            throw err;
        }
    }

    /**
     * Process query results into a format suitable for Google Sheets
     * @param results The SQL query results
     * @returns Processed results ready for Google Sheets
     */
    private _processQueryResults(results: any[]): any[][] {
        if (!results || results.length === 0) {
            return [['No results found']];
        }

        // Extract column headers from the first result object
        const headers = Object.keys(results[0]);
        outputChannel.appendLine(`Processing results: Found ${headers.length} columns`);
        outputChannel.appendLine(`Column headers: ${JSON.stringify(headers)}`);

        try {
            // Convert each result object to an array of values
            const rows = results.map((row, rowIndex) => {
                return headers.map(header => {
                    const value = row[header];

                    // Handle null values
                    if (value === null) {
                        return '';
                    }

                    // Handle dates
                    if (value instanceof Date) {
                        return value.toISOString();
                    }

                    return value;
                });
            });

            // Create the final result with headers and rows
            const processedData = [headers, ...rows];

            // Log the dimensions for debugging
            outputChannel.appendLine(`Processed data dimensions: ${processedData.length} rows x ${processedData[0]?.length || 0} columns`);
            return processedData;
        } catch (err) {
            outputChannel.appendLine(`Error processing query results: ${err instanceof Error ? err.message : String(err)}`);
            // Return a safe fallback
            return [['Error processing results'], [`${err instanceof Error ? err.message : String(err)}`]];
        }
    }
}

// Singleton instance
let googleSheetsExportService: GoogleSheetsExportService | undefined;

/**
 * Get the Google Sheets Export service instance
 */
export function getGoogleSheetsExportService(): GoogleSheetsExportService {
    if (!googleSheetsExportService) {
        googleSheetsExportService = new GoogleSheetsExportService();
    }
    return googleSheetsExportService;
}
