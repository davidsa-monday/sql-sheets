import * as vscode from 'vscode';
import * as fs from 'fs';
import { google, sheets_v4 } from 'googleapis';
import { JWT } from 'google-auth-library';

// Create a dedicated output channel for SQL Sheets logs
const outputChannel = vscode.window.createOutputChannel('SQL Sheets');

type SheetWithTables = sheets_v4.Schema$Sheet & { tables?: SheetsApiTable[] };

type SheetsApiTable = {
    tableId?: string | null;
    name?: string | null;
    range?: sheets_v4.Schema$GridRange;
    columnProperties?: SheetsApiTableColumn[];
};

type SheetsApiTableColumn = {
    columnIndex?: number | null;
    columnName?: string | null;
    columnType?: string | null;
};

type SheetsTableSnapshot = {
    tableId?: string;
    name?: string;
};

interface SheetsTableRequest extends sheets_v4.Schema$Request {
    addTable?: {
        table: SheetsTableSpec;
    };
    updateTable?: {
        table: SheetsTableSpec & { tableId: string };
        fields: string;
    };
}

type SheetsTableSpec = {
    name?: string;
    range: sheets_v4.Schema$GridRange;
    columnProperties: SheetsTableColumnSpec[];
};

type SheetsTableColumnSpec = {
    columnIndex: number;
    columnName: string;
};

export interface SheetUploadResult {
    range: string;
    sheetId?: number;
}

/**
 * Service for handling Google Sheets integration
 */
export class GoogleSheetsService {
    private serviceAccountCredentials: any = null;

    /**
     * Load Google Service Account credentials from file
     */
    private loadServiceAccountCredentials(): any {
        const config = vscode.workspace.getConfiguration('sql-sheets.google');
        const credentialsFile = config.get<string>('serviceAccountFile');

        if (!credentialsFile) {
            throw new Error('No Google Service Account file configured. Please set a service account file in the settings.');
        }

        if (!fs.existsSync(credentialsFile)) {
            throw new Error(`Google Service Account file does not exist: ${credentialsFile}`);
        }

        try {
            const fileContent = fs.readFileSync(credentialsFile, 'utf-8');
            const credentials = JSON.parse(fileContent);

            if (!credentials.type || credentials.type !== 'service_account') {
                throw new Error('Invalid Google Service Account file. Must be a valid service account JSON file.');
            }

            return credentials;
        } catch (err) {
            if (err instanceof Error) {
                throw new Error(`Failed to load Google Service Account file: ${err.message}`);
            } else {
                throw new Error('Failed to load Google Service Account file');
            }
        }
    }

    /**
     * Initialize Google Sheets service with credentials
     */
    public async initialize(): Promise<boolean> {
        try {
            this.serviceAccountCredentials = this.loadServiceAccountCredentials();
            return true;
        } catch (err) {
            console.error('Failed to initialize Google Sheets service:', err);
            return false;
        }
    }

    /**
     * Check if Google Sheets service is configured properly
     */
    public isConfigured(): boolean {
        try {
            this.loadServiceAccountCredentials();
            return true;
        } catch {
            return false;
        }
    }

    /**
     * Get Service Account Email
     */
    public getServiceAccountEmail(): string | null {
        try {
            const credentials = this.loadServiceAccountCredentials();
            return credentials.client_email || null;
        } catch {
            return null;
        }
    }

    /**
     * Get Project ID
     */
    public getProjectId(): string | null {
        try {
            const credentials = this.loadServiceAccountCredentials();
            return credentials.project_id || null;
        } catch {
            return null;
        }
    }

    /**
     * Upload data to a Google Sheet
     * 
     * @param data Data to upload as a 2D array (rows and columns)
     * @param spreadsheetId ID of the spreadsheet
     * @param sheetName Name or ID of the sheet
     * @param startCell Starting cell (e.g., "A1")
     * @param options Additional options
     */
    public async uploadDataToSheet(
        data: any[][],
        spreadsheetId: string,
        sheetName: string | number,
        startCell: string,
        options: {
            transpose?: boolean,
            tableTitle?: string,
            dataOnly?: boolean,
            sqlQuery?: string,
            queryForNote?: string,
            autoCreateSheet?: boolean,
            tableName?: string
        } = {}
    ): Promise<SheetUploadResult> {
        // Default to auto-creating sheets if not specified
        const autoCreateSheet = options.autoCreateSheet ?? true;
        const tableName = typeof options.tableName === 'string' ? options.tableName.trim() : '';

        try {
            // Show progress indicator
            return await vscode.window.withProgress<SheetUploadResult>({
                location: vscode.ProgressLocation.Notification,
                title: 'Uploading data to Google Sheets',
                cancellable: false
            }, async (progress) => {
                // Ensure we have valid credentials
                const credentials = this.loadServiceAccountCredentials();

                // Parse the start cell to get row and column indexes
                const { column, row } = this.parseCellReference(startCell);

                let targetSheetId: number | undefined;
                let knownTables: SheetsTableSnapshot[] | undefined;

                // Validate sheet name
                if (typeof sheetName === 'string' && sheetName.trim() === '') {
                    throw new Error('Sheet name cannot be empty');
                }

                // Log the parameters for debugging
                console.log(`Uploading to sheet: "${sheetName}", starting at cell: ${column}${row}`);

                // Create a JWT client using the service account credentials
                progress.report({ message: 'Authenticating...' });
                outputChannel.appendLine(`Authenticating with Google service account: ${credentials.client_email}`);

                const client = new JWT({
                    email: credentials.client_email,
                    key: credentials.private_key,
                    scopes: ['https://www.googleapis.com/auth/spreadsheets']
                });

                // Create the Google Sheets API client
                const sheets = google.sheets({ version: 'v4', auth: client });
                outputChannel.appendLine('Google Sheets API client created successfully');

                // Verify the spreadsheet and sheet exist before proceeding
                try {
                    outputChannel.appendLine(`Verifying spreadsheet and sheet "${sheetName}"...`);
                    const spreadsheetInfo = await sheets.spreadsheets.get({
                        spreadsheetId,
                        includeGridData: false,
                        fields: 'sheets.properties,sheets.tables'
                    });

                    const availableSheets = spreadsheetInfo.data.sheets ?? [];
                    const typedSheets = availableSheets as SheetWithTables[];

                    if (typedSheets.length > 0) {
                        outputChannel.appendLine('Available sheets in this spreadsheet:');
                        typedSheets.forEach(sheet => {
                            const sheetId = sheet.properties?.sheetId;
                            const sheetTitle = sheet.properties?.title;
                            outputChannel.appendLine(`- ${sheetTitle} (ID: ${sheetId})`);
                        });
                    }

                    const requestedSheetId = typeof sheetName === 'number' ? sheetName : undefined;
                    const matchedSheet = typedSheets.find(sheet => {
                        const props = sheet.properties;
                        if (!props) {
                            return false;
                        }
                        if (requestedSheetId !== undefined) {
                            return props.sheetId === requestedSheetId;
                        }
                        return props.title === sheetName;
                    });

                    let sheetExists = Boolean(matchedSheet);
                    const matchedSheetId = matchedSheet?.properties?.sheetId;
                    if (sheetExists && typeof matchedSheetId === 'number') {
                        targetSheetId = matchedSheetId;
                        const tables = matchedSheet?.tables ?? [];
                        if (Array.isArray(tables) && tables.length > 0) {
                            knownTables = tables
                                .filter(table => Boolean(table?.name) || Boolean(table?.tableId))
                                .map(table => ({
                                    tableId: typeof table?.tableId === 'string' ? table.tableId : undefined,
                                    name: typeof table?.name === 'string' ? table.name : undefined
                                }));
                        }
                    }

                    if (!sheetExists && typeof sheetName === 'string' && autoCreateSheet) {
                        outputChannel.appendLine(`Sheet "${sheetName}" not found in spreadsheet. Creating it now...`);
                        try {
                            const addSheetResponse = await sheets.spreadsheets.batchUpdate({
                                spreadsheetId,
                                requestBody: {
                                    requests: [
                                        {
                                            addSheet: {
                                                properties: {
                                                    title: sheetName,
                                                    gridProperties: {
                                                        hideGridlines: true
                                                    }
                                                }
                                            }
                                        }
                                    ]
                                }
                            });

                            const newSheetId = addSheetResponse.data.replies?.[0]?.addSheet?.properties?.sheetId;
                            outputChannel.appendLine(`Successfully created sheet "${sheetName}" with ID ${newSheetId}`);
                            sheetExists = true;
                            if (typeof newSheetId === 'number') {
                                targetSheetId = newSheetId;
                            }
                            knownTables = [];
                        } catch (err) {
                            outputChannel.appendLine(`Failed to create sheet "${sheetName}": ${err instanceof Error ? err.message : String(err)}`);
                        }
                    } else if (!sheetExists) {
                        outputChannel.appendLine(`WARNING: Sheet with ID or name "${sheetName}" not found in spreadsheet${autoCreateSheet ? '!' : ' and auto-create disabled.'}`);
                    } else {
                        outputChannel.appendLine(`Sheet "${sheetName}" found in spreadsheet.`);
                    }
                } catch (err) {
                    outputChannel.appendLine(`Error verifying spreadsheet/sheet: ${err instanceof Error ? err.message : String(err)}`);
                }
                outputChannel.show(true);

                const rawTableTitle = typeof options.tableTitle === 'string' ? options.tableTitle : '';
                const titleText = rawTableTitle.trim();
                const hasTitleRow = titleText.length > 0;
                const shouldTranspose = Boolean(options.transpose);
                const executedAt = new Date();
                const executedAtDisplay = executedAt.toLocaleString();
                let titleTimestampColumnIndex: number | undefined;

                // Process the data
                let processedData = [...data];  // Create a copy to avoid modifying the original
                const minimumTitleWidth = 2;

                // Add table title if specified
                if (hasTitleRow) {
                    // Create a title row that spans the width of the data
                    let titleWidth = processedData.length > 0 ? processedData[0].length : 1;

                    if (titleWidth < minimumTitleWidth) {
                        processedData = processedData.map(row => {
                            const rowValues = Array.isArray(row) ? [...row] : [row];
                            while (rowValues.length < minimumTitleWidth) {
                                rowValues.push('');
                            }
                            return rowValues;
                        });
                        titleWidth = minimumTitleWidth;
                    }

                    const titleRow = Array(titleWidth).fill('');
                    titleRow[0] = titleText;
                    titleTimestampColumnIndex = titleWidth - 1;
                    titleRow[titleTimestampColumnIndex] = executedAtDisplay;

                    outputChannel.appendLine(`Adding title row: ${JSON.stringify(titleRow)}`);
                    // Add title row without the empty row gap
                    processedData = [titleRow, ...processedData];
                }

                // Transpose the data if requested
                if (shouldTranspose) {
                    processedData = this.transposeData(processedData);
                    outputChannel.appendLine(`Data transposed: now ${processedData.length} rows x ${processedData[0]?.length || 0} columns`);
                }

                // Calculate the range - handle sheet name differently based on type
                let rangeSheetIdentifier;
                if (typeof sheetName === 'number') {
                    // For numeric sheet IDs, just use the number as is
                    rangeSheetIdentifier = sheetName.toString();
                } else {
                    // For sheet names, use the name
                    rangeSheetIdentifier = sheetName.toString();
                }

                // Get the maximum width of any row in the data for accurate range calculation
                let dataMaxWidth = 0;
                processedData.forEach(row => {
                    dataMaxWidth = Math.max(dataMaxWidth, row ? row.length : 0);
                });

                outputChannel.appendLine(`Data dimensions before range calculation: ${processedData.length} rows x max width ${dataMaxWidth}`);

                // Calculate the end cell coordinates based on data dimensions
                const endRow = row + processedData.length - 1;
                const endColumn = this.incrementColumn(column, dataMaxWidth - 1);

                // Get sheet name prefix for range
                let rangePrefix;
                if (/^\d+$/.test(rangeSheetIdentifier)) {
                    rangePrefix = rangeSheetIdentifier; // For sheet IDs
                } else if (/[^a-zA-Z0-9_]/.test(rangeSheetIdentifier) || /^\d/.test(rangeSheetIdentifier)) {
                    rangePrefix = `'${rangeSheetIdentifier.replace(/'/g, "''")}'`; // For names with special chars
                } else {
                    rangePrefix = rangeSheetIdentifier; // For simple names
                }

                // Build the range manually instead of using calculateRange
                const range = `${rangePrefix}!${column}${row}:${endColumn}${endRow}`;
                outputChannel.appendLine(`Calculated range: ${range}`);

                // Upload the data
                progress.report({ message: 'Uploading data...' });

                // Create the request payload
                const requestPayload = {
                    spreadsheetId,
                    range,
                    valueInputOption: 'USER_ENTERED',
                    requestBody: {
                        values: processedData
                    }
                };

                // Log the request payload for debugging
                outputChannel.appendLine('Google Sheets API Request:');

                // Log data width stats to help diagnose range issues
                const rowWidths = processedData.map(row => row ? row.length : 0);
                const minRowWidth = Math.min(...rowWidths);
                const maxRowWidth = Math.max(...rowWidths);

                outputChannel.appendLine(JSON.stringify({
                    spreadsheetId: requestPayload.spreadsheetId,
                    range: requestPayload.range,
                    valueInputOption: requestPayload.valueInputOption,
                    dataSize: {
                        rows: processedData.length,
                        columns: processedData.length > 0 ? processedData[0].length : 0,
                        minWidth: minRowWidth,
                        maxWidth: maxRowWidth,
                        isJagged: minRowWidth !== maxRowWidth
                    },
                    // Include a more detailed preview
                    preview: processedData.slice(0, 3).map((row, i) => {
                        return {
                            rowIndex: i,
                            length: row ? row.length : 0,
                            content: row ? row.slice(0, Math.min(5, row.length)) : []
                        };
                    })
                }, null, 2));
                outputChannel.show(true);

                // Clear existing content and formatting in the continuous range
                try {
                    outputChannel.appendLine('Starting to clear continuous range...');
                    const resolvedSheetId = typeof targetSheetId === 'number'
                        ? targetSheetId
                        : await this.getSheetId(sheets, spreadsheetId, sheetName.toString());
                    targetSheetId = resolvedSheetId;
                    outputChannel.appendLine(`Got sheet ID: ${resolvedSheetId} for sheet: ${sheetName}`);

                    // Log the parameters being passed to clearContinuousRange
                    outputChannel.appendLine(`Clearing range starting at column: ${column}, row: ${row}, data size: ${processedData.length} x ${processedData[0]?.length || 0}`);

                    await this.clearContinuousRange(
                        sheets,
                        spreadsheetId,
                        resolvedSheetId,
                        column,
                        row,
                        processedData
                    );
                    outputChannel.appendLine('Successfully cleared existing content and formatting in the target range');
                } catch (err) {
                    outputChannel.appendLine(`Warning: Failed to clear existing content and formatting: ${err instanceof Error ? err.message : String(err)}`);
                    if (err instanceof Error && err.stack) {
                        outputChannel.appendLine(`Stack trace: ${err.stack}`);
                    }
                    // Continue with the update even if clearing fails
                }

                // Make the API call to update with new data
                try {
                    outputChannel.appendLine(`Making API call to update range: ${requestPayload.range}`);
                    const response = await sheets.spreadsheets.values.update(requestPayload);
                    outputChannel.appendLine(`API call successful. Updated ${response.data.updatedCells} cells.`);
                } catch (err) {
                    // Get detailed error information
                    outputChannel.appendLine(`Google Sheets API Error: ${err instanceof Error ? err.message : String(err)}`);
                    if (err instanceof Error && 'response' in err) {
                        const apiError = err as any;
                        if (apiError.response?.data?.error) {
                            outputChannel.appendLine(`API Error Details: ${JSON.stringify(apiError.response.data.error, null, 2)}`);

                            // Check for specific error messages
                            const errorMessage = apiError.response?.data?.error?.message;
                            if (errorMessage && errorMessage.includes("Unable to parse range")) {
                                // Special handling for range parsing errors, which are often related to missing sheets
                                throw new Error(
                                    `Unable to write to sheet "${sheetName}". Either the sheet doesn't exist, ` +
                                    `you don't have write permission, or there was a problem with the range format. ` +
                                    `Check the Output panel for details.`
                                );
                            }
                        }
                    }
                    throw err;
                }

                // Apply formatting if not data-only mode
                if (!options.dataOnly) {
                    progress.report({ message: 'Applying formatting...' });
                    await this.applyFormatting(
                        sheets,
                        spreadsheetId,
                        sheetName.toString(),
                        column,
                        row,
                        processedData,
                        targetSheetId,
                        hasTitleRow,
                        {
                            timestampColumnIndex: !shouldTranspose ? titleTimestampColumnIndex : undefined,
                            queryNote: options.queryForNote ?? options.sqlQuery
                        }
                    );
                }

                if (tableName) {
                    try {
                        progress.report({ message: `Configuring table "${tableName}"...` });
                        await this.convertRangeToSheetTable(
                            sheets,
                            spreadsheetId,
                            sheetName.toString(),
                            column,
                            row,
                            processedData,
                            {
                                sheetIdOverride: targetSheetId,
                                hasTitleRow,
                                dataOnly: Boolean(options.dataOnly),
                                transpose: shouldTranspose,
                                tableName,
                                existingTables: knownTables
                            }
                        );
                    } catch (err) {
                        const message = err instanceof Error ? err.message : String(err);
                        outputChannel.appendLine(`Warning: Failed to configure table "${tableName}": ${message}`);
                    }
                }

                outputChannel.appendLine(
                    `Successfully uploaded data to spreadsheet ${spreadsheetId}, sheet ${sheetName}, starting at ${startCell}`
                );
                outputChannel.show(true);

                return {
                    range,
                    sheetId: targetSheetId
                };
            });
        } catch (err) {
            vscode.window.showErrorMessage(
                `Failed to upload data to Google Sheets: ${err instanceof Error ? err.message : String(err)}`
            );
            console.error('Google Sheets upload error:', err);
            throw err;
        }
    }

    /**
     * Parse a cell reference (e.g., "A1") into column and row indexes
     * @param cellRef Cell reference (e.g., "A1")
     * @returns Column letter and row number
     */
    private parseCellReference(cellRef: string): { column: string; row: number } {
        // Ensure we have a valid string
        if (!cellRef || typeof cellRef !== 'string') {
            throw new Error(`Invalid cell reference: ${cellRef}`);
        }

        // Trim any whitespace
        const trimmedRef = cellRef.trim();

        // Make sure the cell reference follows the correct format (one or more letters followed by one or more digits)
        const match = trimmedRef.match(/^([A-Za-z]+)([0-9]+)$/);
        if (!match) {
            throw new Error(`Invalid cell reference format: ${cellRef}. Expected format like 'A1', 'B2', 'AA10', etc.`);
        }

        const column = match[1].toUpperCase();
        const row = parseInt(match[2], 10);

        // Additional validation
        if (column === '') {
            throw new Error(`Invalid column reference in cell: ${cellRef}`);
        }

        if (row <= 0) {
            throw new Error(`Invalid row reference in cell: ${cellRef}. Row must be a positive number.`);
        }

        return { column, row };
    }

    /**
     * Transpose a 2D array
     * @param data Data to transpose
     * @returns Transposed data
     */
    private transposeData(data: any[][]): any[][] {
        if (data.length === 0) {
            return [];
        }

        // Get the maximum row length to handle jagged arrays
        const maxRowLength = Math.max(...data.map(row => row.length));

        // Create a new array with transposed dimensions
        const transposed: any[][] = Array(maxRowLength).fill(null).map(() => Array(data.length).fill(null));

        // Fill the transposed array
        for (let i = 0; i < data.length; i++) {
            for (let j = 0; j < data[i].length; j++) {
                transposed[j][i] = data[i][j];
            }
        }

        return transposed;
    }

    /**
     * Calculate the range for the update operation
     * @param sheetName Sheet name
     * @param column Starting column
     * @param row Starting row
     * @param data Data to upload
     * @returns Range in A1 notation
     */
    private calculateRange(sheetName: string, column: string, row: number, data: any[][]): string {
        try {
            // Validate inputs
            if (!sheetName) {
                throw new Error('Sheet name is required');
            }

            if (!column || !/^[A-Za-z]+$/.test(column)) {
                throw new Error(`Invalid column reference: ${column}`);
            }

            if (typeof row !== 'number' || row <= 0) {
                throw new Error(`Invalid row reference: ${row}`);
            }

            // Properly escape sheet name if it contains special characters
            const escapedSheetName = this.escapeSheetName(sheetName);

            // Handle empty data case
            if (!data) {
                outputChannel.appendLine('Warning: Null data provided to calculateRange');
                return `${escapedSheetName}!${column}${row}`;
            }

            if (data.length === 0) {
                outputChannel.appendLine('Warning: Empty data array provided to calculateRange');
                return `${escapedSheetName}!${column}${row}`;
            }

            if (!data[0]) {
                outputChannel.appendLine('Warning: First row of data is null in calculateRange');
                return `${escapedSheetName}!${column}${row}`;
            }

            if (data[0].length === 0) {
                outputChannel.appendLine('Warning: First row of data is empty in calculateRange');
                return `${escapedSheetName}!${column}${row}`;
            }

            // Calculate end row and column
            const endRow = row + data.length - 1;
            const endColumn = this.incrementColumn(column, data[0].length - 1);

            // Check if sheetName is numeric (sheet ID) and handle differently
            let rangePrefix: string;
            if (/^\d+$/.test(sheetName)) {
                // For numeric sheet IDs, don't use quotes in the range
                rangePrefix = `${sheetName}`;
                outputChannel.appendLine(`Using numeric sheet ID: ${sheetName}`);
            } else {
                // For named sheets, Google requires specific escaping of names
                // Single quotes are required for names with spaces or special characters
                // Names that don't need escaping might actually work better without quotes
                if (/[^a-zA-Z0-9_]/.test(sheetName) || /^\d/.test(sheetName)) {
                    // Names with spaces, special chars, or starting with a digit need quotes
                    rangePrefix = `'${sheetName.replace(/'/g, "''")}'`;
                    outputChannel.appendLine(`Using quoted sheet name (has special chars): '${sheetName}'`);
                } else {
                    // Simple names might work better without quotes
                    rangePrefix = sheetName;
                    outputChannel.appendLine(`Using unquoted sheet name (simple name): ${sheetName}`);
                }
            }

            // Construct the range string in A1 notation
            const range = `${rangePrefix}!${column}${row}:${endColumn}${endRow}`;

            // Log detailed information about the range calculation
            outputChannel.appendLine(`Range Calculation Details:`);
            outputChannel.appendLine(JSON.stringify({
                original: {
                    sheetName,
                    startColumn: column,
                    startRow: row,
                    dataRows: data.length,
                    dataColumns: data[0].length
                },
                calculated: {
                    escapedSheetName,
                    rangePrefix,
                    endColumn,
                    endRow,
                    finalRange: range
                }
            }, null, 2));
            outputChannel.show(true); // Show the output channel

            return range;
        } catch (err) {
            throw new Error(`Unable to parse range: ${err instanceof Error ? err.message : String(err)}`);
        }
    }

    /**
     * Escape a sheet name according to Google Sheets A1 notation rules
     * @param sheetName Sheet name to escape
     * @returns Escaped sheet name
     */
    private escapeSheetName(sheetName: string): string {
        // Handle empty sheet names
        if (!sheetName) {
            throw new Error('Sheet name cannot be empty');
        }

        // If sheetName is a number (as a string), it should be treated as a sheet ID
        if (/^\d+$/.test(sheetName)) {
            return sheetName; // Don't wrap sheet IDs in quotes
        }

        // For sheet names with special characters, spaces, or starting with digits,
        // we need to wrap them in single quotes and escape any single quotes within
        if (/[^a-zA-Z0-9_]/.test(sheetName) || /^\d/.test(sheetName)) {
            return `'${sheetName.replace(/'/g, "''")}'`;
        }

        // For simple alphanumeric names, no quotes are needed
        return sheetName;
    }

    /**
     * Increment a column reference by a specific number of steps
     * @param column Starting column (e.g., "A")
     * @param steps Number of steps to increment
     * @returns New column reference
     */
    private incrementColumn(column: string, steps: number): string {
        if (steps <= 0) {
            return column;
        }

        // Convert the column to a number (A=1, B=2, etc.)
        let value = 0;
        for (let i = 0; i < column.length; i++) {
            value = value * 26 + (column.charCodeAt(i) - 64);
        }

        // Increment by the number of steps
        value += steps;

        // Convert back to a column reference
        let result = '';
        while (value > 0) {
            const remainder = (value - 1) % 26;
            result = String.fromCharCode(65 + remainder) + result;
            value = Math.floor((value - 1) / 26);
        }

        return result;
    }

    private async convertRangeToSheetTable(
        sheets: sheets_v4.Sheets,
        spreadsheetId: string,
        sheetName: string,
        column: string,
        row: number,
        data: any[][],
        options: {
            sheetIdOverride?: number;
            hasTitleRow: boolean;
            dataOnly: boolean;
            transpose: boolean;
            tableName: string;
            existingTables?: SheetsTableSnapshot[];
        }
    ): Promise<void> {
        const trimmedTableName = options.tableName.trim();
        if (!trimmedTableName) {
            return;
        }

        if (options.transpose) {
            outputChannel.appendLine(`convertRangeToSheetTable: Skipping table creation for "${trimmedTableName}" because data is transposed.`);
            return;
        }

        if (options.dataOnly) {
            outputChannel.appendLine(`convertRangeToSheetTable: Skipping table creation for "${trimmedTableName}" because data_only is true.`);
            return;
        }

        const dataStartIndex = options.hasTitleRow ? 1 : 0;
        if (data.length <= dataStartIndex) {
            outputChannel.appendLine(`convertRangeToSheetTable: Unable to locate a header row for table "${trimmedTableName}".`);
            return;
        }

        const tableRows = data.length - dataStartIndex;
        if (tableRows <= 0) {
            outputChannel.appendLine(`convertRangeToSheetTable: No data rows detected for table "${trimmedTableName}".`);
            return;
        }

        const tableWidth = data.slice(dataStartIndex).reduce((max, rowValues) => {
            const width = Array.isArray(rowValues) ? rowValues.length : 0;
            return Math.max(max, width);
        }, 0);

        if (tableWidth === 0) {
            outputChannel.appendLine(`convertRangeToSheetTable: No columns detected for table "${trimmedTableName}".`);
            return;
        }

        const sheetId = options.sheetIdOverride ?? await this.getSheetId(sheets, spreadsheetId, sheetName);

        const startRowIndex = (row - 1) + dataStartIndex;
        const endRowIndex = startRowIndex + tableRows;
        const startColumnIndex = this.columnToIndex(column);
        const endColumnIndex = startColumnIndex + tableWidth;

        const gridRange: sheets_v4.Schema$GridRange = {
            sheetId,
            startRowIndex,
            endRowIndex,
            startColumnIndex,
            endColumnIndex
        };

        const escapedSheetName = this.escapeSheetName(sheetName);
        const headerRowNumber = row + dataStartIndex;
        const lastRowNumber = headerRowNumber + tableRows - 1;
        const endColumnLetter = this.incrementColumn(column, tableWidth - 1);
        const readableRange = `${escapedSheetName}!${column}${headerRowNumber}:${endColumnLetter}${lastRowNumber}`;

        const headerRow = Array.isArray(data[dataStartIndex]) ? data[dataStartIndex] : [];
        const dataRows = data.slice(dataStartIndex + 1);
        const columnProperties = this.buildTableColumnProperties(headerRow, dataRows, tableWidth);

        const tableSpec: SheetsTableSpec = {
            name: trimmedTableName,
            range: gridRange,
            columnProperties
        };

        const existingTable = options.existingTables?.find(table => {
            const existingName = table?.name?.trim();
            return existingName ? existingName.toLowerCase() === trimmedTableName.toLowerCase() : false;
        });

        const requests: SheetsTableRequest[] = [];

        if (existingTable?.tableId) {
            outputChannel.appendLine(`convertRangeToSheetTable: Updating existing table "${trimmedTableName}" (ID: ${existingTable.tableId}) to range ${readableRange}.`);
            requests.push({
                updateTable: {
                    table: {
                        ...tableSpec,
                        tableId: existingTable.tableId
                    },
                    fields: 'name,range,columnProperties'
                }
            });
        } else {
            outputChannel.appendLine(`convertRangeToSheetTable: Adding table "${trimmedTableName}" at range ${readableRange}.`);
            requests.push({
                addTable: {
                    table: tableSpec
                }
            });
        }

        if (requests.length === 0) {
            outputChannel.appendLine(`convertRangeToSheetTable: No table requests prepared for "${trimmedTableName}".`);
            return;
        }

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: requests as sheets_v4.Schema$Request[]
            }
        });

        outputChannel.appendLine(`convertRangeToSheetTable: Table "${trimmedTableName}" configured successfully.`);
    }

    private buildTableColumnProperties(
        headerRow: any[],
        dataRows: any[][],
        columnCount: number
    ): SheetsTableColumnSpec[] {
        const normalizedHeader = Array.isArray(headerRow) ? headerRow : [];
        const usedNames = new Map<string, number>();
        const columns: SheetsTableColumnSpec[] = [];

        for (let index = 0; index < columnCount; index++) {
            const headerValue = normalizedHeader[index];
            const columnName = this.generateUniqueColumnName(headerValue, index, usedNames);
            columns.push({
                columnIndex: index,
                columnName
            });
        }

        return columns;
    }

    private generateUniqueColumnName(
        value: any,
        index: number,
        tracker: Map<string, number>
    ): string {
        const fallbackName = `Column ${index + 1}`;
        let baseName: string;

        if (typeof value === 'string') {
            baseName = value.trim();
        } else if (value !== null && value !== undefined) {
            baseName = String(value).trim();
        } else {
            baseName = '';
        }

        if (baseName.length === 0) {
            baseName = fallbackName;
        }

        const key = baseName.toLowerCase();
        const occurrence = tracker.get(key) ?? 0;
        tracker.set(key, occurrence + 1);

        if (occurrence === 0) {
            return baseName;
        }

        return `${baseName} (${occurrence + 1})`;
    }

    /**
     * Apply formatting to the uploaded data
     * @param sheets Google Sheets API client
     * @param spreadsheetId Spreadsheet ID
     * @param sheetName Sheet name
     * @param column Starting column
     * @param row Starting row
     * @param data Data that was uploaded
     */
    private async applyFormatting(
        sheets: sheets_v4.Sheets,
        spreadsheetId: string,
        sheetName: string,
        column: string,
        row: number,
        data: any[][],
        sheetIdOverride?: number,
        hasTitleRow: boolean = false,
        titleDetails: {
            timestampColumnIndex?: number;
            queryNote?: string | undefined;
        } = {}
    ): Promise<void> {
        if (data.length === 0) {
            outputChannel.appendLine('applyFormatting: no data supplied, skipping.');
            return;
        }

        const totalRows = data.length;
        const dataStartIndex = hasTitleRow ? 1 : 0;
        if (totalRows <= dataStartIndex) {
            outputChannel.appendLine('applyFormatting: unable to locate header row, skipping.');
            return;
        }

        const columnCount = data.reduce((max, rowVals) => {
            const length = rowVals ? rowVals.length : 0;
            return Math.max(max, length);
        }, 0);

        if (columnCount === 0) {
            outputChannel.appendLine('applyFormatting: no columns detected, skipping.');
            return;
        }

        const startColumnIndex = this.columnToIndex(column);
        const endColumnIndex = startColumnIndex + columnCount;
        const sheetId = sheetIdOverride ?? await this.getSheetId(sheets, spreadsheetId, sheetName);
        const firstRowIndex = row - 1;
        const headerRowIndex = firstRowIndex + dataStartIndex;

        const titleRange = hasTitleRow ? {
            sheetId,
            startRowIndex: firstRowIndex,
            endRowIndex: firstRowIndex + 1,
            startColumnIndex: startColumnIndex,
            endColumnIndex
        } : undefined;

        const headerRange = {
            sheetId,
            startRowIndex: headerRowIndex,
            endRowIndex: headerRowIndex + 1,
            startColumnIndex: startColumnIndex,
            endColumnIndex
        };

        const borderRange = {
            sheetId,
            startRowIndex: firstRowIndex,
            endRowIndex: firstRowIndex + totalRows,
            startColumnIndex: startColumnIndex,
            endColumnIndex
        };

        const requests = [];

        if (titleRange) {
            requests.push({
                repeatCell: {
                    range: titleRange,
                    cell: {
                        userEnteredFormat: {
                            backgroundColor: { red: 0, green: 0, blue: 0 },
                            textFormat: {
                                foregroundColor: { red: 1, green: 1, blue: 1 },
                                bold: true
                            },
                            horizontalAlignment: 'LEFT'
                        }
                    },
                    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            });
        }

        requests.push({
            repeatCell: {
                range: headerRange,
                cell: {
                    userEnteredFormat: {
                        textFormat: {
                            bold: true
                        },
                        horizontalAlignment: 'LEFT'
                    }
                },
                fields: 'userEnteredFormat(textFormat,horizontalAlignment)'
            }
        });

        const borderColor = { red: 0, green: 0, blue: 0 };
        requests.push({
            updateBorders: {
                range: borderRange,
                top: { style: 'SOLID', width: 1, color: borderColor },
                bottom: { style: 'SOLID', width: 1, color: borderColor },
                left: { style: 'SOLID', width: 1, color: borderColor },
                right: { style: 'SOLID', width: 1, color: borderColor }
            }
        });

        requests.push({
            autoResizeDimensions: {
                dimensions: {
                    sheetId,
                    dimension: 'COLUMNS',
                    startIndex: startColumnIndex,
                    endIndex: endColumnIndex
                }
            }
        });

        requests.push({
            updateSheetProperties: {
                properties: {
                    sheetId,
                    gridProperties: {
                        hideGridlines: true
                    }
                },
                fields: 'gridProperties.hideGridlines'
            }
        });

        const timestampRelativeIndex = typeof titleDetails.timestampColumnIndex === 'number'
            ? titleDetails.timestampColumnIndex
            : undefined;

        let timestampRange;
        if (titleRange && typeof timestampRelativeIndex === 'number') {
            const timestampColumnIndex = startColumnIndex + timestampRelativeIndex;
            timestampRange = {
                sheetId,
                startRowIndex: firstRowIndex,
                endRowIndex: firstRowIndex + 1,
                startColumnIndex: timestampColumnIndex,
                endColumnIndex: timestampColumnIndex + 1
            };

            requests.push({
                repeatCell: {
                    range: timestampRange,
                    cell: {
                        userEnteredFormat: {
                            backgroundColor: { red: 0, green: 0, blue: 0 },
                            textFormat: {
                                fontSize: 8,
                                foregroundColor: { red: 0.6, green: 0.6, blue: 0.6 },
                                bold: false
                            },
                            horizontalAlignment: 'RIGHT'
                        }
                    },
                    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            });
        }

        const queryNote = titleDetails.queryNote?.trim();
        if (titleRange && queryNote && timestampRange) {
            requests.push({
                repeatCell: {
                    range: timestampRange,
                    cell: {
                        note: queryNote
                    },
                    fields: 'note'
                }
            });
        }

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests
            }
        });
    }

    /**
     * Convert a column letter to a 0-based index
     * @param column Column letter (e.g., "A")
     * @returns 0-based index
     */
    private columnToIndex(column: string): number {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + (column.charCodeAt(i) - 64);
        }
        return index - 1;
    }

    /**
     * Convert a 0-based index to a column letter
     * @param index 0-based index
     * @returns Column letter (e.g., "A", "B", "AA")
     */
    private indexToColumn(index: number): string {
        index = index + 1; // Convert from 0-based to 1-based for the conversion
        let columnName = '';
        while (index > 0) {
            const remainder = (index - 1) % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            index = Math.floor((index - 1) / 26);
        }
        return columnName;
    }

    /**
     * Creates a format clearing request for a specific range in a sheet
     * @param sheetId ID of the sheet
     * @param startRow Start row (1-indexed)
     * @param startCol Start column (1-indexed)
     * @param endRow End row (1-indexed, inclusive)
     * @param endCol End column (1-indexed, inclusive)
     * @returns Format clearing request object for batch update
     */
    private createClearFormatsRequest(sheetId: number, startRow: number, startCol: number, endRow: number, endCol: number): any {
        outputChannel.appendLine(`Creating clear formats request for sheetId=${sheetId}, startRow=${startRow}, startCol=${startCol}, endRow=${endRow}, endCol=${endCol}`);

        return {
            repeatCell: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: startRow - 1,  // Convert to 0-indexed
                    endRowIndex: endRow,  // endRow is inclusive
                    startColumnIndex: startCol - 1,  // Convert to 0-indexed
                    endColumnIndex: endCol  // endCol is inclusive
                },
                cell: {
                    userEnteredFormat: {
                        textFormat: {
                            bold: false,
                            foregroundColor: {},  // Reset to default color
                        },
                        backgroundColor: {
                            red: 1,
                            green: 1,
                            blue: 1,
                        },  // White background
                        borders: {
                            top: { style: "NONE" },
                            bottom: { style: "NONE" },
                            left: { style: "NONE" },
                            right: { style: "NONE" },
                        }
                    }
                },
                fields: "userEnteredFormat(textFormat,backgroundColor,borders)"
            }
        };
    }

    /**
     * Clear both content and formatting in the continuous range of a sheet
     * @param sheets Google Sheets API client
     * @param spreadsheetId Spreadsheet ID
     * @param sheetId Sheet ID
     * @param startColumn Starting column letter (e.g., "A")
     * @param startRow Starting row number
     * @param newData New data to be uploaded (for calculating dimensions)
     */
    private async clearContinuousRange(
        sheets: sheets_v4.Sheets,
        spreadsheetId: string,
        sheetId: number,
        startColumn: string,
        startRow: number,
        newData: any[][]
    ): Promise<void> {
        outputChannel.appendLine(`clearContinuousRange called with: sheetId=${sheetId}, startColumn=${startColumn}, startRow=${startRow}`);

        // Get the dimensions of the new data
        const numDataRows = newData.length;
        const numCols = newData.length > 0 ? newData[0].length : 0;
        outputChannel.appendLine(`New data dimensions: ${numDataRows} rows x ${numCols} columns`);

        try {
            // Get the sheet name for the range
            const sheetMetadata = await sheets.spreadsheets.get({
                spreadsheetId,
                ranges: [],
                includeGridData: false
            });

            const sheet = sheetMetadata.data.sheets?.find(s => s.properties?.sheetId === sheetId);
            const sheetTitle = sheet?.properties?.title;

            if (!sheetTitle) {
                throw new Error(`Could not find sheet with ID ${sheetId}`);
            }

            outputChannel.appendLine(`Working with sheet: "${sheetTitle}" (ID: ${sheetId})`);

            // Properly format the sheet name for A1 notation
            let rangePrefix: string;
            if (/[^a-zA-Z0-9_]/.test(sheetTitle) || /^\d/.test(sheetTitle)) {
                rangePrefix = `'${sheetTitle.replace(/'/g, "''")}'`;
            } else {
                rangePrefix = sheetTitle;
            }

            // Get all values from the worksheet to determine the continuous range
            const rangeToFetch = `${rangePrefix}!${startColumn}${startRow}:ZZ10000`;
            outputChannel.appendLine(`Fetching sheet values from ${rangeToFetch}`);

            const response = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: rangeToFetch
            });

            // Extract the values from the response
            const allValues = response.data.values || [];
            const maxRow = allValues.length;
            const maxCol = Math.max(...allValues.map(row => row.length), 0);
            outputChannel.appendLine(`Fetched ${maxRow} rows and max width of ${maxCol} columns from the sheet`);

            // Convert start column letter to index (0-based)
            const startColIndex = this.columnToIndex(startColumn);
            outputChannel.appendLine(`Start column "${startColumn}" converted to 0-based index: ${startColIndex}`);

            // If the sheet is empty or we're starting beyond existing data
            if (maxRow === 0) {
                outputChannel.appendLine("Sheet appears to be empty, no need to clear anything");
                return;
            }

            // Find the last non-empty row
            let lastDataRow = 0;
            for (let r = 0; r < allValues.length; r++) {
                const row = allValues[r];
                // Check if the row has any non-empty cells
                if (row && row.some(cell => cell !== undefined && cell !== null && cell !== "")) {
                    lastDataRow = r;
                }
            }
            lastDataRow++; // Convert to 1-based row number

            // Find the last non-empty column
            let lastDataCol = 0;
            for (let r = 0; r < allValues.length; r++) {
                const row = allValues[r] || [];
                for (let c = 0; c < row.length; c++) {
                    if (row[c] !== undefined && row[c] !== null && row[c] !== "") {
                        lastDataCol = Math.max(lastDataCol, c);
                    }
                }
            }
            lastDataCol++; // Convert to 1-based column index

            // Calculate the full range to clear
            const clearEndRow = Math.max(startRow + lastDataRow, startRow + numDataRows);
            const clearEndColIndex = Math.max(startColIndex + lastDataCol, startColIndex + numCols);
            const clearEndColLetter = this.indexToColumn(clearEndColIndex);

            outputChannel.appendLine(`Determined data ranges: last data row=${lastDataRow}, last data col=${lastDataCol}`);
            outputChannel.appendLine(`Clearing range from ${startColumn}${startRow} to ${clearEndColLetter}${clearEndRow}`);

            // Create the A1 range notation
            const clearRange = `${rangePrefix}!${startColumn}${startRow}:${clearEndColLetter}${clearEndRow}`;

            // First, clear the values
            outputChannel.appendLine(`Clearing values in range: ${clearRange}`);
            await sheets.spreadsheets.values.clear({
                spreadsheetId,
                range: clearRange
            });

            // Then, clear the formatting
            outputChannel.appendLine(`Clearing formatting in range: ${clearRange}`);
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                requestBody: {
                    requests: [
                        this.createClearFormatsRequest(
                            sheetId,
                            startRow,
                            startColIndex + 1, // Convert to 1-based for the function
                            clearEndRow,
                            clearEndColIndex + 1 // Convert to 1-based for the function
                        )
                    ]
                }
            });

            outputChannel.appendLine(`Successfully cleared both values and formatting in range: ${clearRange}`);

        } catch (err) {
            outputChannel.appendLine(`Error clearing continuous range: ${err instanceof Error ? err.message : String(err)}`);
            if (err instanceof Error && err.stack) {
                outputChannel.appendLine(`Stack trace: ${err.stack}`);
            }

            // Fallback to clearing just the area of the new data
            try {
                outputChannel.appendLine(`Attempting fallback clearing method...`);
                // Calculate the end position based on the new data dimensions
                const endRow = startRow + numDataRows - 1;
                const endColIndex = this.columnToIndex(startColumn) + numCols;
                const endColumnLetter = this.indexToColumn(endColIndex);

                // Define the fallback range in A1 notation
                const fallbackRange = `${startColumn}${startRow}:${endColumnLetter}${endRow}`;
                outputChannel.appendLine(`Fallback range: ${fallbackRange}`);

                // Clear the values
                await sheets.spreadsheets.values.clear({
                    spreadsheetId,
                    range: fallbackRange
                });

                // Clear the formatting
                await sheets.spreadsheets.batchUpdate({
                    spreadsheetId,
                    requestBody: {
                        requests: [
                            this.createClearFormatsRequest(
                                sheetId,
                                startRow,
                                this.columnToIndex(startColumn) + 1,
                                endRow,
                                endColIndex + 1
                            )
                        ]
                    }
                });

                outputChannel.appendLine(`Successfully cleared fallback range: ${fallbackRange}`);
            } catch (fallbackErr) {
                outputChannel.appendLine(`Failed to clear fallback range: ${fallbackErr instanceof Error ? fallbackErr.message : String(fallbackErr)}`);
                if (fallbackErr instanceof Error && fallbackErr.stack) {
                    outputChannel.appendLine(`Stack trace: ${fallbackErr.stack}`);
                }
                throw fallbackErr;
            }
        }
    }    /**
     * Get the sheet ID by name
     * @param sheets Google Sheets API client
     * @param spreadsheetId Spreadsheet ID
     * @param sheetName Sheet name
     * @returns Sheet ID
     */
    private async getSheetId(sheets: sheets_v4.Sheets, spreadsheetId: string, sheetName: string): Promise<number> {
        const response = await sheets.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });

        const trimmedSheetName = sheetName.trim();
        const numericCandidate = Number(trimmedSheetName);
        const isNumericIdentifier = trimmedSheetName !== '' && !Number.isNaN(numericCandidate);

        const sheet = response.data.sheets?.find(s => {
            const properties = s.properties;
            if (!properties) {
                return false;
            }

            if (isNumericIdentifier && properties.sheetId === numericCandidate) {
                return true;
            }

            return properties.title === sheetName;
        });

        if (!sheet?.properties?.sheetId) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }

        return sheet.properties.sheetId;
    }
}

// Singleton instance
let googleSheetsService: GoogleSheetsService | undefined;

/**
 * Get the Google Sheets service instance
 */
export function getGoogleSheetsService(): GoogleSheetsService {
    if (!googleSheetsService) {
        googleSheetsService = new GoogleSheetsService();
    }
    return googleSheetsService;
}
