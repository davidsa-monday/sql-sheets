import * as vscode from 'vscode';
import { SqlQuery } from './SqlQuery';
import { SqlSheetConfiguration } from './SqlSheetConfiguration';

const NON_INHERITABLE_DEFAULTS = new Set(['start_cell', 'name', 'table_name']);

export class SqlFile {
    private _queries: SqlQuery[] = [];
    private defaultParameterRanges: Record<string, { start: number; end: number }> = {};
    private topLevelPreFiles: string[] = [];

    constructor(public readonly document: vscode.TextDocument) {
        this.parse();
    }

    public get queries(): SqlQuery[] {
        return this._queries;
    }

    public getGlobalPreFiles(): string[] {
        return [...this.topLevelPreFiles];
    }

    public getQueryAt(position: vscode.Position): SqlQuery | undefined {
        const offset = this.document.offsetAt(position);
        return this._queries.find(q => offset >= q.startOffset && offset <= q.endOffset);
    }

    public async updateParameter(query: SqlQuery, key: string, value: string): Promise<void> {
        // Get the text of the query block
        const queryText = this.document.getText(new vscode.Range(
            this.document.positionAt(query.startOffset),
            this.document.positionAt(query.endOffset)
        ));

        const edit = new vscode.WorkspaceEdit();

        // Format value if it's for a boolean parameter
        if (SqlSheetConfiguration.isBooleanParameter(key)) {
            // Use centralized validation in the SqlSheetConfiguration class
            value = SqlSheetConfiguration.formatBooleanToString(value);
        }

        // Check if the parameter already exists
        const paramRegex = new RegExp(`--${key}:[\\t ]*[^\\r\\n]*`, 'g');
        const match = paramRegex.exec(queryText);

        if (match) {
            // Parameter exists, update its line
            const matchStart = query.startOffset + match.index;
            const matchEnd = matchStart + match[0].length;
            const range = new vscode.Range(
                this.document.positionAt(matchStart),
                this.document.positionAt(matchEnd)
            );
            edit.replace(this.document.uri, range, `--${key}: ${value}`);
        } else {
            const source = query.parameterSources?.[key];
            const defaultRangeOffsets = source === 'default' ? this.defaultParameterRanges[key] : undefined;

            if (defaultRangeOffsets) {
                const range = new vscode.Range(
                    this.document.positionAt(defaultRangeOffsets.start),
                    this.document.positionAt(defaultRangeOffsets.end)
                );
                const replacementText = `--${key}: ${value}`;
                edit.replace(this.document.uri, range, replacementText);
                await vscode.workspace.applyEdit(edit);
                this.parse();
                return;
            }

            // Parameter doesn't exist, add it to the beginning of the query block
            // Find where to insert the new parameter (before the first non-comment line)
            const lines = queryText.split('\n');
            let insertPos = query.startOffset;
            let insertLine = 0;

            // Find the first non-comment line
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line.startsWith('--') && line !== '') {
                    break;
                }
                insertLine = i + 1;
                insertPos += lines[i].length + 1; // +1 for newline
            }

            const position = this.document.positionAt(insertPos);
            edit.insert(this.document.uri, position, `--${key}: ${value}\n`);
        }

        await vscode.workspace.applyEdit(edit);
        this.parse();
    }

    private parse(): void {
        this._queries = [];
        const text = this.document.getText();
        const defaultParameters = this.extractDefaultParameters(text);
        const defaultParams = defaultParameters.values;
        this.defaultParameterRanges = defaultParameters.ranges;
        this.topLevelPreFiles = defaultParameters.preFiles;
        const queryBlocks = text.split(';').filter(b => b.trim().length > 0);

        let currentOffset = 0;
        for (const block of queryBlocks) {
            const blockWithSemicolon = block + ';';
            const startOffset = text.indexOf(block, currentOffset);
            const endOffset = startOffset + blockWithSemicolon.length;

            const regex = /--(\w+):[\t ]*(.*)/g;
            let match;
            const params: { [key: string]: string } = { ...defaultParams };
            const parameterSources: Record<string, 'default' | 'query'> = {};

            for (const key of Object.keys(defaultParams)) {
                parameterSources[key] = 'default';
            }
            const preFiles: string[] = [];
            let queryText = block;

            while ((match = regex.exec(block)) !== null) {
                const key = match[1];
                const value = match[2].trim();

                if (key === 'pre_file') {
                    preFiles.push(value);
                    continue;
                }

                params[key] = value;
                parameterSources[key] = 'query';
            }

            if (queryText.length > 0) {
                const lines = queryText.split('\n');
                let firstStatementLine = 0;

                while (firstStatementLine < lines.length) {
                    const trimmedLine = lines[firstStatementLine].trim();

                    if (trimmedLine.length === 0) {
                        firstStatementLine++;
                        continue;
                    }

                    if (/^--\w+:/.test(trimmedLine)) {
                        firstStatementLine++;
                        continue;
                    }

                    break;
                }

                queryText = lines.slice(firstStatementLine).join('\n');
            }

            const combinedSheetParameter = SqlSheetConfiguration.parseSheetNameParameter(params['sheet_name']);
            const sheetId = combinedSheetParameter.sheetId;
            const sheetName = combinedSheetParameter.sheetName ?? (sheetId === undefined ? params['sheet_name'] : undefined);

            const combinedStartParameter = SqlSheetConfiguration.parseStartCellParameter(params['start_cell']);
            const startNamedRange = combinedStartParameter.startNamedRange ?? params['start_named_range'];
            const startCell = combinedStartParameter.startCell ?? (!startNamedRange ? params['start_cell'] : undefined);

            // Process boolean parameters using centralized validation in SqlSheetConfiguration
            // If the parameter doesn't exist in the SQL comment, it will be undefined
            // and the constructor will use the default value (false)
            const config = new SqlSheetConfiguration(
                params['spreadsheet_id'],
                sheetName,
                sheetId,
                startCell,
                startNamedRange,
                params['name'],
                params['table_name'],
                preFiles,
                SqlSheetConfiguration.stringToBoolean(params['transpose']),
                SqlSheetConfiguration.stringToBoolean(params['data_only']),
                SqlSheetConfiguration.stringToBoolean(params['skip'])
            );

            this._queries.push(new SqlQuery(config, queryText, startOffset, endOffset, this.document.uri, parameterSources));
            currentOffset = endOffset;
        }
    }

    private extractDefaultParameters(text: string): {
        values: Record<string, string>;
        ranges: Record<string, { start: number; end: number }>;
        preFiles: string[];
    } {
        const defaults: Record<string, string> = {};
        const ranges: Record<string, { start: number; end: number }> = {};
        const preFiles: string[] = [];
        const lineRegex = /^.*(?:\r?\n|$)/gm;
        let match: RegExpExecArray | null;

        while ((match = lineRegex.exec(text)) !== null) {
            const lineWithEol = match[0];
            const lineStart = match.index;
            const lineContent = lineWithEol.replace(/\r?\n$/, '');
            const trimmedLine = lineContent.trim();

            if (trimmedLine.length === 0) {
                continue;
            }

            const parameterMatch = /^--(\w+):[\t ]*(.*)$/.exec(trimmedLine);
            if (parameterMatch) {
                const key = parameterMatch[1];
                const value = parameterMatch[2].trim();
                if (key === 'pre_file') {
                    preFiles.push(value);
                    continue;
                }

                if (NON_INHERITABLE_DEFAULTS.has(key)) {
                    continue;
                }

                defaults[key] = value;
                ranges[key] = {
                    start: lineStart,
                    end: lineStart + lineContent.length
                };
                continue;
            }

            if (trimmedLine.startsWith('--')) {
                continue;
            }

            break;
        }

        return { values: defaults, ranges, preFiles };
    }
}
