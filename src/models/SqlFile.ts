import * as vscode from 'vscode';
import { SqlQuery, QueryDestination } from './SqlQuery';
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

            const defaultSourceMap: Record<string, 'default' | 'query'> = {};
            for (const key of Object.keys(defaultParams)) {
                defaultSourceMap[key] = 'default';
            }

            const preFiles: string[] = [];
            const headerLinesMeta: Array<{ lineContent: string; trimmed: string; start: number; end: number }> = [];
            const lineRegex = /^.*(?:\r?\n|$)/gm;
            let headerEndRelative = block.length;
            let lineMatch: RegExpExecArray | null;

            lineRegex.lastIndex = 0;
            while ((lineMatch = lineRegex.exec(block)) !== null) {
                const lineWithEol = lineMatch[0];
                const lineStart = lineMatch.index;
                const lineEnd = lineStart + lineWithEol.length;
                const lineContent = lineWithEol.replace(/\r?\n$/, '');
                const trimmed = lineContent.trim();

                if (trimmed.length === 0 || trimmed.startsWith('--')) {
                    headerLinesMeta.push({ lineContent, trimmed, start: lineStart, end: lineEnd });
                    continue;
                }

                headerEndRelative = lineStart;
                break;
            }

            const destinationParamsList: Record<string, string>[] = [];
            const destinationSourcesList: Record<string, 'default' | 'query'>[] = [];
            const destinationParameterRanges: Array<Map<string, { start: number; end: number }>> = [];

            if (headerLinesMeta.length === 0) {
                destinationParamsList.push({ ...defaultParams });
                destinationSourcesList.push({ ...defaultSourceMap });
                destinationParameterRanges.push(new Map());
            } else {
                const separatorRegex = /^--\s*\d+\s*$/;
                let currentParams: Record<string, string> = { ...defaultParams };
                let currentSources: Record<string, 'default' | 'query'> = { ...defaultSourceMap };
                let currentRanges: Map<string, { start: number; end: number }> = new Map();

                for (const lineMeta of headerLinesMeta) {
                    const trimmed = lineMeta.trimmed;

                    if (separatorRegex.test(trimmed)) {
                        destinationParamsList.push(currentParams);
                        destinationSourcesList.push(currentSources);
                        destinationParameterRanges.push(currentRanges);

                        currentParams = { ...defaultParams };
                        currentSources = { ...defaultSourceMap };
                        currentRanges = new Map();
                        continue;
                    }

                    const parameterMatch = /^--(\w+):[\t ]*(.*)$/.exec(lineMeta.lineContent);
                    if (parameterMatch) {
                        const key = parameterMatch[1];
                        const value = parameterMatch[2].trim();

                        if (key === 'pre_file') {
                            preFiles.push(value);
                            continue;
                        }

                        currentParams[key] = value;
                        currentSources[key] = 'query';

                        const paramStartInLine = lineMeta.lineContent.indexOf(parameterMatch[0]);
                        const absoluteStart = startOffset + lineMeta.start + Math.max(paramStartInLine, 0);
                        const absoluteEnd = absoluteStart + parameterMatch[0].length;
                        currentRanges.set(key, { start: absoluteStart, end: absoluteEnd });
                        continue;
                    }
                }

                destinationParamsList.push(currentParams);
                destinationSourcesList.push(currentSources);
                destinationParameterRanges.push(currentRanges);
            }

            let queryText = block.substring(headerEndRelative);
            if (queryText.length > 0) {
                queryText = queryText.replace(/^[\r\n]+/, '');
            }

            const queryDestinations: QueryDestination[] = destinationParamsList.map((paramValues, index) => {
                const combinedSheetParameter = SqlSheetConfiguration.parseSheetNameParameter(paramValues['sheet_name']);
                const sheetId = combinedSheetParameter.sheetId;
                const sheetName = combinedSheetParameter.sheetName ?? (sheetId === undefined ? paramValues['sheet_name'] : undefined);

                const combinedStartParameter = SqlSheetConfiguration.parseStartCellParameter(paramValues['start_cell']);
                const startNamedRange = combinedStartParameter.startNamedRange ?? paramValues['start_named_range'];
                const startCell = combinedStartParameter.startCell ?? (!startNamedRange ? paramValues['start_cell'] : undefined);

                return {
                    config: new SqlSheetConfiguration(
                        paramValues['spreadsheet_id'],
                        sheetName,
                        sheetId,
                        startCell,
                        startNamedRange,
                        paramValues['name'],
                        paramValues['table_name'],
                        preFiles,
                        SqlSheetConfiguration.stringToBoolean(paramValues['transpose']),
                        SqlSheetConfiguration.stringToBoolean(paramValues['data_only']),
                        SqlSheetConfiguration.stringToBoolean(paramValues['skip'])
                    ),
                    parameterRanges: destinationParameterRanges[index] ?? new Map<string, { start: number; end: number }>()
                };
            });

            const parameterSources = { ...(destinationSourcesList[0] ?? {}) };

            this._queries.push(new SqlQuery(queryDestinations, queryText, startOffset, endOffset, this.document.uri, parameterSources));
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
