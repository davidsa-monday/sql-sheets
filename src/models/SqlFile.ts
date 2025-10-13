import * as vscode from 'vscode';
import { SqlQuery } from './SqlQuery';
import { SqlSheetConfiguration } from './SqlSheetConfiguration';

export class SqlFile {
    private _queries: SqlQuery[] = [];

    constructor(public readonly document: vscode.TextDocument) {
        this.parse();
    }

    public get queries(): SqlQuery[] {
        return this._queries;
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
        const paramRegex = new RegExp(`--${key}:\\s*.*`, 'g');
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
    }

    private parse(): void {
        this._queries = [];
        const text = this.document.getText();
        const queryBlocks = text.split(';').filter(b => b.trim().length > 0);

        let currentOffset = 0;
        for (const block of queryBlocks) {
            const blockWithSemicolon = block + ';';
            const startOffset = text.indexOf(block, currentOffset);
            const endOffset = startOffset + blockWithSemicolon.length;

            const regex = /--(\w+):\s*(.*)/g;
            let match;
            const params: { [key: string]: string } = {};
            let queryText = block;

            while ((match = regex.exec(block)) !== null) {
                params[match[1]] = match[2].trim();
            }

            const firstSelect = block.toLowerCase().indexOf('select');
            if (firstSelect !== -1) {
                queryText = block.substring(firstSelect);
            }

            // Process boolean parameters using centralized validation in SqlSheetConfiguration
            // If the parameter doesn't exist in the SQL comment, it will be undefined
            // and the constructor will use the default value (false)
            const config = new SqlSheetConfiguration(
                params['spreadsheet_id'],
                params['sheet_name'],
                params['start_cell'],
                params['start_named_range'],
                params['name'],
                params['table_name'],
                params['pre_file'],
                SqlSheetConfiguration.stringToBoolean(params['transpose']),
                SqlSheetConfiguration.stringToBoolean(params['data_only']),
                SqlSheetConfiguration.stringToBoolean(params['skip'])
            );

            this._queries.push(new SqlQuery(config, queryText, startOffset, endOffset));
            currentOffset = endOffset;
        }
    }
}
