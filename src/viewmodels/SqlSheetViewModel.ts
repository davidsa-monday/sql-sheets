import * as vscode from 'vscode';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';
import { SqlFile } from '../models/SqlFile';
import { SqlQuery } from '../models/SqlQuery';

export class SqlSheetViewModel {
    // Get all parameter keys from SqlSheetConfiguration using reflection
    private static _parameterKeys: string[] = Object.getOwnPropertyNames(
        new SqlSheetConfiguration())
        .filter(key => key !== 'constructor' && key !== 'start_named_range' && key !== 'sheet_id')
        .map(key => key === 'pre_files' ? 'pre_file' : key);

    // Expose the keys as a property
    public get parameterKeys(): string[] {
        return SqlSheetViewModel._parameterKeys;
    }

    // Expose parameter descriptions
    public get parameterDescriptions(): Record<string, string> {
        return SqlSheetConfiguration.parameterDescriptions;
    }

    // Get parameter types
    public get parameterTypes(): Record<string, string> {
        const config = new SqlSheetConfiguration();
        const types: Record<string, string> = {};

        // Determine the type of each property
        for (const key of SqlSheetViewModel._parameterKeys) {
            // Use centralized parameter type checking
            if (SqlSheetConfiguration.isBooleanParameter(key)) {
                types[key] = 'boolean';
            } else {
                types[key] = 'string';
            }
        }

        return types;
    }

    private _sqlFile?: SqlFile;
    private _activeQuery?: SqlQuery;
    private _onDidChange: vscode.EventEmitter<SqlSheetViewModel> = new vscode.EventEmitter<SqlSheetViewModel>();
    readonly onDidChange: vscode.Event<SqlSheetViewModel> = this._onDidChange.event;

    constructor() {
        vscode.window.onDidChangeActiveTextEditor(() => this.onActiveEditorChanged());
        vscode.workspace.onDidChangeTextDocument(e => this.onDocumentChanged(e));
        vscode.window.onDidChangeTextEditorSelection(e => this.onSelectionChanged(e));
        this.onActiveEditorChanged();
    }

    public get config(): Record<string, unknown> {
        const config = this._activeQuery?.config ?? new SqlSheetConfiguration();
        return {
            ...config,
            pre_file: config.pre_files[0] ?? ''
        };
    }

    public async updateParameter(key: string, value: string): Promise<void> {
        if (this._sqlFile && this._activeQuery) {
            // The SqlFile class will handle the type conversion appropriately
            await this._sqlFile.updateParameter(this._activeQuery, key, value);
        }
    }

    private onActiveEditorChanged(): void {
        const editor = vscode.window.activeTextEditor;
        if (editor && editor.document.fileName.endsWith('.sql')) {
            this._sqlFile = new SqlFile(editor.document);
            this.updateActiveQuery();
        } else {
            this._sqlFile = undefined;
            this._activeQuery = undefined;
        }
        this._onDidChange.fire(this);
    }

    private onDocumentChanged(e: vscode.TextDocumentChangeEvent): void {
        if (this._sqlFile && e.document === this._sqlFile.document) {
            this._sqlFile = new SqlFile(e.document);
            this.updateActiveQuery();
            this._onDidChange.fire(this);
        }
    }

    private onSelectionChanged(e: vscode.TextEditorSelectionChangeEvent): void {
        if (this._sqlFile && e.textEditor.document === this._sqlFile.document) {
            this.updateActiveQuery();
            this._onDidChange.fire(this);
        }
    }

    private updateActiveQuery(): void {
        const editor = vscode.window.activeTextEditor;
        if (this._sqlFile && editor) {
            this._activeQuery = this._sqlFile.getQueryAt(editor.selection.active);
        }
    }
}
