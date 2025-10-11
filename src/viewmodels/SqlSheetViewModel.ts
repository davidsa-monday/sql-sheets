import * as vscode from 'vscode';
import { SqlSheetConfiguration } from '../models/SqlSheetConfiguration';
import { SqlFile } from '../models/SqlFile';
import { SqlQuery } from '../models/SqlQuery';

export class SqlSheetViewModel {
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

    public get config(): SqlSheetConfiguration {
        return this._activeQuery?.config ?? new SqlSheetConfiguration();
    }

    public async updateParameter(key: string, value: string): Promise<void> {
        if (this._sqlFile && this._activeQuery) {
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
