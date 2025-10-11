import * as vscode from 'vscode';
import { SqlSheetViewModel } from '../viewmodels/SqlSheetViewModel';

export class SqlSheetEditorProvider implements vscode.WebviewViewProvider {

    public static readonly viewType = 'sql-sheets.editor';

    private _view?: vscode.WebviewView;

    constructor(
        private readonly _extensionUri: vscode.Uri,
        private readonly _viewModel: SqlSheetViewModel
    ) {
        this._viewModel.onDidChange((vm: SqlSheetViewModel) => {
            if (this._view) {
                this._view.webview.postMessage({
                    type: 'update',
                    config: vm.config
                });
            }
        });
    }

    public resolveWebviewView(
        webviewView: vscode.WebviewView,
        context: vscode.WebviewViewResolveContext,
        _token: vscode.CancellationToken,
    ) {
        this._view = webviewView;

        webviewView.webview.options = {
            enableScripts: true,
            localResourceRoots: [
                this._extensionUri
            ]
        };

        webviewView.webview.html = this._getHtmlForWebview(webviewView.webview);

        webviewView.webview.onDidReceiveMessage(data => {
            switch (data.type) {
                case 'edit':
                    {
                        this._viewModel.updateParameter(data.key, data.value);
                        break;
                    }
            }
        });

        // Post the initial data
        setTimeout(() => {
            webviewView.webview.postMessage({
                type: 'update',
                config: this._viewModel.config
            });
        }, 200);
    }

    private _getHtmlForWebview(webview: vscode.Webview) {
        const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'media', 'main.js'));
        const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'media', 'main.css'));
        const toolkitUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'node_modules', '@vscode', 'webview-ui-toolkit', 'dist', 'toolkit.js'));

        return `<!DOCTYPE html>
			<html lang="en">
			<head>
				<meta charset="UTF-8">
				<meta name="viewport" content="width=device-width, initial-scale=1.0">
				<script type="module" src="${toolkitUri}"></script>
				<link href="${styleUri}" rel="stylesheet">
				<title>SQL Sheet Editor</title>
			</head>
			<body>
				<vscode-data-grid id="config-grid" grid-template-columns="150px 1fr" generate-header="sticky" aria-label="SQL Sheet Configuration"></vscode-data-grid>
				<script src="${scriptUri}"></script>
			</body>
			</html>`;
    }
}
