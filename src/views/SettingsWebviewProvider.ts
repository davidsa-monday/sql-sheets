import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { getSnowflakeService } from '../services/snowflakeService';

export class SettingsWebviewProvider {
    public static readonly viewType = 'sql-sheets.settings';

    private _panel: vscode.WebviewPanel | undefined;
    private _extensionUri: vscode.Uri;

    constructor(extensionUri: vscode.Uri) {
        this._extensionUri = extensionUri;
    }

    public static createOrShow(extensionUri: vscode.Uri) {
        const provider = new SettingsWebviewProvider(extensionUri);
        provider.show();
    }

    public async show() {
        const columnToShowIn = vscode.window.activeTextEditor
            ? vscode.window.activeTextEditor.viewColumn
            : undefined;

        if (this._panel) {
            this._panel.reveal(columnToShowIn);
            return;
        }

        // Create and show the webview
        this._panel = vscode.window.createWebviewPanel(
            SettingsWebviewProvider.viewType,
            'SQL Sheets Settings',
            columnToShowIn || vscode.ViewColumn.One,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [
                    vscode.Uri.joinPath(this._extensionUri, 'media')
                ]
            }
        );

        // Set the webview's initial html content
        this._panel.webview.html = await this._getHtmlForWebview();

        // Handle messages from the webview
        this._panel.webview.onDidReceiveMessage(
            async (message) => {
                switch (message.command) {
                    case 'loadCredentialsFile':
                        await this._loadCredentialsFile();
                        break;
                    case 'testConnection':
                        await this._testConnection();
                        break;
                    case 'createTemplate':
                        await this._createTemplate();
                        break;
                    case 'loadGoogleCredentialsFile':
                        await this._loadGoogleCredentialsFile();
                        break;
                    case 'createGoogleTemplate':
                        await this._createGoogleTemplate();
                        break;
                    case 'refresh':
                        this._panel!.webview.html = await this._getHtmlForWebview();
                        break;
                }
            },
            undefined,
            []
        );

        // When the panel is disposed, clean up resources
        this._panel.onDidDispose(
            () => {
                this._panel = undefined;
            },
            null,
            []
        );
    }

    private async _getHtmlForWebview() {
        // Get the current Snowflake settings
        const snowflakeConfig = vscode.workspace.getConfiguration('sql-sheets.connection');
        const credentialsFile = snowflakeConfig.get<string>('credentialsFile') || '';

        // Get the current Google settings
        const googleConfig = vscode.workspace.getConfiguration('sql-sheets.google');
        const googleCredentialsFile = googleConfig.get<string>('serviceAccountFile') || '';

        // Get Snowflake credentials file content if it exists
        let credentialsSummary = '';
        if (credentialsFile && fs.existsSync(credentialsFile)) {
            try {
                const fileContent = fs.readFileSync(credentialsFile, 'utf-8');
                const credentials = JSON.parse(fileContent);

                credentialsSummary = `
          <div class="credentials-summary">
            <h3>Snowflake Credentials Summary:</h3>
            <p><strong>User:</strong> ${this._escapeHtml(credentials.user || 'Not set')}</p>
            <p><strong>Account:</strong> ${this._escapeHtml(credentials.account || 'Not set')}</p>
            <p><strong>Warehouse:</strong> ${this._escapeHtml(credentials.warehouse || 'Not set')}</p>
            <p><strong>Database:</strong> ${this._escapeHtml(credentials.database || 'Not set')}</p>
            <p><strong>Schema:</strong> ${this._escapeHtml(credentials.schema || 'Not set')}</p>
            <p class="note">To change credentials, edit the JSON file directly or select a new credentials file.</p>
          </div>
        `;
            } catch (err) {
                credentialsSummary = `
          <div class="error-message">
            <p>Error reading Snowflake credentials file: ${this._escapeHtml(String(err))}</p>
          </div>
        `;
            }
        } else {
            credentialsSummary = `
        <div class="warning-message">
          <p>No Snowflake credentials file selected or file does not exist.</p>
          <p>Please select a credentials file using the Browse button or create a new template.</p>
        </div>
      `;
        }

        // Get Google Service Account credentials file content if it exists
        let googleCredentialsSummary = '';
        if (googleCredentialsFile && fs.existsSync(googleCredentialsFile)) {
            try {
                const fileContent = fs.readFileSync(googleCredentialsFile, 'utf-8');
                const credentials = JSON.parse(fileContent);

                googleCredentialsSummary = `
          <div class="credentials-summary google-credentials">
            <h3>Google Service Account Summary:</h3>
            <p><strong>Type:</strong> ${this._escapeHtml(credentials.type || 'Not set')}</p>
            <p><strong>Project ID:</strong> ${this._escapeHtml(credentials.project_id || 'Not set')}</p>
            <p><strong>Client Email:</strong> ${this._escapeHtml(credentials.client_email || 'Not set')}</p>
            <p><strong>Private Key ID:</strong> ${this._escapeHtml(credentials.private_key_id ? '********' : 'Not set')}</p>
            <p class="note">To change credentials, edit the JSON file directly or select a new service account file.</p>
          </div>
        `;
            } catch (err) {
                googleCredentialsSummary = `
          <div class="error-message">
            <p>Error reading Google Service Account credentials file: ${this._escapeHtml(String(err))}</p>
          </div>
        `;
            }
        } else {
            googleCredentialsSummary = `
        <div class="warning-message">
          <p>No Google Service Account credentials file selected or file does not exist.</p>
          <p>Please select a service account file using the Browse button or create a new template.</p>
        </div>
      `;
        }

        // Create the HTML content
        return `<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>SQL Sheets Settings</title>
        <style>
            body {
                font-family: var(--vscode-font-family);
                padding: 20px;
                color: var(--vscode-foreground);
                background-color: var(--vscode-editor-background);
            }
            .container {
                max-width: 600px;
                margin: 0 auto;
            }
            h1 {
                border-bottom: 1px solid var(--vscode-panel-border);
                padding-bottom: 10px;
                margin-bottom: 20px;
            }
            .section-title {
                border-bottom: 1px solid var(--vscode-panel-border);
                padding-bottom: 5px;
                margin-top: 30px;
                margin-bottom: 15px;
            }
            .form-group {
                margin-bottom: 15px;
            }
            label {
                display: block;
                margin-bottom: 5px;
                font-weight: bold;
            }
            input[type="text"] {
                width: 100%;
                padding: 8px;
                border: 1px solid var(--vscode-input-border);
                background-color: var(--vscode-input-background);
                color: var(--vscode-input-foreground);
                border-radius: 2px;
            }
            .buttons {
                display: flex;
                justify-content: space-between;
                margin-top: 20px;
            }
            button {
                padding: 8px 16px;
                background-color: var(--vscode-button-background);
                color: var(--vscode-button-foreground);
                border: none;
                border-radius: 2px;
                cursor: pointer;
                margin-right: 10px;
            }
            button:last-child {
                margin-right: 0;
            }
            button:hover {
                background-color: var(--vscode-button-hoverBackground);
            }
            .credentials-file {
                display: flex;
                align-items: center;
                justify-content: space-between;
            }
            .credentials-file input {
                flex-grow: 1;
                margin-right: 10px;
            }
            .message {
                margin-top: 20px;
                padding: 10px;
                border-radius: 2px;
            }
            .success {
                background-color: rgba(0, 128, 0, 0.1);
                color: var(--vscode-notificationsSuccessForeground);
            }
            .error {
                background-color: rgba(255, 0, 0, 0.1);
                color: var(--vscode-notificationsErrorForeground);
            }
            .warning {
                background-color: rgba(255, 255, 0, 0.1);
                color: var(--vscode-notificationsWarningForeground);
            }
            .warning-message {
                background-color: rgba(255, 255, 0, 0.1);
                color: var(--vscode-notificationsWarningForeground);
                padding: 10px;
                border-radius: 2px;
                margin-top: 20px;
            }
            .error-message {
                background-color: rgba(255, 0, 0, 0.1);
                color: var(--vscode-notificationsErrorForeground);
                padding: 10px;
                border-radius: 2px;
                margin-top: 20px;
            }
            .credentials-summary {
                background-color: var(--vscode-editor-inactiveSelectionBackground);
                padding: 15px;
                border-radius: 4px;
                margin-top: 20px;
            }
            .credentials-summary h3 {
                margin-top: 0;
            }
            .note {
                font-style: italic;
                margin-top: 15px;
                color: var(--vscode-descriptionForeground);
            }
            .google-credentials {
                background-color: var(--vscode-editor-inactiveSelectionBackground);
                border-left: 4px solid #4285F4;
            }
            .tab-container {
                margin-bottom: 20px;
            }
            .tab-buttons {
                display: flex;
                border-bottom: 1px solid var(--vscode-panel-border);
            }
            .tab-button {
                padding: 10px 20px;
                background: none;
                border: none;
                cursor: pointer;
                margin-right: 0;
                color: var(--vscode-foreground);
                opacity: 0.7;
                border-bottom: 2px solid transparent;
            }
            .tab-button.active {
                opacity: 1;
                font-weight: bold;
                border-bottom: 2px solid var(--vscode-button-background);
            }
            .tab-content {
                padding-top: 20px;
            }
            .tab-pane {
                display: none;
            }
            .tab-pane.active {
                display: block;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>SQL Sheets Settings</h1>
            <div id="message" style="display: none;" class="message"></div>
            
            <div class="tab-container">
                <div class="tab-buttons">
                    <button type="button" class="tab-button active" data-tab="snowflake">Snowflake</button>
                    <button type="button" class="tab-button" data-tab="google">Google Service Account</button>
                </div>
                
                <div class="tab-content">
                    <div id="snowflake-tab" class="tab-pane active">
                        <form id="snowflakeForm">
                            <h2>Snowflake Credentials</h2>
                            <p>Configure your Snowflake connection by selecting a credentials JSON file.</p>
                            <div class="form-group credentials-file">
                                <input type="text" id="credentialsFile" placeholder="Path to credentials file" value="${this._escapeHtml(credentialsFile)}" readonly>
                                <button type="button" id="browseButton">Browse...</button>
                            </div>
                            
                            <div class="buttons">
                                <button type="button" id="createTemplateButton">Create Template</button>
                                <button type="button" id="testButton">Test Connection</button>
                            </div>
                        </form>
                        ${credentialsSummary}
                    </div>
                    
                    <div id="google-tab" class="tab-pane">
                        <form id="googleForm">
                            <h2>Google Service Account</h2>
                            <p>Configure your Google Service Account by selecting a service account JSON file.</p>
                            <div class="form-group credentials-file">
                                <input type="text" id="googleCredentialsFile" placeholder="Path to service account file" value="${this._escapeHtml(googleCredentialsFile)}" readonly>
                                <button type="button" id="browseGoogleButton">Browse...</button>
                            </div>
                            
                            <div class="buttons">
                                <button type="button" id="createGoogleTemplateButton">Create Template</button>
                            </div>
                        </form>
                        ${googleCredentialsSummary}
                    </div>
                </div>
            </div>
        </div>

        <script>
            const vscode = acquireVsCodeApi();
            
            // Tab functionality
            document.querySelectorAll('.tab-button').forEach(button => {
                button.addEventListener('click', () => {
                    // Remove active class from all buttons and panes
                    document.querySelectorAll('.tab-button').forEach(btn => {
                        btn.classList.remove('active');
                    });
                    document.querySelectorAll('.tab-pane').forEach(pane => {
                        pane.classList.remove('active');
                    });
                    
                    // Add active class to clicked button and its corresponding pane
                    button.classList.add('active');
                    const tabId = button.getAttribute('data-tab');
                    document.getElementById(tabId + '-tab').classList.add('active');
                });
            });
            
            // Snowflake credentials actions
            document.getElementById('browseButton').addEventListener('click', () => {
                vscode.postMessage({
                    command: 'loadCredentialsFile'
                });
            });
            
            document.getElementById('testButton').addEventListener('click', () => {
                vscode.postMessage({
                    command: 'testConnection'
                });
            });

            document.getElementById('createTemplateButton').addEventListener('click', () => {
                vscode.postMessage({
                    command: 'createTemplate'
                });
            });
            
            // Google credentials actions
            document.getElementById('browseGoogleButton').addEventListener('click', () => {
                vscode.postMessage({
                    command: 'loadGoogleCredentialsFile'
                });
            });
            
            document.getElementById('createGoogleTemplateButton').addEventListener('click', () => {
                vscode.postMessage({
                    command: 'createGoogleTemplate'
                });
            });

            // Show messages
            function showMessage(text, type) {
                const messageDiv = document.getElementById('message');
                messageDiv.textContent = text;
                messageDiv.className = 'message ' + type;
                messageDiv.style.display = 'block';
                
                setTimeout(() => {
                    messageDiv.style.display = 'none';
                }, 5000);
            }

            // Listen for messages from the extension
            window.addEventListener('message', event => {
                const message = event.data;
                switch (message.command) {
                    case 'connectionSuccess':
                        showMessage('Connection successful!', 'success');
                        break;
                    case 'connectionError':
                        showMessage('Connection failed: ' + message.error, 'error');
                        break;
                    case 'credentialsLoaded':
                        document.getElementById('credentialsFile').value = message.filePath;
                        showMessage('Snowflake credentials loaded from file!', 'success');
                        // Reload page to show updated credentials summary
                        setTimeout(() => {
                            vscode.postMessage({ command: 'refresh' });
                        }, 500);
                        break;
                    case 'templateCreated':
                        showMessage('Template file created at: ' + message.filePath, 'success');
                        break;
                    case 'googleCredentialsLoaded':
                        document.getElementById('googleCredentialsFile').value = message.filePath;
                        showMessage('Google Service Account credentials loaded from file!', 'success');
                        // Reload page to show updated credentials summary
                        setTimeout(() => {
                            vscode.postMessage({ command: 'refresh' });
                        }, 500);
                        break;
                    case 'googleTemplateCreated':
                        showMessage('Google Service Account template created at: ' + message.filePath, 'success');
                        break;
                }
            });
        </script>
    </body>
    </html>`;
    }

    private _escapeHtml(unsafe: string): string {
        return unsafe
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    private async _saveSettings(settings: any) {
        try {
            const config = vscode.workspace.getConfiguration('sql-sheets');
            await config.update('connection.user', settings.user, vscode.ConfigurationTarget.Global);
            await config.update('connection.password', settings.password, vscode.ConfigurationTarget.Global);
            await config.update('connection.account', settings.account, vscode.ConfigurationTarget.Global);
            await config.update('connection.warehouse', settings.warehouse, vscode.ConfigurationTarget.Global);
            await config.update('connection.database', settings.database, vscode.ConfigurationTarget.Global);
            await config.update('connection.schema', settings.schema, vscode.ConfigurationTarget.Global);

            this._panel?.webview.postMessage({ command: 'settingsSaved' });
        } catch (err) {
            vscode.window.showErrorMessage(`Failed to save settings: ${err instanceof Error ? err.message : String(err)}`);
        }
    }

    private async _loadCredentialsFile() {
        try {
            // Show file picker
            const options: vscode.OpenDialogOptions = {
                canSelectMany: false,
                openLabel: 'Select Credentials File',
                filters: {
                    'JSON Files': ['json']
                }
            };

            const fileUri = await vscode.window.showOpenDialog(options);

            if (!fileUri || fileUri.length === 0) {
                return;
            }

            const filePath = fileUri[0].fsPath;

            // Read and parse the file to validate it's proper JSON
            const fileContent = fs.readFileSync(filePath, 'utf-8');
            const credentials = JSON.parse(fileContent);

            // Validate the structure
            if (!credentials.user || !credentials.account) {
                throw new Error('Invalid credentials file. Must contain at least "user" and "account" fields.');
            }

            // Save the credentials file path only
            const config = vscode.workspace.getConfiguration('sql-sheets');
            await config.update('connection.credentialsFile', filePath, vscode.ConfigurationTarget.Global);

            // Update the UI
            this._panel?.webview.postMessage({
                command: 'credentialsLoaded',
                filePath
            });

            // Refresh the webview to show updated credentials info
            this._panel!.webview.html = await this._getHtmlForWebview();

        } catch (err) {
            vscode.window.showErrorMessage(`Failed to load credentials: ${err instanceof Error ? err.message : String(err)}`);
        }
    }

    private async _testConnection() {
        const snowflakeService = getSnowflakeService();

        try {
            const connection = await snowflakeService.createConnection();
            this._panel?.webview.postMessage({ command: 'connectionSuccess' });
        } catch (err) {
            this._panel?.webview.postMessage({
                command: 'connectionError',
                error: err instanceof Error ? err.message : String(err)
            });
        }
    }

    private async _createTemplate() {
        try {
            const options: vscode.SaveDialogOptions = {
                saveLabel: 'Create',
                filters: {
                    'JSON Files': ['json']
                },
                title: 'Create Snowflake Credentials Template'
            };

            // Try to use workspace folder or user's home directory as default
            if (vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0) {
                options.defaultUri = vscode.Uri.file(
                    path.join(vscode.workspace.workspaceFolders[0].uri.fsPath, 'snowflake-credentials.json')
                );
            }

            const fileUri = await vscode.window.showSaveDialog(options);
            if (!fileUri) {
                return;
            }

            const templateContent = JSON.stringify({
                account: "your-account",
                user: "your-username",
                password: "your-password",
                warehouse: "COMPUTE_WH",
                database: "your-database",
                schema: "your-schema"
            }, null, 2);

            fs.writeFileSync(fileUri.fsPath, templateContent, 'utf8');

            // Save the credentials file path to settings
            const config = vscode.workspace.getConfiguration('sql-sheets');
            await config.update('connection.credentialsFile', fileUri.fsPath, vscode.ConfigurationTarget.Global);

            // Update the UI
            this._panel?.webview.postMessage({
                command: 'templateCreated',
                filePath: fileUri.fsPath
            });

            // Refresh the webview to show updated credentials info
            this._panel!.webview.html = await this._getHtmlForWebview();

            // Open the file for editing
            await vscode.window.showTextDocument(fileUri);
        } catch (err) {
            vscode.window.showErrorMessage(`Failed to create template: ${err instanceof Error ? err.message : String(err)}`);
        }
    }

    private async _loadGoogleCredentialsFile() {
        try {
            // Show file picker
            const options: vscode.OpenDialogOptions = {
                canSelectMany: false,
                openLabel: 'Select Google Service Account File',
                filters: {
                    'JSON Files': ['json']
                }
            };

            const fileUri = await vscode.window.showOpenDialog(options);

            if (!fileUri || fileUri.length === 0) {
                return;
            }

            const filePath = fileUri[0].fsPath;

            // Read and parse the file to validate it's proper JSON
            const fileContent = fs.readFileSync(filePath, 'utf-8');
            const credentials = JSON.parse(fileContent);

            // Validate the structure
            if (!credentials.type || credentials.type !== 'service_account' || !credentials.project_id) {
                throw new Error('Invalid Google Service Account file. Must be a valid service account JSON file.');
            }

            // Save the credentials file path
            const config = vscode.workspace.getConfiguration('sql-sheets');
            await config.update('google.serviceAccountFile', filePath, vscode.ConfigurationTarget.Global);

            // Update the UI
            this._panel?.webview.postMessage({
                command: 'googleCredentialsLoaded',
                filePath
            });

            // Refresh the webview to show updated credentials info
            this._panel!.webview.html = await this._getHtmlForWebview();

        } catch (err) {
            vscode.window.showErrorMessage(`Failed to load Google Service Account file: ${err instanceof Error ? err.message : String(err)}`);
        }
    }

    private async _createGoogleTemplate() {
        try {
            const options: vscode.SaveDialogOptions = {
                saveLabel: 'Create',
                filters: {
                    'JSON Files': ['json']
                },
                title: 'Create Google Service Account Template'
            };

            // Try to use workspace folder or user's home directory as default
            if (vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0) {
                options.defaultUri = vscode.Uri.file(
                    path.join(vscode.workspace.workspaceFolders[0].uri.fsPath, 'google-credentials.json')
                );
            }

            const fileUri = await vscode.window.showSaveDialog(options);
            if (!fileUri) {
                return;
            }

            const templateContent = JSON.stringify({
                "type": "service_account",
                "project_id": "your-project-id",
                "private_key_id": "your-private-key-id",
                "private_key": "-----BEGIN PRIVATE KEY-----\nYOUR-PRIVATE-KEY\n-----END PRIVATE KEY-----\n",
                "client_email": "your-service-account@your-project-id.iam.gserviceaccount.com",
                "client_id": "your-client-id",
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/your-service-account%40your-project-id.iam.gserviceaccount.com",
                "universe_domain": "googleapis.com"
            }, null, 2);

            fs.writeFileSync(fileUri.fsPath, templateContent, 'utf8');

            // Save the credentials file path to settings
            const config = vscode.workspace.getConfiguration('sql-sheets');
            await config.update('google.serviceAccountFile', fileUri.fsPath, vscode.ConfigurationTarget.Global);

            // Update the UI
            this._panel?.webview.postMessage({
                command: 'googleTemplateCreated',
                filePath: fileUri.fsPath
            });

            // Refresh the webview to show updated credentials info
            this._panel!.webview.html = await this._getHtmlForWebview();

            // Open the file for editing
            await vscode.window.showTextDocument(fileUri);
        } catch (err) {
            vscode.window.showErrorMessage(`Failed to create Google template: ${err instanceof Error ? err.message : String(err)}`);
        }
    }
}