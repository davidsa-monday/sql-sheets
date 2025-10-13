import * as vscode from 'vscode';
import * as snowflake from 'snowflake-sdk';
import * as fs from 'fs';

/**
 * Service for handling Snowflake database connections and queries
 */
export class SnowflakeService {
    private connection: snowflake.Connection | null = null;

    constructor() {
        // Initialize snowflake driver
        snowflake.configure({ insecureConnect: false });
    }

    /**
     * Load credentials from file
     */
    private loadCredentialsFromFile(): {
        user: string | undefined;
        password: string | undefined;
        account: string | undefined;
        warehouse: string | undefined;
        database: string | undefined;
        schema: string | undefined;
    } {
        const config = vscode.workspace.getConfiguration('sql-sheets.connection');
        const credentialsFile = config.get<string>('credentialsFile');

        if (!credentialsFile) {
            throw new Error('No credentials file configured. Please set a credentials file in the settings.');
        }

        if (!fs.existsSync(credentialsFile)) {
            throw new Error(`Credentials file does not exist: ${credentialsFile}`);
        }

        try {
            const fileContent = fs.readFileSync(credentialsFile, 'utf-8');
            const credentials = JSON.parse(fileContent);

            if (!credentials.user || !credentials.account) {
                throw new Error('Invalid credentials file. Must contain at least "user" and "account" fields.');
            }

            return {
                user: credentials.user,
                password: credentials.password,
                account: credentials.account,
                warehouse: credentials.warehouse,
                database: credentials.database,
                schema: credentials.schema,
            };
        } catch (err) {
            if (err instanceof Error) {
                throw new Error(`Failed to load credentials from file: ${err.message}`);
            } else {
                throw new Error('Failed to load credentials from file');
            }
        }
    }

    /**
     * Creates a Snowflake connection using credentials from the extension settings
     */
    public async createConnection(): Promise<snowflake.Connection> {
        // Close any existing connection
        await this.closeConnection();

        // Try to load credentials from file first, then fall back to settings
        const credentials = this.loadCredentialsFromFile();
        const { user, password, account, warehouse, database, schema } = credentials;

        // Validate required fields
        if (!user || !password || !account) {
            throw new Error('Missing required Snowflake connection parameters (user, password, or account)');
        }

        // Create connection options
        const connectionOptions: snowflake.ConnectionOptions = {
            account,
            username: user,
            password,
            application: 'SQLSheetsVSCodeExtension',
        };

        if (warehouse) {
            connectionOptions.warehouse = warehouse;
        }

        if (database) {
            connectionOptions.database = database;
        }

        if (schema) {
            connectionOptions.schema = schema;
        }

        // Create connection
        this.connection = snowflake.createConnection(connectionOptions);

        // Connect and return the connection
        return new Promise<snowflake.Connection>((resolve, reject) => {
            if (!this.connection) {
                reject(new Error('Failed to create Snowflake connection'));
                return;
            }

            this.connection.connect((err: Error | undefined | null, conn: snowflake.Connection) => {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(conn);
            });
        });
    }

    /**
     * Executes a SQL query and returns the results
     * @param query SQL query string to execute
     * @returns Query results
     */
    public async executeQuery(query: string): Promise<any[]> {
        if (!this.connection) {
            await this.createConnection();
        }

        if (!this.connection) {
            throw new Error('No active Snowflake connection');
        }

        return new Promise<any[]>((resolve, reject) => {
            this.connection!.execute({
                sqlText: query,
                // Preserve duplicate column names by letting the driver append numeric suffixes
                rowMode: 'object_with_renamed_duplicated_columns',
                complete: (err: Error | undefined | null, stmt: any, rows: any[] | undefined) => {
                    if (err) {
                        reject(err);
                        return;
                    }
                    resolve(rows || []);
                },
                parameters: {
                    MULTI_STATEMENT_COUNT: 0  // Set to 0 to allow unlimited statements or a specific number
                }
            });
        });
    }

    /**
     * Checks if the connection settings are configured
     */
    public isConfigured(): boolean {
        const credentials = this.loadCredentialsFromFile();
        const { user, password, account } = credentials;

        return Boolean(user && password && account);
    }

    /**
     * Closes the current connection
     */
    public async closeConnection(): Promise<void> {
        return new Promise<void>((resolve) => {
            if (this.connection) {
                this.connection.destroy((err: Error | undefined | null) => {
                    this.connection = null;
                    resolve();
                });
            } else {
                resolve();
            }
        });
    }
}

// Singleton instance
let snowflakeService: SnowflakeService | undefined;

/**
 * Get the snowflake service instance
 */
export function getSnowflakeService(): SnowflakeService {
    if (!snowflakeService) {
        snowflakeService = new SnowflakeService();
    }
    return snowflakeService;
}
