import * as vscode from 'vscode';
import * as snowflake from 'snowflake-sdk';
import * as fs from 'fs';
import { getLogger } from './loggingService';

/**
 * Service for handling Snowflake database connections and queries
 */
export class SnowflakeService {
    private connection: snowflake.Connection | null = null;
    private connectionPromise: Promise<snowflake.Connection> | null = null;
    private readonly logger = getLogger();

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
        if (this.connectionPromise) {
            return this.connectionPromise;
        }

        this.connectionPromise = (async () => {
            // Close any existing connection
            await this.closeConnection();

            // Load credentials from configuration
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
            return await new Promise<snowflake.Connection>((resolve, reject) => {
                if (!this.connection) {
                    reject(new Error('Failed to create Snowflake connection'));
                    return;
                }

                this.connection.connect((err: Error | undefined | null, conn: snowflake.Connection) => {
                    if (err) {
                        reject(err);
                        return;
                    }
                    this.logger.debug('Connected to Snowflake.', { audience: ['developer'] });
                    resolve(conn);
                });
            });
        })();

        try {
            return await this.connectionPromise;
        } finally {
            this.connectionPromise = null;
        }
    }

    /**
     * Executes a SQL query and returns the results
     * @param query SQL query string to execute
     * @returns Query results
     */
    public async executeQuery(query: string): Promise<any[]> {
        return new Promise<any[]>(async (resolve, reject) => {
            try {
                await this.ensureActiveConnection();
            } catch (err) {
                reject(err);
                return;
            }

            const attemptExecution = (retrying: boolean) => {
                if (!this.connection) {
                    reject(new Error('No active Snowflake connection'));
                    return;
                }

                this.connection.execute({
                    sqlText: query,
                    // Preserve duplicate column names by letting the driver append numeric suffixes
                    rowMode: 'object_with_renamed_duplicated_columns',
                    complete: (err: Error | undefined | null, stmt: any, rows: any[] | undefined) => {
                        if (err) {
                            if (!retrying && this.shouldAttemptReconnect(err)) {
                                this.logger.warn('Snowflake connection interrupted. Attempting to reconnect...', {
                                    audience: ['developer'],
                                    data: err
                                });
                                this.invalidateConnection();
                                this.createConnection()
                                    .then(() => attemptExecution(true))
                                    .catch(reject);
                                return;
                            }
                            reject(err);
                            return;
                        }
                        resolve(rows || []);
                    },
                    parameters: {
                        MULTI_STATEMENT_COUNT: 0  // Set to 0 to allow unlimited statements or a specific number
                    }
                });
            };

            attemptExecution(false);
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
                this.connection.destroy(() => {
                    this.connection = null;
                    resolve();
                });
            } else {
                resolve();
            }
        });
    }

    private async ensureActiveConnection(): Promise<snowflake.Connection> {
        if (this.connection && typeof this.connection.isUp === 'function' && this.connection.isUp()) {
            return this.connection;
        }
        return this.createConnection();
    }

    private shouldAttemptReconnect(err: unknown): boolean {
        if (!err || typeof err !== 'object') {
            return false;
        }

        const snowflakeError = err as snowflake.SnowflakeError;
        const sqlState = snowflakeError.sqlState;
        const code = snowflakeError.code;
        const isFatal = Boolean((snowflakeError as { isFatal?: boolean }).isFatal);
        const reconnectableSqlStates = new Set(['08002', '08003', '08006']);
        const reconnectableCodes = new Set([405503, 407001, 407002]);

        if (sqlState && reconnectableSqlStates.has(sqlState)) {
            return true;
        }

        if (typeof code === 'number' && reconnectableCodes.has(code)) {
            return true;
        }

        if (isFatal) {
            return true;
        }

        const message = typeof snowflakeError.message === 'string'
            ? snowflakeError.message.toLowerCase()
            : '';

        const reconnectablePhrases = [
            'session no longer exists',
            'connection is closed',
            'connection already closed',
            'session not found',
            'invalid connection'
        ];

        return reconnectablePhrases.some(phrase => message.includes(phrase));
    }

    private invalidateConnection(): void {
        if (this.connection) {
            try {
                this.connection.destroy(() => { /* noop */ });
            } catch {
                // Ignore errors while tearing down a broken connection
            }
        }
        this.connection = null;
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
