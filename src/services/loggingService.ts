import * as vscode from 'vscode';

type LogLevel = 'DEBUG' | 'INFO' | 'WARN' | 'ERROR';
type LogAudience = 'support' | 'developer';

interface LogOptions {
	/** Which channels should receive the message. Defaults depend on level. */
	audience?: LogAudience[];
	/** Optional structured payload to include. */
	data?: unknown;
	/** Reveal the output panel for support-facing logs. */
	reveal?: boolean;
}

class LoggingService {
	private readonly outputChannel: vscode.OutputChannel;

	constructor() {
		this.outputChannel = vscode.window.createOutputChannel('SQL Sheets');
	}

	public debug(message: string, options: LogOptions = {}): void {
		this.log('DEBUG', message, { ...options, audience: options.audience ?? ['developer'] });
	}

	public info(message: string, options: LogOptions = {}): void {
		this.log('INFO', message, { ...options, audience: options.audience ?? ['developer'] });
	}

	public warn(message: string, options: LogOptions = {}): void {
		this.log('WARN', message, { ...options, audience: options.audience ?? ['support'] });
	}

	public error(message: string, options: LogOptions = {}): void {
		this.log('ERROR', message, { ...options, audience: options.audience ?? ['support', 'developer'] });
	}

	public async notifyInfo(message: string, ...items: string[]): Promise<string | undefined> {
		this.info(message);
		return vscode.window.showInformationMessage(message, ...items);
	}

	public async notifyWarning(message: string, ...items: string[]): Promise<string | undefined> {
		this.warn(message);
		return vscode.window.showWarningMessage(message, ...items);
	}

	public async notifyError(message: string, ...items: string[]): Promise<string | undefined> {
		this.error(message);
		return vscode.window.showErrorMessage(message, ...items);
	}

	public revealOutput(preserveFocus = true): void {
		this.outputChannel.show(preserveFocus);
	}

	private log(level: LogLevel, message: string, options: LogOptions): void {
		const timestamp = new Date().toISOString();
		const formatted = `[${timestamp}] [${level}] ${message}`;
		const audiences = options.audience ?? ['support'];

		if (audiences.includes('support')) {
			this.outputChannel.appendLine(formatted);
			if (options.data !== undefined) {
				for (const line of this.stringifyData(options.data).split('\n')) {
					this.outputChannel.appendLine(line);
				}
			}
			if (options.reveal) {
				this.outputChannel.show(true);
			}
		}

		if (audiences.includes('developer')) {
			const consoleMethod = this.consoleMethodFor(level);
			if (options.data !== undefined) {
				consoleMethod(`[SQL Sheets] [${level}] ${message}`, options.data);
			} else {
				consoleMethod(`[SQL Sheets] [${level}] ${message}`);
			}
		}
	}

	private consoleMethodFor(level: LogLevel): (...args: unknown[]) => void {
		switch (level) {
			case 'ERROR':
				return console.error;
			case 'WARN':
				return console.warn;
			default:
				return console.log;
		}
	}

	private stringifyData(data: unknown): string {
		if (data instanceof Error) {
			return data.stack ?? data.message;
		}
		if (typeof data === 'string') {
			return data;
		}
		try {
			return JSON.stringify(data, null, 2);
		} catch {
			return String(data);
		}
	}
}

const loggingService = new LoggingService();

export function getLogger(): LoggingService {
	return loggingService;
}

