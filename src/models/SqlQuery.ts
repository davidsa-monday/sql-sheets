import * as vscode from 'vscode';
import { SqlSheetConfiguration } from './SqlSheetConfiguration';

export interface QueryDestination {
    readonly config: SqlSheetConfiguration;
    readonly parameterRanges: Map<string, { start: number; end: number }>;
}

export class SqlQuery {
    constructor(
        public readonly destinations: QueryDestination[],
        public readonly queryText: string,
        public readonly startOffset: number,
        public readonly endOffset: number,
        public readonly documentUri: vscode.Uri,
        public readonly parameterSources: Record<string, 'default' | 'query'>,
    ) { }

    public get config(): SqlSheetConfiguration {
        return this.destinations[0]?.config ?? new SqlSheetConfiguration();
    }

    public get configs(): SqlSheetConfiguration[] {
        return this.destinations.map(destination => destination.config);
    }

    public hasMultipleDestinations(): boolean {
        return this.destinations.length > 1;
    }
}
