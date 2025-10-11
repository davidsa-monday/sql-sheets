import { SqlSheetConfiguration } from "./SqlSheetConfiguration";

export class SqlQuery {
    constructor(
        public readonly config: SqlSheetConfiguration,
        public readonly queryText: string,
        public readonly startOffset: number,
        public readonly endOffset: number
    ) { }
}
