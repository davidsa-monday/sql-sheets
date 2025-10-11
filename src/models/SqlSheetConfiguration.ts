export class SqlSheetConfiguration {
    constructor(
        public readonly spreadsheet_id?: string,
        public readonly sheet_name?: string,
        public readonly start_cell?: string,
        public readonly start_named_range?: string,
        public readonly table_name?: string
    ) { }
}
