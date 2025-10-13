export class SqlSheetConfiguration {
    // Static property to store parameter descriptions
    public static readonly parameterDescriptions: Record<string, string> = {
        spreadsheet_id: "The Google Sheet ID where results will be sent",
        sheet_name: "The name or index of the sheet within the spreadsheet",
        start_cell: "The cell where output should begin (e.g., \"A1\") or use \"offset N\" to place N rows after the previous output",
        start_named_range: "Alternative to start_cell, can use a named range in Google Sheets",
        transpose: "Set to \"true\"/\"1\"/\"yes\" to transpose the output data (rows become columns)",
        name: "A title for the table (defaults to the SQL filename if not specified)",
        table_name: "An alternative name for the table",
        data_only: "Set to \"true\"/\"1\"/\"yes\" to output only data without headers",
        skip: "Set to \"true\"/\"1\"/\"yes\" to skip processing this query",
        pre_file: "Path to a SQL file that should be executed before this query"
    };

    // Boolean parameters list
    public static readonly booleanParameters: string[] = ['transpose', 'data_only', 'skip'];

    // Static method to convert string to boolean
    public static stringToBoolean(value?: string): boolean | undefined {
        if (value === undefined) {
            return undefined; // Let constructor use default value
        }
        const strValue = value.toLowerCase();
        return strValue === 'true' || strValue === '1' || strValue === 'yes';
    }

    // Static method to format boolean value to string
    public static formatBooleanToString(value: boolean | string): string {
        if (typeof value === 'boolean') {
            return value ? 'true' : 'false';
        }

        const strValue = value.toLowerCase();
        return (strValue === 'true' || strValue === '1' || strValue === 'yes') ? 'true' : 'false';
    }

    // Static method to check if a parameter is a boolean type
    public static isBooleanParameter(key: string): boolean {
        return this.booleanParameters.includes(key);
    }

    constructor(
        public readonly spreadsheet_id?: string,
        public readonly sheet_name?: string,
        public readonly start_cell?: string,
        public readonly start_named_range?: string,
        public readonly name?: string,
        public readonly table_name?: string,
        public readonly pre_file?: string,
        public readonly transpose: boolean = false,
        public readonly data_only: boolean = false,
        public readonly skip: boolean = false,
    ) { }
}
