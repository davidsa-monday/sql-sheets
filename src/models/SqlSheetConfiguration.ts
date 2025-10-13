export class SqlSheetConfiguration {
    // Static property to store parameter descriptions
    public static readonly parameterDescriptions: Record<string, string> = {
        spreadsheet_id: "The Google Sheet ID where results will be sent",
        sheet_name: "The name or index of the sheet within the spreadsheet",
        start_cell: "The cell where output should begin (e.g., \"A1\"). Optionally prefix with a named range like \"MyRange | A1\" to combine both.",
        start_named_range: "Legacy alternative to start_cell for named ranges. Prefer combining with start_cell as 'MyRange | A1'.",
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

    /**
     * Determine whether this configuration contains either a start cell or a named range
     */
    public hasStartLocation(): boolean {
        const hasCell = typeof this.start_cell === 'string' && this.start_cell.trim().length > 0;
        const hasNamedRange = typeof this.start_named_range === 'string' && this.start_named_range.trim().length > 0;
        return hasCell || hasNamedRange;
    }

    /**
     * Parse a combined start parameter value and extract the optional named range and cell references.
     * Supports inputs like:
     *   - "MyNamedRange | B5"
     *   - "MyNamedRange"
     *   - "B5"
     *   - "offset 2" (no parsing, returned as cell)
     */
    public static parseStartCellParameter(value?: string): { startCell?: string; startNamedRange?: string } {
        if (!value) {
            return {};
        }

        const trimmedValue = value.trim();
        if (trimmedValue.length === 0) {
            return {};
        }

        const separatorIndex = trimmedValue.indexOf('|');
        if (separatorIndex === -1) {
            if (this.looksLikeCellReference(trimmedValue) || this.looksLikeOffsetDirective(trimmedValue)) {
                return { startCell: trimmedValue };
            }
            return { startNamedRange: trimmedValue };
        }

        const beforeSeparator = trimmedValue.substring(0, separatorIndex).trim();
        const afterSeparator = trimmedValue.substring(separatorIndex + 1).trim();

        const result: { startCell?: string; startNamedRange?: string } = {};
        if (beforeSeparator.length > 0) {
            result.startNamedRange = beforeSeparator;
        }
        if (afterSeparator.length > 0) {
            if (this.looksLikeCellReference(afterSeparator) || this.looksLikeOffsetDirective(afterSeparator)) {
                result.startCell = afterSeparator;
            } else {
                result.startCell = afterSeparator;
            }
        }

        return result;
    }

    private static looksLikeCellReference(value: string): boolean {
        const validCellFormat = /^[A-Za-z]+[0-9]+$/;
        return validCellFormat.test(value.trim());
    }

    private static looksLikeOffsetDirective(value: string): boolean {
        return /^offset\s+\d+$/i.test(value.trim());
    }

    /**
     * Format a combined start parameter value given an optional named range and cell reference.
     */
    public static formatStartCellParameter(namedRange?: string, startCell?: string): string {
        const trimmedRange = typeof namedRange === 'string' ? namedRange.trim() : '';
        const trimmedCell = typeof startCell === 'string' ? startCell.trim() : '';

        if (trimmedRange && trimmedCell) {
            return `${trimmedRange} | ${trimmedCell}`;
        }

        if (trimmedRange) {
            return trimmedRange;
        }

        return trimmedCell;
    }
}
