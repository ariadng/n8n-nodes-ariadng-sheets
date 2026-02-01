import type {
    IExecuteFunctions,
    INodeExecutionData,
    INodeType,
    INodeTypeDescription,
    IDataObject,
    JsonObject,
} from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';

import {
    createClient,
    SheetsError,
    type SheetsClient,
    type RawCellValue,
    type ValueRange,
} from '@ariadng/sheets';

export class Sheets implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'Sheets',
        name: 'sheets',
        icon: 'file:sheets.svg',
        group: ['transform'],
        version: 1,
        subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
        description: 'Read and write data to Google Sheets using @ariadng/sheets',
        defaults: {
            name: 'Sheets',
        },
        inputs: ['main'],
        outputs: ['main'],
        credentials: [
            {
                name: 'sheetsServiceAccountApi',
                required: true,
            },
        ],
        properties: [
            // Resource
            {
                displayName: 'Resource',
                name: 'resource',
                type: 'options',
                noDataExpression: true,
                options: [
                    {
                        name: 'Spreadsheet',
                        value: 'spreadsheet',
                    },
                    {
                        name: 'Sheet',
                        value: 'sheet',
                    },
                    {
                        name: 'Values',
                        value: 'values',
                    },
                ],
                default: 'values',
            },

            // Operations - Spreadsheet
            {
                displayName: 'Operation',
                name: 'operation',
                type: 'options',
                noDataExpression: true,
                displayOptions: {
                    show: {
                        resource: ['spreadsheet'],
                    },
                },
                options: [
                    {
                        name: 'Get',
                        value: 'get',
                        description: 'Get spreadsheet metadata',
                        action: 'Get spreadsheet metadata',
                    },
                ],
                default: 'get',
            },

            // Operations - Sheet
            {
                displayName: 'Operation',
                name: 'operation',
                type: 'options',
                noDataExpression: true,
                displayOptions: {
                    show: {
                        resource: ['sheet'],
                    },
                },
                options: [
                    {
                        name: 'List',
                        value: 'list',
                        description: 'List all sheets in the spreadsheet',
                        action: 'List all sheets',
                    },
                ],
                default: 'list',
            },

            // Operations - Values
            {
                displayName: 'Operation',
                name: 'operation',
                type: 'options',
                noDataExpression: true,
                displayOptions: {
                    show: {
                        resource: ['values'],
                    },
                },
                options: [
                    {
                        name: 'Read',
                        value: 'read',
                        description: 'Read values from a range',
                        action: 'Read values from range',
                    },
                    {
                        name: 'Read Formulas',
                        value: 'readFormulas',
                        description: 'Read formulas from a range',
                        action: 'Read formulas from range',
                    },
                    {
                        name: 'Read Multiple',
                        value: 'readMultiple',
                        description: 'Read values from multiple ranges',
                        action: 'Read values from multiple ranges',
                    },
                    {
                        name: 'Write',
                        value: 'write',
                        description: 'Write values to a range',
                        action: 'Write values to range',
                    },
                    {
                        name: 'Write Multiple',
                        value: 'writeMultiple',
                        description: 'Write values to multiple ranges',
                        action: 'Write values to multiple ranges',
                    },
                    {
                        name: 'Append',
                        value: 'append',
                        description: 'Append rows to a table',
                        action: 'Append rows to table',
                    },
                    {
                        name: 'Clear',
                        value: 'clear',
                        description: 'Clear values from a range',
                        action: 'Clear values from range',
                    },
                    {
                        name: 'Clear Multiple',
                        value: 'clearMultiple',
                        description: 'Clear values from multiple ranges',
                        action: 'Clear values from multiple ranges',
                    },
                    {
                        name: 'Search',
                        value: 'search',
                        description: 'Search for values in the spreadsheet',
                        action: 'Search for values',
                    },
                ],
                default: 'read',
            },

            // Spreadsheet ID - Common
            {
                displayName: 'Spreadsheet ID',
                name: 'spreadsheetId',
                type: 'string',
                required: true,
                default: '',
                description: 'The ID of the spreadsheet. Found in the URL: docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit',
                placeholder: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
            },

            // Range - For read/write operations
            {
                displayName: 'Range',
                name: 'range',
                type: 'string',
                required: true,
                default: '',
                description: 'The A1 notation range (e.g., Sheet1!A1:D10 or A1:D10 for first sheet)',
                placeholder: 'Sheet1!A1:D10',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['read', 'readFormulas', 'write', 'append', 'clear'],
                    },
                },
            },

            // Ranges - For multiple operations
            {
                displayName: 'Ranges',
                name: 'ranges',
                type: 'string',
                required: true,
                default: '',
                description: 'Comma-separated list of A1 notation ranges (e.g., Sheet1!A1:B10, Sheet1!D1:E10)',
                placeholder: 'Sheet1!A1:B10, Sheet1!D1:E10',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['readMultiple', 'clearMultiple'],
                    },
                },
            },

            // Data for write operations
            {
                displayName: 'Data Mode',
                name: 'dataMode',
                type: 'options',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['write', 'append'],
                    },
                },
                options: [
                    {
                        name: 'From Input Items',
                        value: 'fromItems',
                        description: 'Use input items as rows',
                    },
                    {
                        name: 'JSON Array',
                        value: 'jsonArray',
                        description: 'Provide a 2D JSON array',
                    },
                ],
                default: 'fromItems',
            },

            // JSON data input
            {
                displayName: 'Data (JSON Array)',
                name: 'jsonData',
                type: 'json',
                default: '[["Column1", "Column2"], ["Value1", "Value2"]]',
                description: 'A 2D array of values to write',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['write', 'append'],
                        dataMode: ['jsonArray'],
                    },
                },
            },

            // Columns for fromItems mode
            {
                displayName: 'Columns',
                name: 'columns',
                type: 'string',
                default: '',
                description: 'Comma-separated list of column names from input items to write. Leave empty to use all columns.',
                placeholder: 'name, email, age',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['write', 'append'],
                        dataMode: ['fromItems'],
                    },
                },
            },

            // Include headers
            {
                displayName: 'Include Headers',
                name: 'includeHeaders',
                type: 'boolean',
                default: true,
                description: 'Whether to include column names as the first row',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['write', 'append'],
                        dataMode: ['fromItems'],
                    },
                },
            },

            // Write Multiple - Data
            {
                displayName: 'Range Data',
                name: 'rangeData',
                type: 'fixedCollection',
                typeOptions: {
                    multipleValues: true,
                },
                default: {},
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['writeMultiple'],
                    },
                },
                options: [
                    {
                        displayName: 'Range',
                        name: 'rangeValues',
                        values: [
                            {
                                displayName: 'Range',
                                name: 'range',
                                type: 'string',
                                default: '',
                                description: 'A1 notation range',
                            },
                            {
                                displayName: 'Data (JSON Array)',
                                name: 'data',
                                type: 'json',
                                default: '[["Value1", "Value2"]]',
                                description: 'A 2D array of values',
                            },
                        ],
                    },
                ],
            },

            // Search query
            {
                displayName: 'Search Query',
                name: 'searchQuery',
                type: 'string',
                required: true,
                default: '',
                description: 'The text or pattern to search for',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['search'],
                    },
                },
            },

            // Search range (optional)
            {
                displayName: 'Search Range',
                name: 'searchRange',
                type: 'string',
                default: '',
                description: 'Optional range to search within. Leave empty to search all sheets.',
                placeholder: 'Sheet1!A1:D100',
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['search'],
                    },
                },
            },

            // Search options
            {
                displayName: 'Search Options',
                name: 'searchOptions',
                type: 'collection',
                placeholder: 'Add Option',
                default: {},
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['search'],
                    },
                },
                options: [
                    {
                        displayName: 'Case Sensitive',
                        name: 'caseSensitive',
                        type: 'boolean',
                        default: false,
                        description: 'Whether the search should be case-sensitive',
                    },
                    {
                        displayName: 'Exact Match',
                        name: 'exactMatch',
                        type: 'boolean',
                        default: false,
                        description: 'Whether to match the entire cell value (not just contains)',
                    },
                    {
                        displayName: 'Use Regex',
                        name: 'regex',
                        type: 'boolean',
                        default: false,
                        description: 'Whether to treat the query as a regular expression',
                    },
                    {
                        displayName: 'Limit',
                        name: 'limit',
                        type: 'number',
                        default: 0,
                        description: 'Maximum number of results to return (0 = unlimited)',
                    },
                ],
            },

            // Options
            {
                displayName: 'Options',
                name: 'options',
                type: 'collection',
                placeholder: 'Add Option',
                default: {},
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['read', 'readFormulas', 'readMultiple'],
                    },
                },
                options: [
                    {
                        displayName: 'Value Render Option',
                        name: 'valueRenderOption',
                        type: 'options',
                        options: [
                            {
                                name: 'Formatted Value',
                                value: 'FORMATTED_VALUE',
                                description: 'Values as displayed in the UI',
                            },
                            {
                                name: 'Unformatted Value',
                                value: 'UNFORMATTED_VALUE',
                                description: 'Raw values without formatting',
                            },
                            {
                                name: 'Formula',
                                value: 'FORMULA',
                                description: 'Formulas if present, otherwise values',
                            },
                        ],
                        default: 'FORMATTED_VALUE',
                    },
                    {
                        displayName: 'Date Time Render Option',
                        name: 'dateTimeRenderOption',
                        type: 'options',
                        options: [
                            {
                                name: 'Formatted String',
                                value: 'FORMATTED_STRING',
                                description: 'Dates as formatted strings',
                            },
                            {
                                name: 'Serial Number',
                                value: 'SERIAL_NUMBER',
                                description: 'Dates as Excel serial numbers',
                            },
                        ],
                        default: 'FORMATTED_STRING',
                    },
                    {
                        displayName: 'Major Dimension',
                        name: 'majorDimension',
                        type: 'options',
                        options: [
                            {
                                name: 'Rows',
                                value: 'ROWS',
                                description: 'Return data as rows',
                            },
                            {
                                name: 'Columns',
                                value: 'COLUMNS',
                                description: 'Return data as columns',
                            },
                        ],
                        default: 'ROWS',
                    },
                ],
            },

            // Write options
            {
                displayName: 'Write Options',
                name: 'writeOptions',
                type: 'collection',
                placeholder: 'Add Option',
                default: {},
                displayOptions: {
                    show: {
                        resource: ['values'],
                        operation: ['write', 'writeMultiple', 'append'],
                    },
                },
                options: [
                    {
                        displayName: 'Value Input Option',
                        name: 'valueInputOption',
                        type: 'options',
                        options: [
                            {
                                name: 'User Entered',
                                value: 'USER_ENTERED',
                                description: 'Parse values as if typed by user (formulas evaluated)',
                            },
                            {
                                name: 'Raw',
                                value: 'RAW',
                                description: 'Store values exactly as provided',
                            },
                        ],
                        default: 'USER_ENTERED',
                    },
                    {
                        displayName: 'Insert Data Option',
                        name: 'insertDataOption',
                        type: 'options',
                        displayOptions: {
                            show: {
                                '/operation': ['append'],
                            },
                        },
                        options: [
                            {
                                name: 'Overwrite',
                                value: 'OVERWRITE',
                                description: 'Add rows after the last row with data',
                            },
                            {
                                name: 'Insert Rows',
                                value: 'INSERT_ROWS',
                                description: 'Insert new rows, pushing existing data down',
                            },
                        ],
                        default: 'OVERWRITE',
                    },
                ],
            },
        ],
    };

    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
        const items = this.getInputData();
        const returnData: INodeExecutionData[] = [];

        const resource = this.getNodeParameter('resource', 0) as string;
        const operation = this.getNodeParameter('operation', 0) as string;
        const spreadsheetId = this.getNodeParameter('spreadsheetId', 0) as string;

        // Get credentials and create client
        const credentials = await this.getCredentials('sheetsServiceAccountApi');
        let client: SheetsClient;

        try {
            const serviceAccountJson = JSON.parse(credentials.serviceAccountJson as string);
            client = createClient({
                auth: {
                    type: 'service-account',
                    credentials: serviceAccountJson,
                },
            });
        } catch (error) {
            throw new NodeOperationError(
                this.getNode(),
                'Invalid Service Account JSON. Please check your credentials.',
                { description: 'The JSON must contain valid service account credentials with client_email and private_key.' }
            );
        }

        try {
            // Spreadsheet operations
            if (resource === 'spreadsheet') {
                if (operation === 'get') {
                    const spreadsheet = await client.getSpreadsheet(spreadsheetId);
                    returnData.push({
                        json: spreadsheet as unknown as IDataObject,
                        pairedItem: { item: 0 },
                    });
                }
            }

            // Sheet operations
            if (resource === 'sheet') {
                if (operation === 'list') {
                    const sheets = await client.getSheets(spreadsheetId);
                    for (let i = 0; i < sheets.length; i++) {
                        returnData.push({
                            json: sheets[i] as unknown as IDataObject,
                            pairedItem: { item: 0 },
                        });
                    }
                }
            }

            // Values operations
            if (resource === 'values') {
                if (operation === 'read' || operation === 'readFormulas') {
                    const range = this.getNodeParameter('range', 0) as string;
                    const options = this.getNodeParameter('options', 0, {}) as IDataObject;

                    let valueRange: ValueRange;
                    if (operation === 'readFormulas') {
                        valueRange = await client.getFormulas(spreadsheetId, range);
                    } else {
                        valueRange = await client.getValues(spreadsheetId, range, {
                            valueRenderOption: options.valueRenderOption as 'FORMATTED_VALUE' | 'UNFORMATTED_VALUE' | 'FORMULA',
                            dateTimeRenderOption: options.dateTimeRenderOption as 'SERIAL_NUMBER' | 'FORMATTED_STRING',
                            majorDimension: options.majorDimension as 'ROWS' | 'COLUMNS',
                        });
                    }

                    // Convert to n8n format - each row becomes an item
                    const values = valueRange.values;
                    if (values.length > 0) {
                        // Use first row as headers if it looks like headers
                        const firstRow = values[0];
                        const headers = firstRow.map((cell, idx) =>
                            cell.value !== null && cell.value !== undefined
                                ? String(cell.value)
                                : `column_${idx}`
                        );

                        // Check if first row looks like headers (all strings)
                        const isFirstRowHeaders = firstRow.every(cell =>
                            typeof cell.value === 'string'
                        );

                        const dataRows = isFirstRowHeaders ? values.slice(1) : values;
                        const useHeaders = isFirstRowHeaders ? headers : headers.map((_, idx) => `column_${idx}`);

                        for (let rowIdx = 0; rowIdx < dataRows.length; rowIdx++) {
                            const row = dataRows[rowIdx];
                            const rowData: IDataObject = {
                                _range: valueRange.range,
                                _rowIndex: isFirstRowHeaders ? rowIdx + 2 : rowIdx + 1,
                            };

                            for (let colIdx = 0; colIdx < useHeaders.length; colIdx++) {
                                const cell = row[colIdx];
                                rowData[useHeaders[colIdx]] = cell?.value ?? null;
                            }

                            returnData.push({
                                json: rowData,
                                pairedItem: { item: 0 },
                            });
                        }
                    }
                }

                if (operation === 'readMultiple') {
                    const rangesStr = this.getNodeParameter('ranges', 0) as string;
                    const ranges = rangesStr.split(',').map(r => r.trim());
                    const options = this.getNodeParameter('options', 0, {}) as IDataObject;

                    const result = await client.batchGetValues(spreadsheetId, ranges, {
                        valueRenderOption: options.valueRenderOption as 'FORMATTED_VALUE' | 'UNFORMATTED_VALUE' | 'FORMULA',
                        dateTimeRenderOption: options.dateTimeRenderOption as 'SERIAL_NUMBER' | 'FORMATTED_STRING',
                        majorDimension: options.majorDimension as 'ROWS' | 'COLUMNS',
                    });

                    for (const vr of result.valueRanges) {
                        for (let rowIdx = 0; rowIdx < vr.values.length; rowIdx++) {
                            const row = vr.values[rowIdx];
                            const rowData: IDataObject = {
                                _range: vr.range,
                                _rowIndex: rowIdx + 1,
                            };

                            for (let colIdx = 0; colIdx < row.length; colIdx++) {
                                rowData[`column_${colIdx}`] = row[colIdx]?.value ?? null;
                            }

                            returnData.push({
                                json: rowData,
                                pairedItem: { item: 0 },
                            });
                        }
                    }
                }

                if (operation === 'write') {
                    const range = this.getNodeParameter('range', 0) as string;
                    const dataMode = this.getNodeParameter('dataMode', 0) as string;
                    const writeOptions = this.getNodeParameter('writeOptions', 0, {}) as IDataObject;

                    let values: RawCellValue[][];

                    if (dataMode === 'jsonArray') {
                        const jsonData = this.getNodeParameter('jsonData', 0) as string;
                        values = JSON.parse(jsonData);
                    } else {
                        // From input items
                        const columnsStr = this.getNodeParameter('columns', 0, '') as string;
                        const includeHeaders = this.getNodeParameter('includeHeaders', 0, true) as boolean;

                        const columns = columnsStr
                            ? columnsStr.split(',').map(c => c.trim())
                            : Object.keys(items[0]?.json || {}).filter(k => !k.startsWith('_'));

                        values = [];
                        if (includeHeaders) {
                            values.push(columns);
                        }

                        for (const item of items) {
                            const row: RawCellValue[] = columns.map(col => {
                                const val = item.json[col];
                                if (val === undefined || val === null) return null;
                                if (typeof val === 'object') return JSON.stringify(val);
                                return val as RawCellValue;
                            });
                            values.push(row);
                        }
                    }

                    const result = await client.updateValues(spreadsheetId, range, values, {
                        valueInputOption: (writeOptions.valueInputOption as 'RAW' | 'USER_ENTERED') || 'USER_ENTERED',
                    });

                    returnData.push({
                        json: result as unknown as IDataObject,
                        pairedItem: { item: 0 },
                    });
                }

                if (operation === 'writeMultiple') {
                    const rangeData = this.getNodeParameter('rangeData', 0) as IDataObject;
                    const writeOptions = this.getNodeParameter('writeOptions', 0, {}) as IDataObject;

                    const data: { range: string; values: RawCellValue[][] }[] = [];
                    const rangeValues = (rangeData.rangeValues as IDataObject[]) || [];

                    for (const rv of rangeValues) {
                        data.push({
                            range: rv.range as string,
                            values: JSON.parse(rv.data as string),
                        });
                    }

                    const result = await client.batchUpdateValues(spreadsheetId, data, {
                        valueInputOption: (writeOptions.valueInputOption as 'RAW' | 'USER_ENTERED') || 'USER_ENTERED',
                    });

                    returnData.push({
                        json: result as unknown as IDataObject,
                        pairedItem: { item: 0 },
                    });
                }

                if (operation === 'append') {
                    const range = this.getNodeParameter('range', 0) as string;
                    const dataMode = this.getNodeParameter('dataMode', 0) as string;
                    const writeOptions = this.getNodeParameter('writeOptions', 0, {}) as IDataObject;

                    let values: RawCellValue[][];

                    if (dataMode === 'jsonArray') {
                        const jsonData = this.getNodeParameter('jsonData', 0) as string;
                        values = JSON.parse(jsonData);
                    } else {
                        const columnsStr = this.getNodeParameter('columns', 0, '') as string;

                        const columns = columnsStr
                            ? columnsStr.split(',').map(c => c.trim())
                            : Object.keys(items[0]?.json || {}).filter(k => !k.startsWith('_'));

                        values = [];
                        for (const item of items) {
                            const row: RawCellValue[] = columns.map(col => {
                                const val = item.json[col];
                                if (val === undefined || val === null) return null;
                                if (typeof val === 'object') return JSON.stringify(val);
                                return val as RawCellValue;
                            });
                            values.push(row);
                        }
                    }

                    const result = await client.appendValues(spreadsheetId, range, values, {
                        valueInputOption: (writeOptions.valueInputOption as 'RAW' | 'USER_ENTERED') || 'USER_ENTERED',
                        insertDataOption: (writeOptions.insertDataOption as 'OVERWRITE' | 'INSERT_ROWS') || 'OVERWRITE',
                    });

                    returnData.push({
                        json: result as unknown as IDataObject,
                        pairedItem: { item: 0 },
                    });
                }

                if (operation === 'clear') {
                    const range = this.getNodeParameter('range', 0) as string;
                    const result = await client.clearValues(spreadsheetId, range);

                    returnData.push({
                        json: result as unknown as IDataObject,
                        pairedItem: { item: 0 },
                    });
                }

                if (operation === 'clearMultiple') {
                    const rangesStr = this.getNodeParameter('ranges', 0) as string;
                    const ranges = rangesStr.split(',').map(r => r.trim());

                    const result = await client.batchClearValues(spreadsheetId, ranges);

                    returnData.push({
                        json: result as unknown as IDataObject,
                        pairedItem: { item: 0 },
                    });
                }

                if (operation === 'search') {
                    const query = this.getNodeParameter('searchQuery', 0) as string;
                    const searchRange = this.getNodeParameter('searchRange', 0, '') as string;
                    const searchOptions = this.getNodeParameter('searchOptions', 0, {}) as IDataObject;

                    const result = await client.searchValues(spreadsheetId, query, {
                        range: searchRange || undefined,
                        caseSensitive: searchOptions.caseSensitive as boolean,
                        exactMatch: searchOptions.exactMatch as boolean,
                        regex: searchOptions.regex as boolean,
                        limit: (searchOptions.limit as number) || undefined,
                    });

                    // Return each match as a separate item
                    for (const match of result.matches) {
                        returnData.push({
                            json: {
                                query: result.query,
                                matchType: result.matchType,
                                sheet: match.sheet,
                                sheetId: match.sheetId,
                                address: match.address,
                                row: match.row,
                                column: match.column,
                                value: match.value,
                            },
                            pairedItem: { item: 0 },
                        });
                    }

                    // If no matches, return summary
                    if (result.matches.length === 0) {
                        returnData.push({
                            json: {
                                query: result.query,
                                matchType: result.matchType,
                                totalMatches: 0,
                                message: 'No matches found',
                            },
                            pairedItem: { item: 0 },
                        });
                    }
                }
            }

            return [returnData];
        } catch (error) {
            if (error instanceof SheetsError) {
                throw new NodeApiError(this.getNode(), {
                    message: error.message,
                    code: error.code,
                    status: error.status,
                } as JsonObject, {
                    message: `Google Sheets API Error: ${error.message}`,
                    description: `Status: ${error.status}, Code: ${error.code}`,
                });
            }
            throw error;
        }
    }
}
