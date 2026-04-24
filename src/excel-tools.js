'use strict';

const { appendRows, getHeaders, updateCell } = require('./excel-writer');
const { resolveOutputPath } = require('./config');

const EXCEL_TOOLS = [
  {
    name: 'append_row',
    description: [
      'Append one or more rows to an Excel (.xlsx) file.',
      'The file is created automatically if it does not exist.',
      'Supports positional rows and key-value data matched to the live header row.',
    ].join(' '),
    inputSchema: {
      type: 'object',
      required: ['file_path'],
      properties: {
        file_path: { type: 'string', description: 'Absolute path or filename relative to OUTPUT_DIR.' },
        sheet: { type: 'string', description: 'Worksheet name. Defaults to the first sheet.' },
        data: { type: 'object', description: 'Key-value row mapped by the workbook header row.' },
        row: { type: 'array', description: 'Single row as an array of cell values.', items: {} },
        rows: { type: 'array', description: 'Multiple rows: array of row arrays.', items: { type: 'array', items: {} } },
        header: { type: 'array', description: 'Column headers. Written only when creating a new file.', items: { type: 'string' } },
      },
    },
  },
  {
    name: 'update_cell',
    description: [
      'Update one or more cells in the last populated data row of an Excel (.xlsx) file.',
      'Use column+value for one cell, or header+row for multiple cells.',
    ].join(' '),
    inputSchema: {
      type: 'object',
      required: ['file_path'],
      properties: {
        file_path: { type: 'string', description: 'Absolute path or filename relative to OUTPUT_DIR.' },
        sheet: { type: 'string', description: 'Worksheet name. Defaults to the first sheet.' },
        row_number: { type: 'number', description: 'Backward-compatible field. Current writer targets the last data row.' },
        column: { description: 'Column name from header row or 1-based column index.' },
        value: { description: 'New value for the cell.' },
        header: { type: 'array', description: 'Column names for multi-cell update.', items: { type: 'string' } },
        row: { type: 'array', description: 'Values parallel to header. Empty values are skipped.', items: {} },
      },
    },
  },
  {
    name: 'append_edit_record',
    description: [
      'Append a revision record row to an Excel 改正管理表.',
      'Maps document_meta to columns by live header names where possible.',
    ].join(' '),
    inputSchema: {
      type: 'object',
      required: ['file_path', 'document_meta'],
      properties: {
        file_path: { type: 'string', description: 'Absolute path or filename relative to OUTPUT_DIR.' },
        sheet: { type: 'string', description: 'Worksheet name. Defaults to the first sheet.' },
        document_meta: {
          type: 'object',
          properties: {
            file_type: { type: ['string', 'null'], description: '文書区分: 基準 | マニュアル | null' },
            file_name: { type: 'string', description: '規程・基準・マニュアル名称' },
            main_changes: { type: 'string', description: '内容（改正概要）' },
          },
        },
      },
    },
  },
  {
    name: 'get_excel_schema',
    description: 'Return Excel API usage guidance. Prefer read_excel_headers for a live workbook schema.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_excel_headers',
    description: 'Read the first row of an existing Excel file as a live schema for key-value row writes.',
    inputSchema: {
      type: 'object',
      required: ['file_path'],
      properties: {
        file_path: { type: 'string', description: 'Absolute path or filename relative to OUTPUT_DIR.' },
        sheet: { type: 'string', description: 'Worksheet name. Defaults to the first sheet.' },
      },
    },
  },
  {
    name: 'get_schema',
    description: 'Backward-compatible alias for get_excel_schema.',
    inputSchema: { type: 'object', properties: {} },
  },
];

function excelSchemaDescription() {
  return {
    description: 'Excel workbook append/update tools',
    recommended_flow: [
      'Call read_excel_headers for an existing workbook.',
      'Ask the LLM to output key-value data using those header names.',
      'Call append_row with data for robust column mapping.',
    ],
    append_row: {
      file_path: 'string - absolute path or filename relative to OUTPUT_DIR',
      sheet: 'string? - worksheet name',
      data: 'object? - key-value pairs keyed by header name, recommended for existing files',
      row: 'array? - positional row',
      rows: 'array? - multiple positional rows',
      header: 'array? - header row when creating a new file',
    },
    append_edit_record: {
      document_meta: {
        file_type: '文書区分',
        file_name: '規程・基準・マニュアル名称',
        main_changes: '内容',
      },
    },
  };
}

async function handleAppendRow(args = {}) {
  const { sheet, row, rows, data, header } = args;
  const file_path = resolveOutputPath(args.file_path);

  if (!file_path) return toolError('file_path is required');
  if (!data && !row && (!rows || rows.length === 0)) {
    return toolError('Provide data, row, or rows');
  }

  try {
    const result = await appendRows({ file_path, sheet, row, rows, data, header });
    return toolText([
      'Row(s) appended successfully',
      `file: ${result.file_path}`,
      `sheet: ${result.sheet}`,
      `appended_rows: ${result.appended_rows}`,
      `total_rows: ${result.total_rows}`,
    ]);
  } catch (error) {
    return toolError(error.message);
  }
}

async function handleUpdateCell(args = {}) {
  const file_path = resolveOutputPath(args.file_path);
  if (!file_path) return toolError('file_path is required');
  if (args.column === undefined && (!args.header || !args.row)) {
    return toolError('Provide either column+value or header+row');
  }

  try {
    const result = await updateCell({ ...args, file_path });
    return toolText([
      'Cell(s) updated successfully',
      `file: ${result.file_path}`,
      `sheet: ${result.sheet}`,
      `updated_cells: ${result.updated_cells}`,
      `excel_row: ${result.excel_row}`,
    ]);
  } catch (error) {
    return toolError(error.message);
  }
}

async function handleAppendEditRecord(args = {}) {
  const file_path = resolveOutputPath(args.file_path);
  const documentMeta = args.document_meta || {};

  if (!file_path) return toolError('file_path is required');
  if (!documentMeta.file_type && !documentMeta.file_name && !documentMeta.main_changes) {
    return toolError('document_meta with at least one field is required');
  }

  const data = {};
  if (documentMeta.file_type) data['文書区分'] = documentMeta.file_type;
  if (documentMeta.file_name) data['規程・基準・マニュアル名称'] = documentMeta.file_name;
  if (documentMeta.main_changes) data['内容'] = documentMeta.main_changes;

  try {
    const result = await appendRows({ file_path, sheet: args.sheet, data });
    return toolText([
      'Edit record appended successfully',
      `file: ${result.file_path}`,
      `sheet: ${result.sheet}`,
      `total_rows: ${result.total_rows}`,
    ]);
  } catch (error) {
    return toolError(error.message);
  }
}

async function handleReadExcelHeaders(args = {}) {
  const file_path = resolveOutputPath(args.file_path);
  if (!file_path) return toolError('file_path is required');

  try {
    const headers = await getHeaders(file_path, args.sheet);
    return toolJson({ file_path, sheet: args.sheet || null, headers });
  } catch (error) {
    return toolError(error.message);
  }
}

async function handleExcelTool(name, args = {}) {
  switch (name) {
    case 'append_row': return handleAppendRow(args);
    case 'update_cell': return handleUpdateCell(args);
    case 'append_edit_record': return handleAppendEditRecord(args);
    case 'read_excel_headers': return handleReadExcelHeaders(args);
    case 'get_excel_schema':
    case 'get_schema':
      return toolJson(excelSchemaDescription());
    default:
      return null;
  }
}

function toolText(lines) {
  return { content: [{ type: 'text', text: Array.isArray(lines) ? lines.join('\n') : String(lines) }] };
}

function toolJson(value) {
  return toolText(JSON.stringify(value, null, 2));
}

function toolError(message) {
  return { content: [{ type: 'text', text: String(message) }], isError: true };
}

module.exports = {
  EXCEL_TOOLS,
  excelSchemaDescription,
  handleAppendEditRecord,
  handleAppendRow,
  handleExcelTool,
  handleReadExcelHeaders,
  handleUpdateCell,
};
