import * as XLSX from 'xlsx';

const TEST_SUITE_COLUMNS = [
  { key: 'protocol', header: 'Protocol' },
  { key: 'serverPath', header: 'Server Path' },
  { key: 'path', header: 'Path' },
  { key: 'method', header: 'Method' },
  { key: 'params', header: 'Params' },
  { key: 'urlTemplate', header: 'URL' },
  { key: 'paramValues', header: 'Param Values' },
  { key: 'headerValues', header: 'Headers' },
  { key: 'body', header: 'Body' },
  { key: 'responseCode', header: 'Response Code' },
  { key: 'responseHeaders', header: 'Response Headers' },
  { key: 'responseBody', header: 'Response Body' },
  { key: 'responseTimeMs', header: 'Response Time (ms)' },
];

function normalizeCell(value) {
  if (value === undefined || value === null) {
    return '';
  }

  return String(value);
}

export function exportTestSuiteToExcel(testSuiteRows) {
  const rows = testSuiteRows.map((row) => {
    const mapped = {};

    TEST_SUITE_COLUMNS.forEach((column) => {
      mapped[column.header] = normalizeCell(row[column.key]);
    });

    return mapped;
  });

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows, {
    header: TEST_SUITE_COLUMNS.map((column) => column.header),
  });

  XLSX.utils.book_append_sheet(workbook, worksheet, 'TestSuite');
  XLSX.writeFile(workbook, 'SwiftAPI_TestSuite.xlsx');
}

export async function importTestSuiteFromExcel(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });

  const targetSheetName = workbook.SheetNames.includes('TestSuite')
    ? 'TestSuite'
    : workbook.SheetNames[0];

  if (!targetSheetName) {
    return [];
  }

  const worksheet = workbook.Sheets[targetSheetName];
  const sheetRows = XLSX.utils.sheet_to_json(worksheet, {
    defval: '',
  });

  return sheetRows.map((row, index) => {
    const normalized = {
      id: `imported:${index}`,
      protocol: '',
      serverPath: '',
      path: '',
      method: 'GET',
      params: '',
      urlTemplate: '',
      paramValues: '',
      headerValues: '',
      body: '',
      responseCode: '',
      responseHeaders: '',
      responseBody: '',
      responseTimeMs: '',
    };

    TEST_SUITE_COLUMNS.forEach((column) => {
      normalized[column.key] = normalizeCell(row[column.header]);
    });

    return normalized;
  });
}
