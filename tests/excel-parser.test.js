const assert = require('node:assert/strict');
const test = require('node:test');
const ExcelJS = require('exceljs');
const { parseExcelBuffer } = require('../dist/excel-parser.js');

const defaultLimits = {
  maxRows: 5000,
  maxWorksheetRows: 100000,
  maxColumns: 256,
  autoDetectHeader: true,
  headerScanMaxRows: 40,
  headerHintKeywords: [],
};

async function workbookBuffer(build) {
  const workbook = new ExcelJS.Workbook();
  build(workbook);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

test('uses explicit rows, preserves columns after Z, and counts blank rows', async () => {
  const buffer = await workbookBuffer((workbook) => {
    const sheet = workbook.addWorksheet('Employees');
    sheet.getCell('A1').value = 'Employee report';
    for (let column = 1; column <= 27; column += 1) {
      sheet.getRow(2).getCell(column).value = `Header ${column}`;
    }
    sheet.getCell('A3').value = '  NV001  ';
    sheet.getCell('B3').value = 0;
    sheet.getCell('AA3').value = 'after-z';
    sheet.getCell('A5').value = 'NV002';
  });

  const result = await parseExcelBuffer(
    buffer,
    'employees.xlsx',
    { sheetName: 'employees', headerRow: 2, dataStartRow: 3 },
    defaultLimits
  );

  assert.equal(result.sheetName, 'Employees');
  assert.equal(result.sheetFallbackUsed, false);
  assert.equal(result.headerAutoDetected, false);
  assert.equal(result.parsedRows, 2);
  assert.equal(result.skippedRows, 1);
  assert.equal(result.rows[0].rowNumber, 3);
  assert.equal(result.rows[0].values.length, 27);
  assert.equal(result.rows[0].values[0], 'NV001');
  assert.equal(result.rows[0].values[1], '0');
  assert.equal(result.rows[0].values[26], 'after-z');
});

test('auto-detects a header after a report title', async () => {
  const buffer = await workbookBuffer((workbook) => {
    const sheet = workbook.addWorksheet('Data');
    sheet.addRow(['Employee report: September']);
    sheet.addRow([]);
    sheet.addRow(['Employee code', 'Full name', 'Birthday']);
    sheet.addRow(['NV001', 'Nguyen Van A', '2000-01-02']);
  });

  const result = await parseExcelBuffer(buffer, 'employees.xlsx', {}, defaultLimits);
  assert.equal(result.headerRow, 3);
  assert.equal(result.dataStartRow, 4);
  assert.equal(result.headerAutoDetected, true);
  assert.equal(result.dataStartAutoDetected, true);
  assert.deepEqual(result.headers, ['Employee code', 'Full name', 'Birthday']);
});

test('rejects worksheets over the configured column limit', async () => {
  const buffer = await workbookBuffer((workbook) => {
    const sheet = workbook.addWorksheet('Data');
    sheet.getCell('A1').value = 'A';
    sheet.getCell('C1').value = 'C';
  });

  await assert.rejects(
    parseExcelBuffer(buffer, 'wide.xlsx', { headerRow: 1 }, { ...defaultLimits, maxColumns: 2 }),
    /Column limit exceeded/
  );
});

test('rejects input over the configured non-empty row limit', async () => {
  const buffer = await workbookBuffer((workbook) => {
    const sheet = workbook.addWorksheet('Data');
    sheet.addRow(['Code', 'Name']);
    sheet.addRow(['A', 'First']);
    sheet.addRow(['B', 'Second']);
  });

  await assert.rejects(
    parseExcelBuffer(buffer, 'rows.xlsx', { headerRow: 1, dataStartRow: 2 }, { ...defaultLimits, maxRows: 1 }),
    /Row limit exceeded: 1/
  );
});

test('rejects a sparse worksheet with an excessive last row position', async () => {
  const buffer = await workbookBuffer((workbook) => {
    const sheet = workbook.addWorksheet('Data');
    sheet.addRow(['Code', 'Name']);
    sheet.getCell('A100').value = 'late-row';
  });

  await assert.rejects(
    parseExcelBuffer(
      buffer,
      'sparse.xlsx',
      { headerRow: 1, dataStartRow: 2 },
      { ...defaultLimits, maxWorksheetRows: 50 }
    ),
    /Worksheet row-position limit exceeded/
  );
});
