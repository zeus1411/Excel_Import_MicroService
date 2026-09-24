const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const { createServer } = require('node:http');
const test = require('node:test');
const path = require('node:path');
const ExcelJS = require('exceljs');

function listen(server) {
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolve(server.address().port));
  });
}

function close(server) {
  return new Promise((resolve) => server.close(() => resolve()));
}

async function getFreePort() {
  const server = createServer();
  const port = await listen(server);
  await close(server);
  return port;
}

async function waitForHealth(url, child) {
  for (let attempt = 0; attempt < 50; attempt += 1) {
    if (child.exitCode !== null) throw new Error(`Microservice exited with code ${child.exitCode}`);
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // The child process may still be starting.
    }
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error('Timed out waiting for microservice health endpoint');
}

test('HTTP API authenticates, downloads an allowlisted source, and returns normalized rows', async (context) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.addRow(['Code', 'Name']);
  sheet.addRow(['NV001', 'Nguyen Van A']);
  const excelBuffer = Buffer.from(await workbook.xlsx.writeBuffer());

  const sourceServer = createServer((request, response) => {
    if (request.url !== '/employees.xlsx') {
      response.writeHead(404).end();
      return;
    }
    response.writeHead(200, {
      'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'content-length': excelBuffer.length,
    });
    response.end(excelBuffer);
  });
  const sourcePort = await listen(sourceServer);
  context.after(() => close(sourceServer));

  const servicePort = await getFreePort();
  const child = spawn(process.execPath, [path.resolve(__dirname, '../dist/server.js')], {
    env: {
      ...process.env,
      HOST: '127.0.0.1',
      PORT: String(servicePort),
      EXCEL_SERVICE_API_KEY: 'test-api-key',
      EXCEL_SERVICE_PROJECT_KEYS_JSON: JSON.stringify({ 'test-project': 'test-api-key' }),
      EXCEL_SOURCE_ALLOWED_HOSTS: '127.0.0.1',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  context.after(() => {
    if (child.exitCode === null) child.kill('SIGTERM');
  });

  await waitForHealth(`http://127.0.0.1:${servicePort}/health`, child);
  const requestBody = {
    source: {
      type: 'presigned-url',
      url: `http://127.0.0.1:${sourcePort}/employees.xlsx`,
      fileName: 'employees.xlsx',
    },
    options: { headerRow: 1, dataStartRow: 2 },
    clientReference: { projectCode: 'test-project', sourceFileId: 93 },
  };

  const unauthorized = await fetch(`http://127.0.0.1:${servicePort}/v1/excel/parse`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(requestBody),
  });
  assert.equal(unauthorized.status, 401);

  const wrongProject = await fetch(`http://127.0.0.1:${servicePort}/v1/excel/parse`, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-api-key': 'test-api-key' },
    body: JSON.stringify({
      ...requestBody,
      clientReference: { projectCode: 'another-project', sourceFileId: 93 },
    }),
  });
  assert.equal(wrongProject.status, 403);

  const response = await fetch(`http://127.0.0.1:${servicePort}/v1/excel/parse`, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-api-key': 'test-api-key' },
    body: JSON.stringify(requestBody),
  });
  const body = await response.json();
  assert.equal(response.status, 200);
  assert.equal(body.success, true);
  assert.equal(body.data.sheetName, 'Data');
  assert.deepEqual(body.data.headers, ['Code', 'Name']);
  assert.deepEqual(body.data.rows, [{ rowNumber: 2, values: ['NV001', 'Nguyen Van A'] }]);
});
