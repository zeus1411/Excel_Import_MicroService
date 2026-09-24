"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const crypto_1 = require("crypto");
const http_1 = require("http");
const config_1 = require("./config");
const excel_parser_1 = require("./excel-parser");
class HttpError extends Error {
    status;
    constructor(status, message) {
        super(message);
        this.status = status;
    }
}
function sendJson(response, status, payload) {
    response.writeHead(status, {
        'content-type': 'application/json; charset=utf-8',
        'cache-control': 'no-store',
    });
    response.end(JSON.stringify(payload));
}
function secretMatches(received, expected) {
    if (!received)
        return false;
    const receivedBuffer = Buffer.from(received);
    const expectedBuffer = Buffer.from(expected);
    return receivedBuffer.length === expectedBuffer.length && (0, crypto_1.timingSafeEqual)(receivedBuffer, expectedBuffer);
}
function authenticateProject(received) {
    for (const credential of config.credentials) {
        if (secretMatches(received, credential.apiKey))
            return credential.projectCode;
    }
    return undefined;
}
async function readJsonBody(request, maxBytes) {
    const chunks = [];
    let total = 0;
    for await (const chunk of request) {
        const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
        total += buffer.length;
        if (total > maxBytes)
            throw new HttpError(413, `Request body exceeds ${maxBytes} bytes`);
        chunks.push(buffer);
    }
    if (total === 0)
        throw new HttpError(400, 'JSON request body is required');
    try {
        return JSON.parse(Buffer.concat(chunks).toString('utf8'));
    }
    catch {
        throw new HttpError(400, 'Request body is not valid JSON');
    }
}
function validateRequest(input) {
    if (!input || typeof input !== 'object')
        throw new HttpError(400, 'Request body must be an object');
    const request = input;
    if (!request.source || request.source.type !== 'presigned-url') {
        throw new HttpError(400, 'source.type must be presigned-url');
    }
    if (typeof request.source.url !== 'string' || !request.source.url.trim()) {
        throw new HttpError(400, 'source.url is required');
    }
    if (typeof request.source.fileName !== 'string' || !request.source.fileName.trim()) {
        throw new HttpError(400, 'source.fileName is required');
    }
    if (!request.source.fileName.toLowerCase().endsWith('.xlsx')) {
        throw new HttpError(415, 'Only .xlsx files are supported');
    }
    return request;
}
function validateSourceUrl(rawUrl, allowedHosts) {
    let url;
    try {
        url = new URL(rawUrl);
    }
    catch {
        throw new HttpError(400, 'source.url is invalid');
    }
    if (!['http:', 'https:'].includes(url.protocol))
        throw new HttpError(400, 'source.url must use HTTP or HTTPS');
    if (url.username || url.password)
        throw new HttpError(400, 'source.url must not contain URL credentials');
    if (allowedHosts.length === 0) {
        throw new HttpError(503, 'EXCEL_SOURCE_ALLOWED_HOSTS is not configured');
    }
    const hostAllowed = allowedHosts.includes('*')
        || allowedHosts.includes(url.hostname.toLowerCase())
        || allowedHosts.includes(url.host.toLowerCase());
    if (!hostAllowed)
        throw new HttpError(403, `Source host '${url.host}' is not allowed`);
    return url;
}
async function downloadSource(url, maxBytes, timeoutMs, expectedSha256) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
        const response = await fetch(url, { signal: controller.signal, redirect: 'manual' });
        if (response.status >= 300 && response.status < 400) {
            throw new HttpError(400, 'Source URL redirects are not allowed');
        }
        if (!response.ok)
            throw new HttpError(502, `Source download failed with HTTP ${response.status}`);
        const contentLength = Number(response.headers.get('content-length'));
        if (Number.isFinite(contentLength) && contentLength > maxBytes) {
            throw new HttpError(413, `Source file exceeds ${maxBytes} bytes`);
        }
        if (!response.body)
            throw new HttpError(502, 'Source response has no body');
        const chunks = [];
        let total = 0;
        for await (const chunk of response.body) {
            const buffer = Buffer.from(chunk);
            total += buffer.length;
            if (total > maxBytes)
                throw new HttpError(413, `Source file exceeds ${maxBytes} bytes`);
            chunks.push(buffer);
        }
        const result = Buffer.concat(chunks);
        if (!result.length)
            throw new HttpError(400, 'Downloaded source file is empty');
        if (expectedSha256) {
            const expected = expectedSha256.trim().toLowerCase();
            if (!/^[a-f0-9]{64}$/.test(expected))
                throw new HttpError(400, 'source.sha256 must be 64 hex characters');
            const actual = (0, crypto_1.createHash)('sha256').update(result).digest('hex');
            if (actual !== expected)
                throw new HttpError(400, 'Downloaded source SHA-256 does not match');
        }
        return result;
    }
    catch (error) {
        if (error instanceof HttpError)
            throw error;
        if (error instanceof Error && error.name === 'AbortError')
            throw new HttpError(504, 'Source download timed out');
        throw new HttpError(502, `Source download failed: ${error instanceof Error ? error.message : String(error)}`);
    }
    finally {
        clearTimeout(timer);
    }
}
const config = (0, config_1.loadConfig)();
const server = (0, http_1.createServer)(async (request, response) => {
    const requestId = (0, crypto_1.randomUUID)();
    try {
        const url = new URL(request.url || '/', 'http://service.local');
        if (request.method === 'GET' && url.pathname === '/health') {
            return sendJson(response, 200, { status: 'ok', service: 'excel-import-service' });
        }
        if (request.method !== 'POST' || url.pathname !== '/v1/excel/parse') {
            throw new HttpError(404, 'Route not found');
        }
        const authenticatedProject = authenticateProject(request.headers['x-api-key']);
        if (authenticatedProject === undefined) {
            throw new HttpError(401, 'Invalid API key');
        }
        if (!String(request.headers['content-type'] || '').toLowerCase().includes('application/json')) {
            throw new HttpError(415, 'Content-Type must be application/json');
        }
        const body = validateRequest(await readJsonBody(request, config.requestBodyMaxBytes));
        if (authenticatedProject !== null && body.clientReference?.projectCode !== authenticatedProject) {
            throw new HttpError(403, 'API key is not authorized for clientReference.projectCode');
        }
        const sourceUrl = validateSourceUrl(body.source.url, config.allowedSourceHosts);
        const buffer = await downloadSource(sourceUrl, config.sourceFileMaxBytes, config.sourceDownloadTimeoutMs, body.source.sha256);
        const result = await (0, excel_parser_1.parseExcelBuffer)(buffer, body.source.fileName.trim(), body.options || {}, config.parser);
        console.info(JSON.stringify({
            level: 'info',
            requestId,
            event: 'excel_parsed',
            fileName: result.fileName,
            sheetName: result.sheetName,
            parsedRows: result.parsedRows,
            skippedRows: result.skippedRows,
            projectCode: authenticatedProject || body.clientReference?.projectCode || null,
        }));
        return sendJson(response, 200, { success: true, requestId, data: result });
    }
    catch (error) {
        const status = error instanceof HttpError ? error.status : 422;
        const message = error instanceof Error ? error.message : 'Excel parsing failed';
        console.error(JSON.stringify({ level: 'error', requestId, event: 'request_failed', status, message }));
        return sendJson(response, status, { success: false, requestId, error: { message } });
    }
});
server.requestTimeout = Math.max(config.sourceDownloadTimeoutMs + 5000, 10_000);
server.listen(config.port, config.host, () => {
    console.info(JSON.stringify({
        level: 'info',
        event: 'service_started',
        host: config.host,
        port: config.port,
        allowedSourceHosts: config.allowedSourceHosts,
        configuredProjects: config.credentials.map((credential) => credential.projectCode || '(legacy-single-key)'),
    }));
});
function shutdown(signal) {
    console.info(JSON.stringify({ level: 'info', event: 'service_stopping', signal }));
    server.close((error) => {
        if (error) {
            console.error(error);
            process.exitCode = 1;
        }
    });
}
process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
