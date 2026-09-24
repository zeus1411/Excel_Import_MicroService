"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.loadConfig = loadConfig;
function getPositiveInt(name, defaultValue) {
    const raw = process.env[name];
    if (!raw)
        return defaultValue;
    const parsed = Number(raw);
    if (!Number.isFinite(parsed) || parsed < 1)
        return defaultValue;
    return Math.floor(parsed);
}
function getBool(name, defaultValue) {
    const raw = process.env[name];
    if (!raw)
        return defaultValue;
    const normalized = raw.trim().toLowerCase();
    if (['1', 'true', 'yes', 'y', 'on'].includes(normalized))
        return true;
    if (['0', 'false', 'no', 'n', 'off'].includes(normalized))
        return false;
    return defaultValue;
}
function normalizeForDetection(input) {
    return String(input || '')
        .replace(/[đĐ]/g, (character) => (character === 'đ' ? 'd' : 'D'))
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim()
        .replace(/\s+/g, ' ');
}
function getHeaderHintKeywords() {
    const raw = process.env.EXCEL_SERVICE_HEADER_HINT_KEYWORDS || '';
    const values = raw
        .split(/[;,\n|]/g)
        .map(normalizeForDetection)
        .filter(Boolean);
    return Array.from(new Set(values));
}
function getAllowedSourceHosts() {
    return String(process.env.EXCEL_SOURCE_ALLOWED_HOSTS || '')
        .split(',')
        .map((value) => value.trim().toLowerCase())
        .filter(Boolean);
}
function getCredentials() {
    const projectKeysJson = process.env.EXCEL_SERVICE_PROJECT_KEYS_JSON?.trim();
    if (projectKeysJson) {
        let parsed;
        try {
            parsed = JSON.parse(projectKeysJson);
        }
        catch {
            throw new Error('EXCEL_SERVICE_PROJECT_KEYS_JSON must be valid JSON');
        }
        if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
            throw new Error('EXCEL_SERVICE_PROJECT_KEYS_JSON must be an object of projectCode -> API key');
        }
        const credentials = Object.entries(parsed)
            .map(([projectCode, apiKey]) => ({ projectCode: projectCode.trim(), apiKey: String(apiKey).trim() }))
            .filter((item) => item.projectCode && item.apiKey);
        if (!credentials.length)
            throw new Error('EXCEL_SERVICE_PROJECT_KEYS_JSON contains no usable credentials');
        return credentials;
    }
    const apiKey = process.env.EXCEL_SERVICE_API_KEY?.trim();
    if (!apiKey) {
        throw new Error('Missing EXCEL_SERVICE_PROJECT_KEYS_JSON or EXCEL_SERVICE_API_KEY');
    }
    return [{ projectCode: null, apiKey }];
}
function loadConfig() {
    return {
        host: process.env.HOST?.trim() || '0.0.0.0',
        port: getPositiveInt('PORT', 3100),
        credentials: getCredentials(),
        allowedSourceHosts: getAllowedSourceHosts(),
        requestBodyMaxBytes: getPositiveInt('EXCEL_SERVICE_REQUEST_MAX_BYTES', 1024 * 1024),
        sourceFileMaxBytes: getPositiveInt('EXCEL_SERVICE_FILE_MAX_BYTES', 25 * 1024 * 1024),
        sourceDownloadTimeoutMs: getPositiveInt('EXCEL_SERVICE_DOWNLOAD_TIMEOUT_MS', 30_000),
        parser: {
            maxRows: getPositiveInt('EXCEL_SERVICE_MAX_ROWS', 5000),
            maxWorksheetRows: getPositiveInt('EXCEL_SERVICE_MAX_WORKSHEET_ROWS', 100_000),
            maxColumns: getPositiveInt('EXCEL_SERVICE_MAX_COLUMNS', 256),
            autoDetectHeader: getBool('EXCEL_SERVICE_AUTO_DETECT_HEADER', true),
            headerScanMaxRows: getPositiveInt('EXCEL_SERVICE_HEADER_SCAN_MAX_ROWS', 40),
            headerHintKeywords: getHeaderHintKeywords(),
        },
    };
}
