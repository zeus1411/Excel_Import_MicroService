"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseExcelBuffer = parseExcelBuffer;
const exceljs_1 = __importDefault(require("exceljs"));
function parseOptionalPositiveInt(input) {
    if (input === null || input === undefined || input === '')
        return undefined;
    const parsed = Number(input);
    if (!Number.isFinite(parsed) || parsed < 1)
        return undefined;
    return Math.floor(parsed);
}
function cellToText(cell) {
    const value = cell.value;
    if (value === null || value === undefined)
        return null;
    if (typeof value === 'string') {
        const trimmed = value.trim();
        return trimmed || null;
    }
    if (typeof value === 'number' || typeof value === 'boolean')
        return String(value);
    if (value instanceof Date)
        return value.toISOString();
    const text = String(cell.text || '').trim();
    return text || null;
}
function toAsciiLabel(input) {
    return String(input)
        .replace(/[đĐ]/g, (character) => (character === 'đ' ? 'd' : 'D'))
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');
}
function normalizeForDetection(input) {
    return toAsciiLabel(String(input || ''))
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim()
        .replace(/\s+/g, ' ');
}
function isNumericLike(input) {
    const value = String(input || '').trim();
    if (!value)
        return false;
    if (/^[+-]?\d+(?:[.,]\d+)?$/.test(value))
        return true;
    if (/^\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}$/.test(value))
        return true;
    return /^\d{4}[/.\-]\d{1,2}[/.\-]\d{1,2}$/.test(value);
}
function getRowValues(worksheet, rowNumber, maxCol) {
    const row = worksheet.getRow(rowNumber);
    const values = [];
    for (let column = 1; column <= maxCol; column += 1) {
        values.push(cellToText(row.getCell(column)));
    }
    return values;
}
function getLastNonEmptyColumn(values) {
    for (let index = values.length - 1; index >= 0; index -= 1) {
        if (values[index] !== null)
            return index + 1;
    }
    return 0;
}
function countHeaderKeywordHits(normalizedInput, keywords) {
    let hits = 0;
    for (const keyword of keywords) {
        if (keyword.length <= 3) {
            if (normalizedInput === keyword || normalizedInput.split(' ').includes(keyword))
                hits += 1;
        }
        else if (normalizedInput.includes(keyword)) {
            hits += 1;
        }
    }
    return hits;
}
function analyzeRow(values, keywords) {
    let nonEmpty = 0;
    let keywordHits = 0;
    let colonCells = 0;
    let longCells = 0;
    let numericLikeCells = 0;
    const unique = new Set();
    for (const value of values) {
        if (value === null)
            continue;
        const text = value.trim();
        if (!text)
            continue;
        nonEmpty += 1;
        const normalized = normalizeForDetection(text);
        if (normalized) {
            unique.add(normalized);
            keywordHits += countHeaderKeywordHits(normalized, keywords);
        }
        if (text.includes(':'))
            colonCells += 1;
        if (text.length > 40)
            longCells += 1;
        if (isNumericLike(text))
            numericLikeCells += 1;
    }
    return { nonEmpty, uniqueNonEmpty: unique.size, keywordHits, colonCells, longCells, numericLikeCells };
}
function scoreHeaderRow(current, next) {
    if (current.uniqueNonEmpty < 2 || current.nonEmpty < 2)
        return -9999;
    let score = current.uniqueNonEmpty * 3 + Math.min(current.nonEmpty, 10) + current.keywordHits * 2;
    score += Math.round((current.uniqueNonEmpty / current.nonEmpty) * 4);
    score -= current.colonCells * 4;
    score -= current.longCells * 2;
    score -= current.numericLikeCells * 2;
    if (next) {
        if (next.nonEmpty > 0)
            score += 1;
        if (next.numericLikeCells > 0)
            score += 2;
        if (next.keywordHits < current.keywordHits)
            score += 1;
    }
    return score;
}
function detectDataStartRow(worksheet, headerRow, lastCol, keywords, scanWindow = 30) {
    const end = Math.min(worksheet.rowCount || headerRow + 1, headerRow + Math.max(scanWindow, 1));
    for (let rowNumber = headerRow + 1; rowNumber <= end; rowNumber += 1) {
        const signal = analyzeRow(getRowValues(worksheet, rowNumber, lastCol), keywords);
        if (signal.nonEmpty === 0)
            continue;
        if (signal.uniqueNonEmpty === 1 && signal.nonEmpty > 1 && signal.colonCells > 0)
            continue;
        return rowNumber;
    }
    return headerRow + 1;
}
function detectHeaderAndDataStartRow(worksheet, maxCol, scanMaxRows, keywords) {
    const endRow = Math.min(Math.max(worksheet.rowCount || 1, 1), Math.max(scanMaxRows, 1));
    let bestHeaderRow = 1;
    let bestLastCol = 0;
    let bestScore = Number.NEGATIVE_INFINITY;
    for (let rowNumber = 1; rowNumber <= endRow; rowNumber += 1) {
        const currentValues = getRowValues(worksheet, rowNumber, maxCol);
        const currentSignal = analyzeRow(currentValues, keywords);
        if (currentSignal.nonEmpty === 0)
            continue;
        const currentLastCol = getLastNonEmptyColumn(currentValues);
        if (currentLastCol === 0)
            continue;
        const nextSignal = rowNumber < endRow
            ? analyzeRow(getRowValues(worksheet, rowNumber + 1, maxCol), keywords)
            : null;
        const score = scoreHeaderRow(currentSignal, nextSignal);
        if (score > bestScore) {
            bestScore = score;
            bestHeaderRow = rowNumber;
            bestLastCol = currentLastCol;
        }
    }
    if (bestLastCol === 0)
        throw new Error('Cannot auto-detect header row: worksheet has no non-empty rows');
    return {
        headerRow: bestHeaderRow,
        dataStartRow: detectDataStartRow(worksheet, bestHeaderRow, bestLastCol, keywords),
        lastCol: bestLastCol,
    };
}
function normalizeHeaderLabel(raw, index) {
    const cleaned = String(raw ?? '')
        .replace(/\u0000/g, '')
        .replace(/\r\n|\r|\n/g, ' ')
        .trim();
    return cleaned || `COL_${index}`;
}
function ensureValidDataStartRow(headerRow, requested) {
    if (!requested)
        return undefined;
    return requested <= headerRow ? headerRow + 1 : requested;
}
async function parseExcelBuffer(buffer, fileName, options, limits) {
    if (!buffer.length)
        throw new Error('Excel source is empty');
    const workbook = new exceljs_1.default.Workbook();
    await workbook.xlsx.load(buffer);
    const availableSheetNames = workbook.worksheets.map((worksheet) => worksheet.name);
    const requestedSheet = typeof options.sheetName === 'string' ? options.sheetName.trim() : '';
    let worksheet;
    let sheetFallbackUsed = false;
    if (requestedSheet) {
        worksheet = workbook.getWorksheet(requestedSheet);
        if (!worksheet) {
            const requestedLower = requestedSheet.toLowerCase();
            worksheet = workbook.worksheets.find((candidate) => candidate.name.toLowerCase() === requestedLower);
        }
        if (!worksheet && workbook.worksheets.length > 0) {
            worksheet = workbook.worksheets[0];
            sheetFallbackUsed = true;
        }
    }
    else {
        worksheet = workbook.worksheets[0];
    }
    if (!worksheet) {
        throw new Error(`Worksheet not found. Requested=${requestedSheet || '(first sheet)'}, available=${JSON.stringify(availableSheetNames)}`);
    }
    if (worksheet.rowCount > limits.maxWorksheetRows) {
        throw new Error(`Worksheet row-position limit exceeded: ${worksheet.rowCount} > ${limits.maxWorksheetRows}`);
    }
    // columnCount/rowCount preserve sparse positions (for example data in AA or row 5
    // with blank columns/rows in between); actual*Count only counts populated entries.
    const sheetColumnCount = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0, 1);
    if (sheetColumnCount > limits.maxColumns) {
        throw new Error(`Column limit exceeded: ${sheetColumnCount} > ${limits.maxColumns}`);
    }
    const explicitHeaderRow = parseOptionalPositiveInt(options.headerRow);
    const explicitDataStartRow = parseOptionalPositiveInt(options.dataStartRow);
    let headerRow = 1;
    let dataStartRow = 2;
    let lastCol = 0;
    let headerAutoDetected = false;
    let dataStartAutoDetected = false;
    if (explicitHeaderRow) {
        headerRow = explicitHeaderRow;
        lastCol = getLastNonEmptyColumn(getRowValues(worksheet, headerRow, sheetColumnCount));
        if (lastCol === 0)
            throw new Error(`Header row ${headerRow} is empty`);
        dataStartRow = ensureValidDataStartRow(headerRow, explicitDataStartRow)
            || detectDataStartRow(worksheet, headerRow, lastCol, limits.headerHintKeywords);
        dataStartAutoDetected = !explicitDataStartRow;
    }
    else if (limits.autoDetectHeader) {
        const detected = detectHeaderAndDataStartRow(worksheet, sheetColumnCount, limits.headerScanMaxRows, limits.headerHintKeywords);
        headerRow = detected.headerRow;
        lastCol = detected.lastCol;
        dataStartRow = ensureValidDataStartRow(headerRow, explicitDataStartRow) || detected.dataStartRow;
        headerAutoDetected = true;
        dataStartAutoDetected = !explicitDataStartRow;
    }
    else {
        lastCol = getLastNonEmptyColumn(getRowValues(worksheet, 1, sheetColumnCount));
        if (lastCol === 0)
            throw new Error('Header row is empty');
        dataStartRow = ensureValidDataStartRow(headerRow, explicitDataStartRow)
            || detectDataStartRow(worksheet, headerRow, lastCol, limits.headerHintKeywords);
        dataStartAutoDetected = !explicitDataStartRow;
    }
    const headers = getRowValues(worksheet, headerRow, sheetColumnCount)
        .map((value, index) => normalizeHeaderLabel(value || '', index + 1));
    const rows = [];
    let skippedRows = 0;
    const endRow = worksheet.rowCount || dataStartRow;
    for (let rowNumber = dataStartRow; rowNumber <= endRow; rowNumber += 1) {
        const values = getRowValues(worksheet, rowNumber, sheetColumnCount);
        if (!values.some((value) => value !== null)) {
            skippedRows += 1;
            continue;
        }
        if (rows.length >= limits.maxRows) {
            throw new Error(`Row limit exceeded: ${limits.maxRows}`);
        }
        rows.push({ rowNumber, values });
    }
    return {
        fileName,
        sheetName: worksheet.name,
        availableSheetNames,
        sheetFallbackUsed,
        headerRow,
        dataStartRow,
        headerAutoDetected,
        dataStartAutoDetected,
        headers,
        rows,
        parsedRows: rows.length,
        skippedRows,
    };
}
