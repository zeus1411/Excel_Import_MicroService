export type ParseExcelOptions = {
  sheetName?: string;
  headerRow?: number;
  dataStartRow?: number;
};

export type ParseExcelRequest = {
  source: {
    type: 'presigned-url';
    url: string;
    fileName: string;
    sha256?: string;
  };
  options?: ParseExcelOptions;
  clientReference?: {
    projectCode?: string;
    sourceFileId?: string | number;
  };
};

export type ParsedExcelRow = {
  rowNumber: number;
  values: Array<string | null>;
};

export type ParsedExcelResult = {
  fileName: string;
  sheetName: string;
  availableSheetNames: string[];
  sheetFallbackUsed: boolean;
  headerRow: number;
  dataStartRow: number;
  headerAutoDetected: boolean;
  dataStartAutoDetected: boolean;
  headers: string[];
  rows: ParsedExcelRow[];
  parsedRows: number;
  skippedRows: number;
};

export type ParserLimits = {
  maxRows: number;
  maxWorksheetRows: number;
  maxColumns: number;
  autoDetectHeader: boolean;
  headerScanMaxRows: number;
  headerHintKeywords: string[];
};
