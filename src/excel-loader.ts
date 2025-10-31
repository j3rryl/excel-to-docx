import XLSX from "xlsx";
import type { ExcelRecord } from "./types.js";

/**
 * Load and validate Excel workbook
 */
export function loadWorkbook(dataBuffer: ArrayBuffer): XLSX.WorkBook {
  try {
    return XLSX.read(dataBuffer, {
      type: "buffer",
      cellDates: true,
    });
  } catch (readError) {
    throw new Error(
      `Failed to read Excel file: ${
        readError instanceof Error ? readError.message : "File may be corrupted"
      }`
    );
  }
}

/**
 * Get first sheet from workbook with validation
 */
export function getFirstSheet(workbook: XLSX.WorkBook): XLSX.WorkSheet {
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error("Excel file contains no sheets");
  }

  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("First worksheet has an invalid name");
  }

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`Worksheet "${sheetName}" is empty or corrupted`);
  }

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in Excel file`);
  }

  // Check if sheet has data
  const range = sheet["!ref"];
  if (!range) {
    throw new Error(`Sheet "${sheetName}" contains no data`);
  }

  return sheet;
}

/**
 * Parse sheet data into records with validation
 */
export function parseSheetData(sheet: XLSX.WorkSheet): ExcelRecord[] {
  let records: ExcelRecord[];

  try {
    records = XLSX.utils.sheet_to_json(sheet, {
      raw: false,
      defval: "",
      blankrows: false,
    });
  } catch (parseError) {
    throw new Error(
      `Failed to parse Excel data: ${
        parseError instanceof Error
          ? parseError.message
          : "Unknown parsing error"
      }`
    );
  }

  if (!Array.isArray(records)) {
    throw new Error("Excel data could not be parsed into records array");
  }

  if (records.length === 0) {
    throw new Error("No data records found in Excel sheet");
  }

  // Filter out completely empty records
  const validRecords = records.filter(
    (record) =>
      record &&
      typeof record === "object" &&
      Object.keys(record).length > 0 &&
      Object.values(record).some((value) => value !== "" && value != null)
  );

  if (validRecords.length === 0) {
    throw new Error("All records in Excel file are empty");
  }

  return validRecords;
}

/**
 * Load and validate Excel data from buffer
 */
export function loadExcelData(
  dataBuffer: ArrayBuffer,
  verbose: boolean = false
): ExcelRecord[] {
  const workbook = loadWorkbook(dataBuffer);
  const sheet = getFirstSheet(workbook);
  const records = parseSheetData(sheet);

  if (verbose) {
    const sampleRecord = records[0] as ExcelRecord;
    const fields = Object.keys(sampleRecord).filter(
      (key) => sampleRecord[key] !== "" && sampleRecord[key] != null
    );
    console.log(
      `Loaded ${records.length} records with fields: ${fields.join(", ")}`
    );
  }

  return records;
}
