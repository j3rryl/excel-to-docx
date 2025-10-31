import type {
  ExcelRecord,
  GenerateOptions,
  GenerationResult,
} from "./types.js";
import { validateFiles, ensureOutputDir } from "./utils.js";
import { loadExcelData } from "./excel-loader.js";
import { processRecord } from "./document-generator.js";

/**
 * Generate multiple DOCX documents from Excel data
 * @param excelPath - Path to Excel file
 * @param templatePath - Path to Word template file
 * @param options - Generation options
 * @returns Promise with generation results
 */
export async function generateDocuments(
  excelPath: string,
  templatePath: string,
  options: GenerateOptions = {}
): Promise<GenerationResult> {
  const {
    outputDir = "./output",
    fileNameTemplate = "{{CustomerName}}",
    cleanFileName = true,
    verbose = false,
  } = options;

  const result: GenerationResult = {
    success: true,
    generatedFiles: [],
    totalRecords: 0,
    successfulRecords: 0,
    errors: [],
  };

  try {
    // Validate inputs
    if (!excelPath || !templatePath) {
      throw new Error("Excel path and template path are required");
    }

    // Check if files exist
    await validateFiles(excelPath, templatePath);

    // Read files using Bun
    const dataBuffer = await Bun.file(excelPath).arrayBuffer();
    const templateBuffer = await Bun.file(templatePath).arrayBuffer();

    // Load Excel data
    const records = loadExcelData(dataBuffer, verbose);
    result.totalRecords = records.length;

    if (records.length === 0) {
      throw new Error("No data records found in Excel file");
    }

    // Create output directory
    await ensureOutputDir(outputDir);

    // Process each record
    for (const [index, record] of records.entries()) {
      const recordResult = await processRecord(
        record,
        index,
        templateBuffer,
        outputDir,
        fileNameTemplate,
        cleanFileName,
        verbose
      );

      if (recordResult.success && recordResult.filePath) {
        result.generatedFiles.push(recordResult.filePath);
        result.successfulRecords++;
      } else {
        result.errors.push({
          record: index + 1,
          error: recordResult.error || "Unknown error",
        });
      }
    }

    // Set overall success status
    result.success = result.errors.length === 0;

    if (verbose) {
      console.log(
        `Successfully generated ${result.successfulRecords} documents in ${outputDir}`
      );
      if (result.errors.length > 0) {
        console.log(`⚠️  ${result.errors.length} records had errors`);
      }
    }

    return result;
  } catch (error) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    result.success = false;
    result.errors.push({ record: 0, error: errorMsg });

    if (options.verbose) {
      console.error("Fatal error:", errorMsg);
    }

    return result;
  }
}

/**
 * Generate documents from buffers (useful for web applications)
 * @param excelBuffer - Excel file buffer
 * @param templateBuffer - Word template buffer
 * @param options - Generation options
 * @returns Promise with generation results
 */
export async function generateDocumentsFromBuffer(
  excelBuffer: ArrayBuffer,
  templateBuffer: ArrayBuffer,
  options: GenerateOptions = {}
): Promise<GenerationResult> {
  const {
    outputDir = "./output",
    fileNameTemplate = "{{CustomerName}}",
    cleanFileName = true,
    verbose = false,
  } = options;

  const result: GenerationResult = {
    success: true,
    generatedFiles: [],
    totalRecords: 0,
    successfulRecords: 0,
    errors: [],
  };

  try {
    // Load Excel data from buffer
    const records = loadExcelData(excelBuffer, verbose);
    result.totalRecords = records.length;

    if (records.length === 0) {
      throw new Error("No data records found in Excel buffer");
    }

    // Create output directory
    await ensureOutputDir(outputDir);

    // Process each record
    for (const [index, record] of records.entries()) {
      const recordResult = await processRecord(
        record,
        index,
        templateBuffer,
        outputDir,
        fileNameTemplate,
        cleanFileName,
        verbose
      );

      if (recordResult.success && recordResult.filePath) {
        result.generatedFiles.push(recordResult.filePath);
        result.successfulRecords++;
      } else {
        result.errors.push({
          record: index + 1,
          error: recordResult.error || "Unknown error",
        });
      }
    }

    result.success = result.errors.length === 0;

    if (verbose) {
      console.log(
        `Successfully generated ${result.successfulRecords} documents in ${outputDir}`
      );
      if (result.errors.length > 0) {
        console.log(`${result.errors.length} records had errors`);
      }
    }

    return result;
  } catch (error) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    result.success = false;
    result.errors.push({ record: 0, error: errorMsg });

    if (verbose) {
      console.error("Fatal error:", errorMsg);
    }

    return result;
  }
}

/**
 * Validate Excel and template files
 * @param excelPath - Path to Excel file
 * @param templatePath - Path to Word template file
 * @returns Promise that resolves if files are valid
 */
export { validateFiles } from "./utils.js";

/**
 * Get information about Excel file structure
 * @param excelPath - Path to Excel file
 * @returns Sheet names and record count
 */
export async function inspectExcelFile(excelPath: string): Promise<{
  sheetNames: string[];
  firstSheetName: string;
  recordCount: number;
  fields: string[];
}> {
  // Validate file exists
  await validateFiles(excelPath, excelPath);

  const buffer = await Bun.file(excelPath).arrayBuffer();

  // Load workbook to get sheet information
  const XLSX = await import("xlsx");
  const workbook = XLSX.read(buffer, { type: "buffer" });

  const sheetNames = workbook.SheetNames;
  const firstSheetName = sheetNames[0] || "Sheet1";

  // Load records to get field information
  const records = loadExcelData(buffer, false);
  const fields = records[0] ? Object.keys(records[0]) : [];

  return {
    sheetNames,
    firstSheetName,
    recordCount: records.length,
    fields,
  };
}

/**
 * Preview template data without generating documents
 * @param excelPath - Path to Excel file
 * @param templatePath - Path to Word template file
 * @returns Preview information
 */
export async function previewGeneration(
  excelPath: string,
  templatePath: string
): Promise<{
  recordCount: number;
  sampleRecord: ExcelRecord;
  fields: string[];
  outputFileNames: string[];
}> {
  await validateFiles(excelPath, templatePath);

  const dataBuffer = await Bun.file(excelPath).arrayBuffer();
  const records = loadExcelData(dataBuffer, false);

  if (records.length === 0) {
    throw new Error("No records found in Excel file");
  }

  const sampleRecord = records[0];
  if (!sampleRecord) {
    throw new Error("No sample record found in Excel file");
  }
  const fields = Object.keys(sampleRecord);

  // Generate sample output filenames for first 3 records
  const { generateFilename } = await import("./utils.js");
  const outputFileNames = records
    .slice(0, 3)
    .map((record) => generateFilename(record, "{{CustomerName}}", true));

  return {
    recordCount: records.length,
    sampleRecord,
    fields,
    outputFileNames,
  };
}

// Export individual components for advanced usage
export { loadExcelData } from "./excel-loader.js";
export { generateDocument, processRecord } from "./document-generator.js";
export { generateFilename, ensureOutputDir } from "./utils.js";

// Export types for library users
export type { ExcelRecord, GenerateOptions, GenerationResult };
