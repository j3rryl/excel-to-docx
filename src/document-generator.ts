import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import type { ExcelRecord } from "./types.js";
import { generateFilename } from "./utils.js";

/**
 * Generate a single document from template and record data
 */
export function generateDocument(
  templateBuffer: ArrayBuffer,
  record: ExcelRecord
): Buffer {
  try {
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    doc.render(record);

    return doc.getZip().generate({ type: "nodebuffer" });
  } catch (error) {
    throw new Error(
      `Document generation failed: ${
        error instanceof Error ? error.message : "Unknown error"
      }`
    );
  }
}

/**
 * Process a single record and generate document
 */
export async function processRecord(
  record: ExcelRecord,
  index: number,
  templateBuffer: ArrayBuffer,
  outputDir: string,
  fileNameTemplate: string,
  cleanFileName: boolean,
  verbose: boolean
): Promise<{ success: boolean; filePath?: string; error?: string }> {
  try {
    // Generate document
    const documentBuffer = generateDocument(templateBuffer, record);

    // Generate filename
    const filename = generateFilename(record, fileNameTemplate, cleanFileName);
    const filePath = `${outputDir}/${filename}`;

    // Write file
    await Bun.write(filePath, documentBuffer);

    if (verbose) {
      console.log(`Generated: ${filename}`);
    }

    return { success: true, filePath };
  } catch (error) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    if (verbose) {
      console.error(`Error processing record ${index + 1}:`, errorMsg);
    }
    return { success: false, error: errorMsg };
  }
}

/**
 * Generate multiple documents from records
 */
export async function generateMultipleDocuments(
  records: ExcelRecord[],
  templateBuffer: ArrayBuffer,
  outputDir: string,
  fileNameTemplate: string,
  cleanFileName: boolean,
  verbose: boolean
): Promise<{
  generatedFiles: string[];
  errors: Array<{ record: number; error: string }>;
}> {
  const generatedFiles: string[] = [];
  const errors: Array<{ record: number; error: string }> = [];

  for (const [index, record] of records.entries()) {
    const result = await processRecord(
      record,
      index,
      templateBuffer,
      outputDir,
      fileNameTemplate,
      cleanFileName,
      verbose
    );

    if (result.success && result.filePath) {
      generatedFiles.push(result.filePath);
    } else {
      errors.push({
        record: index + 1,
        error: result.error || "Unknown error",
      });
    }
  }

  return { generatedFiles, errors };
}
