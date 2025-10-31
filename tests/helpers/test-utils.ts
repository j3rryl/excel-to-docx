import XLSX from "xlsx";
import { existsSync } from "fs";

/**
 * Create a test Excel file with sample data
 */
export function createTestExcelData(data: any[][]): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, "TestData");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

/**
 * Create a simple DOCX template for testing
 */
export function createTestTemplate(
  content: string = "Test Template {{Name}}"
): ArrayBuffer {
  // In real implementation, you'd create an actual DOCX buffer
  // For testing, we can use a simple text buffer
  return new TextEncoder().encode(content).buffer;
}

/**
 * Wait for file to be created (useful for async file operations)
 */
export async function waitForFile(
  path: string,
  timeout: number = 5000
): Promise<boolean> {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    if (existsSync(path)) {
      return true;
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  return false;
}
