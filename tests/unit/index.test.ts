import { describe, expect, test, afterEach } from "bun:test";
import {
  generateDocuments,
  generateDocumentsFromBuffer,
  inspectExcelFile,
} from "../../src/index.js";

describe("Main Index", () => {
  afterEach(async () => {
    // Cleanup test files
    await Bun.$`rm -rf test-output output`.quiet();
  });

  describe("generateDocuments", () => {
    test("should generate documents successfully", async () => {
      const result = await generateDocuments({
        excelPath: "tests/fixtures/integration-test-data.xlsx",
        templatePath: "tests/fixtures/simple-template.docx",
        options: {
          outputDir: "test-output",
          verbose: false,
        },
      });

      expect(result.success).toBe(true);
      expect(result.totalRecords).toBe(2);
      expect(result.successfulRecords).toBe(2);
      expect(result.generatedFiles).toHaveLength(2);
      expect(result.errors).toHaveLength(0);
    });

    test("should handle file not found errors", async () => {
      const result = await generateDocuments({
        excelPath: "nonexistent.xlsx",
        templatePath: "tests/fixtures/simple-template.docx",
        options: {},
      });

      expect(result.success).toBe(false);
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0].error).toContain("not found");
    });
  });

  describe("generateDocumentsFromBuffer", () => {
    test("should generate documents from buffers", async () => {
      const excelBuffer = await Bun.file(
        "tests/fixtures/integration-test-data.xlsx"
      ).arrayBuffer();
      const templateBuffer = await Bun.file(
        "tests/fixtures/simple-template.docx"
      ).arrayBuffer();

      const result = await generateDocumentsFromBuffer({
        excelBuffer,
        templateBuffer,
        options: {
          outputDir: "test-output",
          verbose: false,
        },
      });

      expect(result.success).toBe(true);
      expect(result.totalRecords).toBe(2);
      expect(result.successfulRecords).toBe(2);
    });
  });

  describe("inspectExcelFile", () => {
    test("should return file structure information", async () => {
      const result = await inspectExcelFile({
        excelPath: "tests/fixtures/integration-test-data.xlsx",
      });

      expect(result.sheetNames).toContain("Sheet1");
      expect(result.firstSheetName).toBe("Sheet1");
      expect(result.recordCount).toBe(2);
      expect(result.fields).toEqual(["name", "company"]);
    });
  });
});
