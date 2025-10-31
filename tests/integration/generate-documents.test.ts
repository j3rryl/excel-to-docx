import { describe, expect, test, afterEach } from "bun:test";
import { generateDocuments } from "../../src/index.js";

describe("Integration Tests", () => {
  afterEach(async () => {
    await Bun.$`rm -rf integration-output`.quiet();
  });

  test("should handle complex Excel data structures", async () => {
    const result = await generateDocuments({
      excelPath: "tests/fixtures/integration-test-data.xlsx",
      templatePath: "tests/fixtures/simple-template.docx",
      options: {
        outputDir: "integration-output",
        fileNameTemplate: "Employee_{{name}}",
        cleanFileName: true,
        verbose: false,
      },
    });

    expect(result.success).toBe(true);
    expect(result.totalRecords).toBe(2); // Should filter empty row
    expect(result.successfulRecords).toBe(2);
    expect(result.generatedFiles).toHaveLength(2);
  });

  test("should handle custom file naming templates", async () => {
    const result = await generateDocuments({
      excelPath: "tests/fixtures/integration-test-data.xlsx",
      templatePath: "tests/fixtures/simple-template.docx",
      options: {
        outputDir: "integration-output",
        fileNameTemplate: "{{name}}_{{company}}_Profile",
        verbose: false,
      },
    });

    expect(result.success).toBe(true);
    // Check that files were created with correct naming pattern
    result.generatedFiles.forEach((file) => {
      expect(file).toMatch(/integration-output\/.+_.+_Profile\.docx/);
    });
  });
});
