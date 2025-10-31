import { describe, expect, test } from "bun:test";
import { generateFilename, validateFiles } from "../../src";

describe("Utils", () => {
  describe("generateFilename", () => {
    test("should replace placeholders with record values", () => {
      const record = { Name: "John", Company: "Acme" };
      const fileNameTemplate = "Contract_{{Name}}_{{Company}}";

      const result = generateFilename({ record, fileNameTemplate });
      expect(result).toBe("Contract_John_Acme.docx");
    });

    test("should clean special characters from filename", () => {
      const record = { Name: "John/Doe", Company: "Acme & Co" };
      const fileNameTemplate = "{{Name}} from {{Company}}";

      const result = generateFilename({
        record,
        fileNameTemplate,
        cleanFileName: true,
      });
      expect(result).toBe("John_Doe_from_Acme___Co.docx");
    });

    test("should use fallback name when no placeholders match", () => {
      const record = { Name: "John" };
      const fileNameTemplate = "{{InvalidField}}";

      const result = generateFilename({ record, fileNameTemplate });
      expect(result).toMatch(/^Document_\d+\.docx$/);
    });

    test("should ensure .docx extension", () => {
      const record = { Name: "John" };
      const fileNameTemplate = "{{Name}}";

      const result = generateFilename({ record, fileNameTemplate });
      expect(result).toBe("John.docx");
    });
  });

  describe("validateFiles", () => {
    test("should not throw for existing files", async () => {
      await expect(
        validateFiles({
          excelPath: "tests/fixtures/integration-test-data.xlsx",
          templatePath: "tests/fixtures/simple-template.docx",
        })
      ).resolves.toBeUndefined();

      // Cleanup
      await Bun.$`rm -f test-excel.xlsx test-template.docx`;
    });

    test("should throw for missing Excel file", async () => {
      await Bun.write("test-template.docx", "test");

      await expect(
        validateFiles({
          excelPath: "nonexistent.xlsx",
          templatePath: "test-template.docx",
        })
      ).rejects.toThrow("Excel file not found");

      await Bun.$`rm -f test-template.docx`;
    });

    test("should throw for missing template file", async () => {
      await Bun.write("test-excel.xlsx", "test");

      await expect(
        validateFiles({
          excelPath: "test-excel.xlsx",
          templatePath: "nonexistent.docx",
        })
      ).rejects.toThrow("Template file not found");

      await Bun.$`rm -f test-excel.xlsx`;
    });
  });
});
