import { beforeEach, describe, expect, test } from "bun:test";
import { generateDocument, processRecord } from "../../src";
import { existsSync } from "fs";

describe("DocumentGenerator", () => {
  let templateBuffer: ArrayBuffer;

  beforeEach(async () => {
    // Load a real template file for testing
    templateBuffer = await Bun.file(
      "tests/fixtures/simple-template.docx"
    ).arrayBuffer();
  });
  describe("generateDocument", () => {
    test("should generate document buffer from template", () => {
      const record = { Name: "John", Company: "Acme Inc" };

      const result = generateDocument({ templateBuffer, record });
      expect(result).toBeInstanceOf(Buffer);
      expect(result.length).toBeGreaterThan(0);
    });

    test("should throw for invalid template buffer", () => {
      const invalidBuffer = new ArrayBuffer(0);
      const record = { Name: "John" };

      expect(() =>
        generateDocument({ templateBuffer: invalidBuffer, record })
      ).toThrow("Document generation failed");
    });
  });

  describe("processRecord", () => {
    test("should successfully process record and create file", async () => {
      const record = { Name: "Test User", Company: "Test Corp" };

      // Create test output directory
      await Bun.$`mkdir -p test-output`.quiet();
      const result = await processRecord({
        record,
        index: 0,
        templateBuffer,
        outputDir: "test-output",
        fileNameTemplate: "{{Name}}",
        cleanFileName: true,
        verbose: false,
      });

      expect(result.success).toBe(true);
      expect(result.filePath).toMatch(/test-output\/Test_User\.docx/);

      // Verify file was created
      expect(existsSync(result.filePath!)).toBe(true);

      // Cleanup
      await Bun.$`rm -rf test-output`;
    });

    test("should handle processing errors gracefully", async () => {
      const invalidBuffer = new ArrayBuffer(0);
      const record = { Name: "Test User" };
      const result = await processRecord({
        record,
        index: 0,
        templateBuffer: invalidBuffer,
        outputDir: "test-output",
        fileNameTemplate: "{{Name}}",
        cleanFileName: true,
        verbose: false,
      });

      expect(result.success).toBe(false);
      expect(result.error).toBeDefined();
    });
  });
});
