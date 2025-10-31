import { describe, expect, test } from "bun:test";
import XLSX from "xlsx";
import {
  getFirstSheet,
  loadWorkbook,
  parseSheetData,
} from "../../src/excel-loader";

describe("ExcelLoader", () => {
  describe("loadWorkbook", () => {
    test("should load valid Excel buffer", () => {
      // Create a simple Excel buffer
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet([
        ["Name", "Email"],
        ["John", "john@test.com"],
      ]);
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

      const result = loadWorkbook({ dataBuffer: buffer });
      expect(result.SheetNames).toEqual(["Sheet1"]);
    });

    test("should throw for invalid buffer", () => {
      const invalidBuffer = new ArrayBuffer(0);
      expect(() => loadWorkbook({ dataBuffer: invalidBuffer })).toThrow(
        "Failed to read Excel file: Empty file buffer"
      );
    });
  });

  describe("getFirstSheet", () => {
    test("should return first sheet", () => {
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet([["Name"], ["John"]]);
      XLSX.utils.book_append_sheet(workbook, worksheet, "TestSheet");

      const sheet = getFirstSheet({ workbook });
      expect(sheet).toBeDefined();
    });

    test("should throw for workbook with no sheets", () => {
      const workbook = XLSX.utils.book_new();
      expect(() => getFirstSheet({ workbook })).toThrow(
        "Excel file contains no sheets"
      );
    });
  });

  describe("parseSheetData", () => {
    test("should parse sheet data to records", () => {
      const sheet = XLSX.utils.aoa_to_sheet([
        ["Name", "Email"],
        ["John", "john@test.com"],
        ["Jane", "jane@test.com"],
      ]);

      const records = parseSheetData({ sheet });
      expect(records).toEqual([
        { Name: "John", Email: "john@test.com" },
        { Name: "Jane", Email: "jane@test.com" },
      ]);
    });

    test("should filter empty rows", () => {
      const sheet = XLSX.utils.aoa_to_sheet([
        ["Name", "Email"],
        ["John", "john@test.com"],
        ["", ""], // Empty row
        ["Jane", "jane@test.com"],
      ]);

      const records = parseSheetData({ sheet });
      expect(records).toHaveLength(2);
    });

    test("should throw for empty sheet", () => {
      const sheet = XLSX.utils.aoa_to_sheet([]);
      expect(() => parseSheetData({ sheet })).toThrow("No data records found");
    });
  });
});
