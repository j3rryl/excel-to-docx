#!/usr/bin/env bun

import { program } from "commander";
import { generateDocuments } from "./index.js";
import { fileURLToPath } from "url";
import path from "path";
import type { CliOptions } from "./types.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const packageJson = JSON.parse(
  await Bun.file(path.join(__dirname, "../package.json")).text()
);

program
  .name("excel-to-docx")
  .description(
    "Generate multiple DOCX documents from Excel data using Word templates"
  )
  .version(packageJson.version)
  .argument("<excel-file>", "Path to Excel file with data (.xlsx, .xls)")
  .argument("<template-file>", "Path to Word template file (.docx)")
  .option("-o, --output <dir>", "Output directory", "./output")
  .option(
    "-n, --name <template>",
    "Filename template (use {{FieldName}} for placeholders)",
    "{{name}}"
  )
  .option("-v, --verbose", "Show detailed output", false)
  .action(
    async (excelFile: string, templateFile: string, options: CliOptions) => {
      try {
        const result = await generateDocuments(excelFile, templateFile, {
          outputDir: options.output,
          fileNameTemplate: options.name,
          verbose: options.verbose,
        });

        if (result.success && !options.verbose) {
          console.log(
            `Successfully generated ${result.successfulRecords} documents.`
          );
        }

        if (result.errors.length > 0 && options.verbose) {
          console.log("\nErrors:");
          result.errors.forEach((error) => {
            console.log(`Record ${error.record}: ${error.error}`);
          });
        }
      } catch (error) {
        console.error("Error:", error instanceof Error ? error.message : error);
        process.exit(1);
      }
    }
  );

program.parse();
