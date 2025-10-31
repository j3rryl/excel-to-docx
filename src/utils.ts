/**
 * Generate a safe filename from template and record data
 */
export function generateFilename({
  record,
  fileNameTemplate,
  cleanFileName = true,
}: {
  record: Record<string, any>;
  fileNameTemplate: string;
  cleanFileName?: boolean;
}): string {
  let filename = fileNameTemplate;

  // Replace placeholders with record data
  Object.keys(record).forEach((key) => {
    const placeholder = `{{${key}}}`;

    if (filename.includes(placeholder) && record[key]) {
      filename = filename.replace(
        new RegExp(placeholder, "g"),
        String(record[key])
      );
    }
  });

  // Fallback if no placeholders were replaced
  if (filename === fileNameTemplate) {
    filename = `Document_${Date.now()}`;
  }

  // Clean filename if requested
  if (cleanFileName) {
    filename = filename.replace(/[^\w\d-_.]/g, "_");
  }

  // Ensure .docx extension
  if (!filename.endsWith(".docx")) {
    filename += ".docx";
  }

  return filename;
}

/**
 * Validate file existence
 */
export async function validateFiles({
  excelPath,
  templatePath,
}: {
  excelPath: string;
  templatePath: string;
}): Promise<void> {
  const dataExists = await Bun.file(excelPath).exists();
  const templateExists = await Bun.file(templatePath).exists();

  if (!dataExists) {
    throw new Error(`Excel file not found: ${excelPath}`);
  }

  if (!templateExists) {
    throw new Error(`Template file not found: ${templatePath}`);
  }
}

/**
 * Create output directory if it doesn't exist
 */
export async function ensureOutputDir({
  outputDir,
}: {
  outputDir: string;
}): Promise<void> {
  await Bun.$`mkdir -p ${outputDir}`.quiet();
}
