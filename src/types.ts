export interface ExcelRecord {
  [key: string]: any;
  CustomerName?: string;
  Email?: string;
  Date?: string;
}

export interface GenerateOptions {
  /** Output directory for generated documents */
  outputDir?: string;
  /** Template for filenames using {{FieldName}} syntax */
  fileNameTemplate?: string;
  /** Clean filenames of special characters */
  cleanFileName?: boolean;
  /** Show detailed output */
  verbose?: boolean;
}

export interface GenerationResult {
  /** Overall success status */
  success: boolean;
  /** Array of generated file paths */
  generatedFiles: string[];
  /** Total records processed */
  totalRecords: number;
  /** Number of successfully generated documents */
  successfulRecords: number;
  /** Array of errors encountered */
  errors: Array<{
    /** Record number that failed (1-based) */
    record: number;
    /** Error description */
    error: string;
  }>;
}

/** Utility type for file validation results */
export interface FileValidationResult {
  isValid: boolean;
  errors: string[];
  excelFile?: {
    exists: boolean;
    size: number;
  };
  templateFile?: {
    exists: boolean;
    size: number;
  };
}

export interface CliOptions {
  output?: string;
  name?: string;
  verbose?: boolean;
}
