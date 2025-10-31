# Excel to DOCX 📄✨

A powerful, type-safe document generation tool that creates multiple DOCX documents from Excel data using Word templates. Built with Bun and TypeScript for maximum performance and reliability.

## Features

- **Blazing Fast** - Built with Bun for superior performance
- **Excel to DOCX** - Convert Excel data into multiple Word documents
- **Type-Safe** - Full TypeScript support with excellent IDE completion
- **Template Engine** - Use simple `{{placeholder}}` syntax in your Word templates
- **Error Handling** - Comprehensive error reporting and validation
- **Flexible** - Customizable file naming and output options
- **CLI & API** - Use as command-line tool or programmatic API

## Installation

### For Bun Users (Recommended)

```bash
# Install globally
bun install -g excel-to-docx

# Or use with bunx
bunx excel-to-docx
```

### For Node.js Users

```bash
# Install globally
npm install -g excel-to-docx

# Or use with npx
npx excel-to-docx
```

## Quick Start

### 1. Create Your Excel Data (`data.xlsx`)

| name        | email         | company     | date       |
| ----------- | ------------- | ----------- | ---------- |
| John Doe    | john@test.com | Acme Inc    | 2024-01-01 |
| Jane Smith  | jane@test.com | Globex Corp | 2024-01-02 |
| Bob Johnson | bob@test.com  | Stark Ind   | 2024-01-03 |

### 2. Create Your Word Template (`template.docx`)

Create a Word document with placeholders:

```
# CONTRACT AGREEMENT

This agreement is made between **{{Company}}** and the undersigned.

**Client Information:**
- Name: {{name}}
- Email: {{email}}
- Date: {{date}}

## Terms and Conditions

1. Payment terms: Net 30 days
2. Delivery: Within 14 business days
3. Support: 24/7 available

Signed,
{{name}}
{{date}}
```

### 3. Generate Documents

```bash
excel-to-docx data.xlsx template.docx
```

### 4. Find Your Documents

The generated files will be in the `output/` folder:

```
output/
├── John_Doe.docx
├── Jane_Smith.docx
└── Bob_Johnson.docx
```

## Advanced Usage

### Custom Output Directory

```bash
excel-to-docx data.xlsx template.docx -o ./contracts
```

### Custom File Naming

```bash
excel-to-docx data.xlsx template.docx -n "Contract_{{company}}_{{date}}"
```

This creates files like: `Name.docx`

### Verbose Output

```bash
excel-to-docx data.xlsx template.docx -v
```

### All Options Combined

```bash
excel-to-docx data.xlsx template.docx -o ./reports -n "Report_{{name}}" -v
```

## API Usage

You can also use `excel-to-docx` programmatically in your TypeScript/JavaScript projects:

### Installation

```bash
bun add excel-to-docx
```

### Basic Usage

```typescript
import { generateDocuments } from "excel-to-docx";

const result = await generateDocuments("data.xlsx", "template.docx", {
  outputDir: "./contracts",
  fileNameTemplate: "{{name}}_Agreement",
  verbose: true,
});

console.log(`Generated ${result.successfulRecords} documents`);
console.log(`Files: ${result.generatedFiles.join(", ")}`);
```

### Advanced API Usage

```typescript
import { generateDocuments } from "excel-to-docx";

async function generateContractDocuments() {
  try {
    const result = await generateDocuments("data.xlsx", "template.docx", {
      outputDir: "./legal-docs",
      fileNameTemplate: "Contract_{{company}}_{{cate}}",
      cleanFileName: true,
      verbose: true,
    });

    if (result.success) {
      console.log("All documents generated successfully!");
    } else {
      console.log("Generated with some errors:");
      result.errors.forEach((error) => {
        console.log(`Record ${error.record}: ${error.error}`);
      });
    }

    return result;
  } catch (error) {
    console.error("Fatal error:", error);
  }
}

// Run the function
generateContractDocuments();
```

## Template Syntax

### Basic Placeholders

Use `{{FieldName}}` in your Word template where `FieldName` matches your Excel column headers.

**Excel:**
| FirstName | LastName | Email |
|-----------|----------|-------|
| John | Doe | john@test.com |

**Word Template:**

```
Hello {{FirstName}} {{LastName}}!

Your email is: {{Email}}
```

### Supported Data Types

- **Strings**: `{{Name}}`, `{{Email}}`
- **Dates**: `{{Date}}`, `{{CreatedAt}}`
- **Numbers**: `{{Amount}}`, `{{Quantity}}`
- **Boolean**: `{{IsActive}}`, `{{Approved}}`

### Advanced Template Features

The underlying [docxtemplater](https://docxtemplater.com/) engine supports:

- Loops and conditionals
- Tables and lists
- Images and rich content
- Custom parsers and modules

## Error Handling

The tool provides comprehensive error reporting:

### Common Issues and Solutions

**"Excel file not found"**

- Check the file path is correct
- Ensure the file has .xlsx or .xls extension

**"No data records found"**

- Verify your Excel file has data beyond the header row
- Check that there are no completely empty rows

**"Template file not found"**

- Ensure the template is a .docx file
- Check file permissions

**"Placeholder not replaced"**

- Verify column names in Excel match placeholders in template
- Check for typos in placeholder syntax

## Examples

Check the `examples/` directory for sample files:

```bash
# Clone the repository
git clone https://github.com/j3rryl/excel-to-docx
cd excel-to-docx

# Run with example files
excel-to-docx examples/data.xlsx examples/template.docx -v
```

## Configuration Options

### CLI Options

| Option      | Short | Default    | Description                              |
| ----------- | ----- | ---------- | ---------------------------------------- |
| `--output`  | `-o`  | `./output` | Output directory for generated documents |
| `--name`    | `-n`  | `{{name}}` | Filename template pattern                |
| `--verbose` | `-v`  | `false`    | Show detailed output                     |
| `--help`    | `-h`  |            | Show help message                        |
| `--version` |       |            | Show version number                      |

### API Options

| Option             | Type      | Default      | Description                              |
| ------------------ | --------- | ------------ | ---------------------------------------- |
| `outputDir`        | `string`  | `"./output"` | Output directory path                    |
| `fileNameTemplate` | `string`  | `"{{name}}"` | Filename pattern with placeholders       |
| `cleanFileName`    | `boolean` | `true`       | Remove special characters from filenames |
| `verbose`          | `boolean` | `false`      | Enable detailed logging                  |

## Programmatic Result Object

When using the API, you receive a detailed result object:

```typescript
interface GenerationResult {
  success: boolean; // Overall success status
  generatedFiles: string[]; // Array of generated file paths
  totalRecords: number; // Total records processed
  successfulRecords: number; // Number of successfully generated documents
  errors: Array<{
    // Array of errors encountered
    record: number; // Record number that failed (1-based)
    error: string; // Error description
  }>;
}
```

## Contributing

We welcome contributions! Here's how to get started:

### Development Setup

```bash
# Clone the repository
git clone https://github.com/j3rryl/excel-to-docx
cd excel-to-docx

# Install dependencies
bun install

# Run in development mode
bun run dev --help

# Run tests (add your test files)
bun test
```

### Project Structure

```
src/
├── index.ts              # Main API entry point
├── cli.ts               # Command-line interface
├── types.ts             # TypeScript type definitions
├── utils.ts             # Utility functions
├── excel-loader.ts      # Excel file parsing and validation
└── document-generator.ts # DOCX template processing
```

### Building for Distribution

```bash
# Build for production
bun run build

# Test the built version
bun run start --help
```

## License

MIT License - see LICENSE file for details.

## Support

- 📖 **Documentation**: [GitHub Repository](https://github.com/j3rryl/excel-to-docx)
- 🐛 **Issues**: [GitHub Issues](https://github.com/j3rryl/excel-to-docx/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/j3rryl/excel-to-docx/discussions)

## Related Projects

- [docxtemplater](https://docxtemplater.com/) - The powerful template engine behind this tool
- [SheetJS](https://sheetjs.com/) - Excel file parsing library
- [PizZip](https://stuk.github.io/jszip/) - ZIP library for DOCX manipulation

---

**Happy Document Generating!** 🎉

If you find this tool useful, please consider giving it a ⭐ on [GitHub](https://github.com/j3rryl/excel-to-docx)!
