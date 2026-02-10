# Banking Document Generator API

A .NET Web API for generating professional banking documents in PDF and Word (DOCX) formats using template-based generation with pure C# libraries - no external dependencies required!

## Features

- **10 Document Types**: Loan Proposal, Promissory Note, Letter of Continuity, Letter of Revival, Letter of Arrangement, Standing Order, Letter of Indemnity, Letter of Lien, Personal Guarantee, and UDC
- **Dual Output Formats**: PDF and Word (DOCX)
- **Template-Based Generation**: Uses Word templates with placeholder replacement
- **Pure C# PDF Generation**: No LibreOffice or external tools needed
- **Batch Generation**: Generate multiple documents as a ZIP file
- **Full Formatting Preservation**: Bold, italic, colors, tables, headers, footers

## Technology Stack

- **.NET 9** - ASP.NET Core Web API
- **DocumentFormat.OpenXml** (v3.0.1) - For reading and manipulating Word documents
- **QuestPDF** (2025.12.4) - For PDF generation without external dependencies
- **Syncfusion.DocIO** - Used only for template generation (not required at runtime)

## Project Structure

```
dynamic_report/
├── BankingDocumentAPI/
│   ├── Controllers/
│   │   └── DocumentController.cs      # API endpoints
│   ├── Models/
│   │   ├── DocumentElement.cs          # Parsed document models
│   │   └── *Data.cs                    # Document data models
│   ├── Services/
│   │   ├── WordToPdfConverterService.cs # Word to PDF converter
│   │   └── WordDocumentService.cs      # Document generation service
│   ├── Templates/                      # Word document templates
│   └── Program.cs                      # API configuration
├── GenerateTemplates/
│   └── Program.cs                      # Template generation utility
└── index.html                          # Web interface
```

## Getting Started

### Prerequisites

- .NET 9 SDK
- Any web browser

### Running the API

1. **Navigate to the API directory:**
   ```bash
   cd BankingDocumentAPI
   ```

2. **Restore dependencies:**
   ```bash
   dotnet restore
   ```

3. **Run the API:**
   ```bash
   dotnet run
   ```

The API will start on `http://localhost:5165`

### Using the Web Interface

1. Open `index.html` in your web browser
2. Enter a Loan ID (e.g., 12345)
3. Select a Document Type
4. Click "Generate PDF" or "Generate Word"

### API Endpoints

#### Generate Single Document

```http
POST /api/document/generate
Content-Type: application/json

{
  "loanId": 12345,
  "documentType": 0,
  "outputFormat": 1
}
```

**Parameters:**
- `loanId` (integer): Loan/Account ID
- `documentType` (integer): Document type (0-9)
  - 0: LoanProposal
  - 1: PromissoryNote
  - 2: LetterOfContinuity
  - 3: LetterOfRevival
  - 4: LetterOfArrangement
  - 5: StandingOrder
  - 6: LetterOfIndemnity
  - 7: LetterOfLien
  - 8: PersonalGuarantee
  - 9: UDC
- `outputFormat` (integer): Output format
  - 0: DOCX (Word)
  - 1: PDF

#### Generate Batch Documents (ZIP)

```http
POST /api/document/batch-generate
Content-Type: application/json

[
  {
    "loanId": 12345,
    "documentType": 0,
    "outputFormat": 1
  },
  {
    "loanId": 12345,
    "documentType": 1,
    "outputFormat": 1
  }
]
```

#### Health Check

```http
GET /api/document/health
```

## How It Works

### Template-Based Generation

1. **Word Templates**: Located in `BankingDocumentAPI/Templates/`
2. **Placeholder Replacement**: Templates contain placeholders like `{ApplicantName}`, `{LoanAmount}`, etc.
3. **OpenXML Parsing**: The `DocumentFormat.OpenXml` library reads and manipulates the Word documents
4. **PDF Generation**: `QuestPDF` recreates the document structure as a PDF

### No External Dependencies

Unlike traditional approaches that require LibreOffice or other external tools, this solution uses:
- **DocumentFormat.OpenXml** to parse Word document structure
- **QuestPDF** to generate PDFs programmatically
- Pure C# code - no external processes required

## Modifying Templates

To modify document templates:

1. Edit the Word documents in `BankingDocumentAPI/Templates/`
2. Use placeholders like `{PropertyName}` for dynamic data
3. Rebuild the project to include updated templates
4. Templates are auto-copied to the build output directory

### Supported Formatting

- **Text Formatting**: Bold, italic, underline, font size, font family, colors
- **Paragraph Formatting**: Alignment (left, center, right, justify), spacing, indentation
- **Tables**: Full table support with borders and cell merging
- **Headers & Footers**: Including page numbers

## Development

### Building the Project

```bash
dotnet build
```

### Running Tests

```bash
dotnet test
```

### Generating New Templates

```bash
cd GenerateTemplates
dotnet run
```

This will regenerate all Word document templates in `BankingDocumentAPI/Templates/`.

## License

This project uses the following libraries:
- **QuestPDF** - Community License (free for open source and small businesses)
- **DocumentFormat.OpenXml** - MIT License
- **Syncfusion.DocIO** - Used only for template generation (trial version, can be replaced)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- QuestPDF for the excellent PDF generation library
- Microsoft for DocumentFormat.OpenXml
- Syncfusion for Word document manipulation tools
