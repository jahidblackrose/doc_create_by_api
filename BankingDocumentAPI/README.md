# Banking Document API

A .NET Web API for generating banking documents including Loan Proposals, Promissory Notes, Letters of Continuity, and more.

## Features

- ✅ Template-based document generation
- ✅ PDF and Word (DOCX) output formats
- ✅ Batch document generation (ZIP)
- ✅ Mock data for testing
- ✅ RESTful API with Swagger documentation
- ✅ Support for 9 document types

## Supported Document Types

1. Loan Proposal
2. Promissory Note
3. Letter of Continuity
4. Letter of Revival
5. Letter of Arrangement
6. Standing Order
7. Letter of Indemnity
8. Letter of Lien
9. Personal Guarantee
10. UDC (Undertaking of Debt Creation)

## Getting Started

### Prerequisites

- .NET 8.0 SDK or later
- Syncfusion Community License (free) or paid license

### Installation

1. Clone the repository
2. Navigate to the project folder
3. Restore NuGet packages:

```bash
dotnet restore
```

4. Build the project:

```bash
dotnet build
```

5. Run the application:

```bash
dotnet run
```

The API will be available at `http://localhost:5000` (or similar port).
Swagger UI will be available at the root URL.

## API Endpoints

### Generate Document

**POST** `/api/document/generate`

Generate a single document.

**Request Body:**
```json
{
  "loanId": 12345,
  "documentType": "LoanProposal",
  "outputFormat": "PDF"
}
```

**Response:** Document file (PDF or DOCX)

### Download Document

**GET** `/api/document/download/{loanId}/{documentType}?format=PDF`

Download a document via GET request.

### Batch Generate

**POST** `/api/document/batch-generate`

Generate multiple documents and return as ZIP.

**Request Body:**
```json
[
  {
    "loanId": 12345,
    "documentType": "LoanProposal",
    "outputFormat": "PDF"
  },
  {
    "loanId": 12345,
    "documentType": "PromissoryNote",
    "outputFormat": "DOCX"
  }
]
```

**Response:** ZIP file containing all documents

### Get Document Types

**GET** `/api/document/document-types`

Get list of all available document types.

### Health Check

**GET** `/api/document/health`

Check API health status.

## Templates

Place your Word document templates in the `Templates/` folder. Templates should contain placeholders in the format `{PropertyName}`.

Example:
```
{ApplicantName}
{LoanAmount}
{InterestRate}
```

For detailed template specifications, see `Templates/README.md`.

## Mock Data

The API includes mock data for testing. When you call any endpoint with a loan ID, it will return sample data:

- **Loan ID 12345**: Returns sample data for all document types
- Custom loan IDs will also work with generated mock data

## Syncfusion License

This project uses Syncfusion DocIO for Word document generation.

For development/testing, you can get a free community license from: https://www.syncfusion.com/products/communitylicense

To register your license, uncomment the following line in `Program.cs` and replace with your key:

```csharp
SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY_HERE");
```

## Project Structure

```
BankingDocumentAPI/
├── Controllers/
│   └── DocumentController.cs      # API endpoints
├── Services/
│   ├── IDocumentService.cs        # Service interface
│   └── WordDocumentService.cs     # Document generation logic + Mock data
├── Models/
│   ├── DocumentRequest.cs         # Request models and enums
│   ├── LoanProposalData.cs        # Loan proposal data model
│   ├── PromissoryNoteData.cs      # Promissory note data model
│   ├── LetterOfContinuityData.cs  # Letter of continuity data model
│   └── BaseDocumentData.cs        # Other document data models
├── Templates/
│   └── README.md                  # Template specifications
├── Program.cs                     # Application configuration
└── appsettings.json               # Application settings
```

## Usage Examples

### cURL

Generate a Loan Proposal PDF:
```bash
curl -X POST http://localhost:5000/api/document/generate \
  -H "Content-Type: application/json" \
  -d '{"loanId":12345,"documentType":"LoanProposal","outputFormat":"PDF"}' \
  --output loan_proposal.pdf
```

### JavaScript/Fetch

```javascript
const response = await fetch('http://localhost:5000/api/document/generate', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    loanId: 12345,
    documentType: 'LoanProposal',
    outputFormat: 'PDF'
  })
});

const blob = await response.blob();
const url = window.URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'loan_proposal.pdf';
a.click();
```

## Development

### Adding New Document Types

1. Add enum value to `DocumentType` in `Models/DocumentRequest.cs`
2. Create data model class in `Models/`
3. Add mock data method in `Services/WordDocumentService.cs`
4. Add case to `GetDocumentDataAsync` switch statement
5. Add template mapping in `GetTemplateName` method

### Database Integration

To integrate with a real database:

1. Add your database context to `WordDocumentService`
2. Replace mock data methods with database queries
3. Example:

```csharp
private async Task<LoanProposalData> GetLoanProposalDataAsync(long loanId)
{
    var loan = await _dbContext.Loans
        .Include(l => l.Applicant)
        .Include(l => l.Branch)
        .FirstOrDefaultAsync(l => l.Id == loanId);

    return new LoanProposalData
    {
        ApplicantName = loan.Applicant.Name,
        LoanAmount = loan.Amount,
        // ... map other properties
    };
}
```

## License

This project is for educational and commercial use.

## Support

For issues or questions, please create an issue in the repository.
