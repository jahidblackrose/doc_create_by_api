# Developer Guide - Adding New Document Types

This guide explains how to add a new document type to the Banking Document Generator API.

## Overview

To add a new document type, you need to:
1. Create a data model for the document
2. Generate or create a Word template
3. Add the document type to the API
4. Update the frontend (optional)
5. Test the new document type

## Step-by-Step Guide

### Step 1: Create the Data Model

Create a new model class in `BankingDocumentAPI/Models/`:

```csharp
namespace BankingDocumentAPI.Models
{
    public class NewDocumentData
    {
        // Add your properties here
        public string CustomerName { get; set; } = string.Empty;
        public string AccountNumber { get; set; } = string.Empty;
        public decimal Amount { get; set; }
        public DateTime TransactionDate { get; set; }
        public string ReferenceNumber { get; set; } = string.Empty;

        // Common fields (optional)
        public string BranchName { get; set; } = "Main Branch";
        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BankAddress { get; set; } = "123 Banking Street, City";
        public string GeneratedDate { get; set; } = string.Empty;
    }
}
```

### Step 2: Add the Document Type Enum

Update `BankingDocumentAPI/Models/DocumentType.cs`:

```csharp
public enum DocumentType
{
    LoanProposal = 0,
    PromissoryNote = 1,
    LetterOfContinuity = 2,
    LetterOfRevival = 3,
    LetterOfArrangement = 4,
    StandingOrder = 5,
    LetterOfIndemnity = 6,
    LetterOfLien = 7,
    PersonalGuarantee = 8,
    UDC = 9,
    NewDocument = 10  // Add your new document type here
}
```

### Step 3: Generate a Word Template

#### Option A: Generate Template Programmatically

Add a new method to `GenerateTemplates/Program.cs`:

```csharp
static void GenerateNewDocumentTemplate(string path)
{
    using (WordDocument document = new WordDocument())
    {
        IWSection section = document.AddSection();
        IWParagraph p;

        // Add title
        AddCenteredHeader(section, "NEW DOCUMENT TITLE", BuiltinStyle.Heading1);

        // Add reference line
        p = section.AddParagraph();
        p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");

        // Add content
        p = section.AddParagraph();
        p.AppendText("Customer Details\n");
        p.ApplyStyle(BuiltinStyle.Heading3);

        p = section.AddParagraph();
        AddField(p, "Customer Name", "CustomerName");
        AddField(p, "Account Number", "AccountNumber");

        p = section.AddParagraph();
        p.AppendText("Transaction Details\n");
        p.ApplyStyle(BuiltinStyle.Heading3);

        p = section.AddParagraph();
        p.AppendText("Amount: Rs. {Amount}\n");
        p.AppendText("Transaction Date: {TransactionDate}\n");

        // Add footer
        p = section.AddParagraph();
        p.AppendText("For {BankName}\n");
        p.AppendText("Authorized Signature: _______________\n");

        SaveTemplate(document, path, "NewDocument.docx");
        Console.WriteLine("✓ Generated: NewDocument.docx");
    }
}
```

Then call it from `Main()`:

```csharp
static void Main(string[] args)
{
    // ... existing code ...
    GenerateNewDocumentTemplate(templatePath);  // Add this line
}
```

Run the template generator:

```bash
cd GenerateTemplates
dotnet run
```

#### Option B: Create Template Manually

1. Open Microsoft Word
2. Create your document with placeholders like `{PropertyName}`
3. Save as `NewDocument.docx`
4. Copy to `BankingDocumentAPI/Templates/`

### Step 4: Add Mock Data Method

Add a method to generate mock data in `BankingDocumentAPI/Services/WordDocumentService.cs`:

```csharp
private NewDocumentData GetNewDocumentMockData(long loanId)
{
    return new NewDocumentData
    {
        CustomerName = "John Doe",
        AccountNumber = "1234567890",
        Amount = 50000,
        TransactionDate = DateTime.Now,
        ReferenceNumber = $"ND/{DateTime.Now.Year}/{loanId}",
        BranchName = "Main Branch",
        BankName = "ABC Bank Ltd.",
        BankAddress = "123 Banking Street, City",
        GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy")
    };
}
```

### Step 5: Add Document Type Case

Add a case in the `GenerateSimpleWordFile` method:

```csharp
private byte[] GenerateSimpleWordFile(DocumentType documentType, long loanId)
{
    // ... existing code ...
    switch (documentType)
    {
        // ... existing cases ...
        case DocumentType.NewDocument:
            data = GetNewDocumentMockData(loanId);
            templateName = "NewDocument.docx";
            break;
        default:
            throw new ArgumentException($"Unknown document type: {documentType}");
    }
    // ... rest of the method ...
}
```

### Step 6: Update the Frontend (Optional)

Update `index.html` to add the new document type:

```html
<select id="documentType">
    <!-- existing options -->
    <option value="10">New Document</option>
</select>
```

Also update the JavaScript `documentTypes` array:

```javascript
const documentTypes = [
    'LoanProposal',
    'PromissoryNote',
    // ... existing types ...
    'UDC',
    'NewDocument'  // Add new type here
];
```

### Step 7: Build and Test

1. **Build the project:**

```bash
cd BankingDocumentAPI
dotnet build
```

2. **Run the API:**

```bash
dotnet run
```

3. **Test the new document type:**

```bash
curl -X POST "http://localhost:5165/api/document/generate" \
  -H "Content-Type: application/json" \
  -d '{"loanId":12345,"documentType":10,"outputFormat":1}' \
  --output new_document.pdf
```

Or use the web interface at `index.html`.

## Quick Reference

### Files to Modify

| File | Purpose |
|------|---------|
| `Models/NewDocumentData.cs` | Create new data model |
| `Models/DocumentType.cs` | Add enum value |
| `Services/WordDocumentService.cs` | Add mock data & case |
| `GenerateTemplates/Program.cs` | Generate template |
| `index.html` | Update frontend (optional) |

### Placeholder Syntax

In Word templates, use placeholders in the format:
```
{PropertyName}
```

For example:
- `{CustomerName}` → Replaced with customer name
- `{AccountNumber}` → Replaced with account number
- `{Amount}` → Replaced with amount

### Common Properties

Include these properties in your data model for consistency:

```csharp
public string ReferenceNumber { get; set; } = string.Empty;
public string GeneratedDate { get; set; } = string.Empty;
public string BranchName { get; set; } = "Main Branch";
public string BankName { get; set; } = "ABC Bank Ltd.";
public string BankAddress { get; set; } = "123 Banking Street, City";
```

## Testing Checklist

- [ ] Data model created with all required properties
- [ ] Document type enum updated
- [ ] Template created with proper placeholders
- [ ] Mock data method implemented
- [ ] Case added to switch statement
- [ ] Frontend updated (if needed)
- [ ] API builds successfully
- [ ] Word document generates correctly
- [ ] PDF generates correctly
- [ ] All placeholders replaced with data
- [ ] Formatting preserved in PDF

## Troubleshooting

### Template Not Found

**Error:** "Template not found: NewDocument.docx"

**Solution:** Ensure the template file exists in:
- `BankingDocumentAPI/Templates/` (source)
- `BankingDocumentAPI/bin/Debug/net9.0/Templates/` (build output)

### Placeholders Not Replaced

**Issue:** Placeholders like `{CustomerName}` appear in generated document

**Solutions:**
1. Check property names match between template and data model
2. Ensure mock data method returns non-empty values
3. Verify template is using correct placeholder format `{PropertyName}`

### PDF Formatting Issues

**Issue:** PDF doesn't match Word template formatting

**Notes:**
- The PDF converter (`WordToPdfConverterService`) parses Word structure and recreates it
- Some advanced formatting may have minor differences
- Tables, bold, italic, colors, and alignment are supported

## Example: Adding "Account Statement" Document

### 1. Data Model (`AccountStatementData.cs`)

```csharp
public class AccountStatementData
{
    public string AccountNumber { get; set; } = string.Empty;
    public string AccountHolder { get; set; } = string.Empty;
    public DateTime StatementDate { get; set; }
    public DateTime PeriodStart { get; set; }
    public DateTime PeriodEnd { get; set; }
    public decimal OpeningBalance { get; set; }
    public decimal ClosingBalance { get; set; }
    public string BranchName { get; set; } = "Main Branch";
    public string BankName { get; set; } = "ABC Bank Ltd.";
    public string ReferenceNumber { get; set; } = string.Empty;
}
```

### 2. Update DocumentType Enum

```csharp
AccountStatement = 10,
```

### 3. Template

Use Syncfusion to generate a template with these placeholders:
```
{AccountNumber}
{AccountHolder}
{StatementDate}
{PeriodStart}
{PeriodEnd}
{OpeningBalance}
{ClosingBalance}
```

### 4. Mock Data

```csharp
private AccountStatementData GetAccountStatementMockData(long loanId)
{
    return new AccountStatementData
    {
        AccountNumber = "1234567890",
        AccountHolder = "John Doe",
        StatementDate = DateTime.Now,
        PeriodStart = DateTime.Now.AddMonths(-1),
        PeriodEnd = DateTime.Now,
        OpeningBalance = 100000,
        ClosingBalance = 95000,
        BranchName = "Main Branch",
        BankName = "ABC Bank Ltd.",
        ReferenceNumber = $"AS/{DateTime.Now.Year}/{loanId}"
    };
}
```

### 5. Switch Case

```csharp
case DocumentType.AccountStatement:
    data = GetAccountStatementMockData(loanId);
    templateName = "AccountStatement.docx";
    break;
```

That's it! Your new document type is ready to use.
