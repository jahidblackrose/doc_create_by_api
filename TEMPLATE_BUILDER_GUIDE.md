# Template Builder Integration Guide

This guide explains how to add the **Template Builder** feature to your existing Banking Document Generator API project.

---

## Overview

The Template Builder is a **client-side tool** that allows you to:
- Upload a Word document (.docx)
- Auto-extract placeholders like `{Name}`, `{Address}`, `{Amount}`
- Generate all C# code needed for your API
- Get API call format examples

**No backend changes needed!** - It's a standalone HTML page that works in the browser.

---

## Files to Add to Your Project

### 1. Template Builder Page (NEW FILE)

**File:** `template-builder.html`

**Location:** Root of your project (same level as `index.html`)

**Purpose:** The main template builder interface

**How to add:**
1. Copy `template-builder.html` to your project root
2. No modifications needed - it's ready to use

---

### 2. Update Main index.html (MODIFY EXISTING)

**File:** `index.html`

**Location:** Project root

**Add this button inside the `.container` div** (around line 500, after the batch section):

```html
<!-- Template Builder Link -->
<div style="margin-top: 25px; text-align: center;">
    <button onclick="window.location.href='template-builder.html'"
            style="background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
                   color: white;
                   border: none;
                   padding: 15px 30px;
                   border-radius: 10px;
                   font-size: 16px;
                   font-weight: 600;
                   cursor: pointer;
                   transition: all 0.3s;">
        🛠️ Open Template Builder
    </button>
    <p style="margin-top: 10px; color: #666; font-size: 13px;">
        Upload Word templates and auto-generate API code
    </p>
</div>
```

---

## Project Structure After Adding Template Builder

```
dynamic_report/
├── BankingDocumentAPI/           # Existing API project
│   ├── Controllers/
│   ├── Models/
│   ├── Services/
│   ├── Templates/
│   └── ...
├── GenerateTemplates/            # Existing template generator
├── index.html                     # Main web interface (MODIFIED)
├── template-builder.html          # Template builder tool (NEW)
├── README.md                      # Existing
├── DEVELOPER.md                  # Existing
├── TEMPLATE_BUILDER_GUIDE.md      # This file (NEW)
└── history.md                    # Existing
```

---

## How the Template Builder Works

### Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    template-builder.html                      │
│                  (Client-Side Tool)                          │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  1. File Upload         2. Extract Placeholders              │
│  ┌─────────────┐      ┌──────────────────┐                  │
│  │  .docx file │ ───> │ {Name}, {Amount}, │                  │
│  │             │      │ {Address}, etc.   │                  │
│  └─────────────┘      └──────────────────┘                  │
│         │                                                   │
│         ▼                                                   │
│  3. Generate Code                                           │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ • Model Class (NameData.cs)                         │  │
│  │ • DocumentType Enum (DocumentType.cs)               │  │
│  │ • Service Methods (WordDocumentService.cs)           │  │
│  │ • Frontend Update (index.html)                      │  │
│  │ • API Documentation (API-README.md)                 │  │
│  └──────────────────────────────────────────────────────┘  │
│         │                                                   │
│         ▼                                                   │
│  4. Copy/Download Code                                      │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ • Copy button for each file                          │  │
│  │ • Download All button                                │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

### Technology Stack

- **Pure HTML/CSS/JavaScript** - No server needed
- **mammoth.js** - For reading .docx files in the browser (loaded via CDN)
- **Works offline** - Once loaded, no internet connection needed

---

## Step-by-Step: Using Template Builder

### Step 1: Open Template Builder

Two ways to open it:

**Option A: From index.html**
```html
<!-- Click the "🛠️ Open Template Builder" button -->
```

**Option B: Direct URL**
```
http://localhost:5165/template-builder.html
```

### Step 2: Upload Your Word Template

1. Create a Word document with placeholders:
   ```
   Loan Proposal

   Applicant Name: {ApplicantName}
   Loan Amount: Rs. {LoanAmount}
   Address: {ApplicantAddress}
   Date: {GeneratedDate}
   ```

2. Save as `.docx` (e.g., `MyNewDocument.docx`)

3. Upload to the template builder (drag & drop or click)

### Step 3: Configure

Fill in the form:

| Field | Example | Description |
|-------|---------|-------------|
| **Document Name** | `AccountStatement` | Name of your document (PascalCase) |
| **Document Type Value** | `10` | Next available enum number |

**How to find the next enum value:**
- Open `BankingDocumentAPI/Models/DocumentType.cs`
- Find the highest number
- Add 1

Example:
```csharp
public enum DocumentType
{
    LoanProposal = 0,
    PromissoryNote = 1,
    // ...
    UDC = 9,
    // ↑ Highest is 9, so use 10
}
```

### Step 4: Generate Code

Click **"Generate Code ✨"**

You'll get:

1. **DocumentType.cs** - Add enum value
2. **[Name]Data.cs** - Model class
3. **WordDocumentService.cs** - Service methods
4. **index.html** - Frontend dropdown update
5. **API-README.md** - API documentation

### Step 5: Copy Code to Your Project

#### File 1: Update DocumentType.cs

**Location:** `BankingDocumentAPI/Models/DocumentType.cs`

**Copy this code:**
```csharp
public enum DocumentType
{
    // ... existing values ...
    UDC = 9,

    AccountStatement = 10  // <-- Add this line (your document)
}
```

#### File 2: Create New Model

**Location:** `BankingDocumentAPI/Models/AccountStatementData.cs` (create new file)

**Copy this code:**
```csharp
using System;

namespace BankingDocumentAPI.Models
{
    public class AccountStatementData
    {
        public string ApplicantName { get; set; } = string.Empty;
        public decimal LoanAmount { get; set; }
        public string ApplicantAddress { get; set; } = string.Empty;
        public DateTime GeneratedDate { get; set; }

        // Common fields
        public string BranchName { get; set; } = "Main Branch";
        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BankAddress { get; set; } = "123 Banking Street, City";
        public string ReferenceNumber { get; set; } = string.Empty;
    }
}
```

#### File 3: Update WordDocumentService.cs

**Location:** `BankingDocumentAPI/Services/WordDocumentService.cs`

**Add this method:**
```csharp
private AccountStatementData GetAccountStatementMockData(long loanId)
{
    return new AccountStatementData
    {
        ApplicantName = "John Doe",
        LoanAmount = 50000,
        ApplicantAddress = "123 Main Street, City",
        GeneratedDate = DateTime.Now,
        BranchName = "Main Branch",
        BankName = "ABC Bank Ltd.",
        BankAddress = "123 Banking Street, City",
        ReferenceNumber = $"AS/{DateTime.Now.Year}/{loanId}"
    };
}
```

**Add this case to the switch statement:**
```csharp
case DocumentType.AccountStatement:
    data = GetAccountStatementMockData(loanId);
    templateName = "AccountStatement.docx";
    break;
```

#### File 4: Update index.html

**Location:** `index.html`

**Add this option to the select dropdown:**
```html
<option value="10">Account Statement</option>
```

**Add this to the documentTypes array:**
```javascript
const documentTypes = [
    'LoanProposal',
    'PromissoryNote',
    // ... existing ...
    'UDC',
    'AccountStatement'  // <-- Add this
];
```

#### File 5: Add Template

**Location:** `BankingDocumentAPI/Templates/AccountStatement.docx`

**Copy your Word template here** (same one you uploaded to template builder)

### Step 6: Build and Test

```bash
cd BankingDocumentAPI
dotnet build
dotnet run
```

Then test at `index.html` or via API:

```bash
curl -X POST "http://localhost:5165/api/document/generate" \
  -H "Content-Type: application/json" \
  -d '{"loanId":12345,"documentType":10,"outputFormat":1}' \
  --output account_statement.pdf
```

---

## Generated Code Features

### Smart Type Detection

The template builder automatically detects C# types based on placeholder names:

| Placeholder Pattern | Detected Type |
|--------------------|---------------|
| `{Amount}`, `{LoanAmount}`, `{Balance}` | `decimal` |
| `{Date}`, `{CreatedDate}`, `{Dob}` | `DateTime` |
| `{Age}`, `{Count}`, `{Term}` | `int` |
| `{IsActive}`, `{HasPermission}` | `bool` |
| Everything else | `string` |

### Smart Mock Data

Mock data is generated based on field names:

| Field Name | Mock Value |
|-----------|------------|
| `{Name}` | `"John Doe"` |
| `{Address}` | `"123 Main Street, City"` |
| `{PhoneNumber}` | `"+880 1XXX-XXXXXX"` |
| `{Email}` | `"john.doe@example.com"` |
| `{Amount}` | `50000` |
| `{Rate}` | `12.5m` |
| `{Date}` | `DateTime.Now` |

---

## File Copy Summary

### Create New Files

| File | Location | Purpose |
|------|----------|---------|
| `template-builder.html` | Project root | Template builder UI |
| `[DocumentName]Data.cs` | `BankingDocumentAPI/Models/` | Data model |
| `[DocumentName].docx` | `BankingDocumentAPI/Templates/` | Word template |

### Modify Existing Files

| File | Location | Change |
|------|----------|--------|
| `index.html` | Project root | Add link + document option |
| `DocumentType.cs` | `BankingDocumentAPI/Models/` | Add enum value |
| `WordDocumentService.cs` | `BankingDocumentAPI/Services/` | Add mock data + switch case |

---

## Quick Reference Cards

### Card 1: Adding New Document Workflow

```
1. Create Word template with {placeholders}
2. Open template-builder.html
3. Upload .docx file
4. Enter Document Name & Type Value
5. Click "Generate Code"
6. Copy code to respective files
7. Build & Run
8. Test in index.html
```

### Card 2: Enum Values

```
Current Project Enum Values:
0 = LoanProposal
1 = PromissoryNote
2 = LetterOfContinuity
3 = LetterOfRevival
4 = LetterOfArrangement
5 = StandingOrder
6 = LetterOfIndemnity
7 = LetterOfLien
8 = PersonalGuarantee
9 = UDC
10 = YOUR_NEW_DOCUMENT ← Use this
```

### Card 3: API Call Format

```
POST http://localhost:5165/api/document/generate

{
  "loanId": 12345,
  "documentType": 10,     // Your enum value
  "outputFormat": 1       // 0=DOCX, 1=PDF
}
```

---

## Troubleshooting

### Template Builder Not Loading

**Problem:** Page shows blank or errors

**Solution:**
- Check browser console for errors
- Ensure mammoth.js CDN is accessible
- Try opening directly (file:///) or via local server

### Placeholders Not Detected

**Problem:** "No placeholders found" error

**Solution:**
- Ensure placeholders use format `{PropertyName}`
- Check spelling: `{Name}` not `{name}` or `{NAME}`
- Verify file is actual .docx (not .doc)

### Code Won't Compile

**Problem:** Build errors after adding generated code

**Solutions:**

1. **Missing using statement:**
   ```csharp
   using System;  // For DateTime
   ```

2. **Namespace mismatch:**
   - Ensure model is in `BankingDocumentAPI.Models` namespace

3. **Template not found:**
   - Copy .docx to both:
     - `BankingDocumentAPI/Templates/`
     - `BankingDocumentAPI/bin/Debug/net9.0/Templates/`

### Enum Value Conflict

**Problem:** Multiple documents have same enum value

**Solution:**
- Use unique numbers for each document type
- Highest existing value + 1

---

## Advanced Usage

### Custom Type Mappings

Edit `template-builder.html` to customize type detection:

```javascript
function inferCSharpType(propertyName) {
    const lower = propertyName.toLowerCase();

    // Add your custom rules here
    if (lower.includes('salary')) return 'decimal';
    if (lower.includes('employeeid')) return 'int';

    // ... rest of function
}
```

### Custom Mock Values

Edit `template-builder.html` to customize mock data:

```javascript
function getMockValue(propertyName, csType) {
    const lower = propertyName.toLowerCase();

    // Add your custom mock values
    if (lower.includes('salary')) return '50000m';

    // ... rest of function
}
```

---

## Summary

### What You Need to Do:

1. ✅ Copy `template-builder.html` to project root
2. ✅ Add link button to `index.html`
3. ✅ Done! (No API changes needed)

### How to Use:

1. Open `template-builder.html` in browser
2. Upload your Word document
3. Generate and copy code
4. Add to your API project
5. Build and test

---

## Need Help?

- Check `DEVELOPER.md` for manual document creation guide
- Check `README.md` for API documentation
- Open browser DevTools (F12) for JavaScript errors

---

**Last Updated:** 2025-02-13
