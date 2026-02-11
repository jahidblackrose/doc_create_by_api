Perfect! Here's a complete .NET solution for generating banking documents via API:

Complete .NET API Solution

1\. Install Required NuGet Packages

bashdotnet add package DocX

dotnet add package iTextSharp.LGPLv2.Core

dotnet add package Syncfusion.DocIO.Net.Core

dotnet add package Syncfusion.DocIORenderer.Net.Core

dotnet add package Syncfusion.Pdf.Net.Core

```



\*\*Note:\*\* DocX is free but limited. For production banking apps, I recommend \*\*Syncfusion\*\* (free with community license) or \*\*Aspose\*\* (paid but very robust).



---



\### \*\*2. Project Structure\*\*

```

BankingDocumentAPI/

├── Controllers/

│   └── DocumentController.cs

├── Services/

│   ├── IDocumentService.cs

│   ├── WordDocumentService.cs

│   └── PdfDocumentService.cs

├── Models/

│   ├── LoanProposalData.cs

│   ├── PromissoryNoteData.cs

│   └── DocumentRequest.cs

├── Templates/

│   ├── LoanProposal.docx

│   ├── PromissoryNote.docx

│   └── LetterOfContinuity.docx

└── Helpers/

&nbsp;   └── DocumentHelper.cs



3\. Models

csharp// Models/DocumentRequest.cs

public class DocumentRequest

{

&nbsp;   public long LoanId { get; set; }

&nbsp;   public DocumentType DocumentType { get; set; }

&nbsp;   public OutputFormat OutputFormat { get; set; } = OutputFormat.PDF;

}



public enum DocumentType

{

&nbsp;   LoanProposal,

&nbsp;   PromissoryNote,

&nbsp;   LetterOfContinuity,

&nbsp;   LetterOfRevival,

&nbsp;   LetterOfArrangement,

&nbsp;   StandingOrder,

&nbsp;   LetterOfIndemnity,

&nbsp;   LetterOfLien,

&nbsp;   PersonalGuarantee,

&nbsp;   UDC

}



public enum OutputFormat

{

&nbsp;   DOCX,

&nbsp;   PDF

}



// Models/LoanProposalData.cs

public class LoanProposalData

{

&nbsp;   // From Database

&nbsp;   public string ApplicantName { get; set; }

&nbsp;   public string AccountNumber { get; set; }

&nbsp;   public decimal LoanAmount { get; set; }

&nbsp;   public string LoanPurpose { get; set; }

&nbsp;   public int TenureMonths { get; set; }

&nbsp;   public decimal InterestRate { get; set; }

&nbsp;   public DateTime ApplicationDate { get; set; }

&nbsp;   public string BranchName { get; set; }

&nbsp;   

&nbsp;   // Static/Calculated

&nbsp;   public string BankName { get; set; } = "ABC Bank Ltd.";

&nbsp;   public string BankAddress { get; set; }

&nbsp;   public decimal MonthlyEMI { get; set; }

&nbsp;   public string GeneratedDate { get; set; }

&nbsp;   public string ReferenceNumber { get; set; }

}



4\. Service Interface

csharp// Services/IDocumentService.cs

public interface IDocumentService

{

&nbsp;   Task<byte\[]> GenerateDocumentAsync(DocumentRequest request);

&nbsp;   Task<byte\[]> GenerateWordDocumentAsync(string templateName, object data);

&nbsp;   Task<byte\[]> GeneratePdfFromWordAsync(byte\[] wordDocument);

}



5\. Word Document Service (Using Syncfusion - Recommended)

csharp// Services/WordDocumentService.cs

using Syncfusion.DocIO;

using Syncfusion.DocIO.DLS;

using Syncfusion.DocIORenderer;

using Syncfusion.Pdf;



public class WordDocumentService : IDocumentService

{

&nbsp;   private readonly IConfiguration \_configuration;

&nbsp;   private readonly ILogger<WordDocumentService> \_logger;

&nbsp;   private readonly string \_templatePath;



&nbsp;   public WordDocumentService(

&nbsp;       IConfiguration configuration, 

&nbsp;       ILogger<WordDocumentService> logger)

&nbsp;   {

&nbsp;       \_configuration = configuration;

&nbsp;       \_logger = logger;

&nbsp;       \_templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");

&nbsp;   }



&nbsp;   public async Task<byte\[]> GenerateDocumentAsync(DocumentRequest request)

&nbsp;   {

&nbsp;       try

&nbsp;       {

&nbsp;           // Get data from database

&nbsp;           var data = await GetDocumentDataAsync(request.LoanId, request.DocumentType);

&nbsp;           

&nbsp;           // Get template name

&nbsp;           string templateName = GetTemplateName(request.DocumentType);

&nbsp;           

&nbsp;           // Generate Word document

&nbsp;           var wordBytes = await GenerateWordDocumentAsync(templateName, data);

&nbsp;           

&nbsp;           // Convert to PDF if needed

&nbsp;           if (request.OutputFormat == OutputFormat.PDF)

&nbsp;           {

&nbsp;               return await GeneratePdfFromWordAsync(wordBytes);

&nbsp;           }

&nbsp;           

&nbsp;           return wordBytes;

&nbsp;       }

&nbsp;       catch (Exception ex)

&nbsp;       {

&nbsp;           \_logger.LogError(ex, "Error generating document");

&nbsp;           throw;

&nbsp;       }

&nbsp;   }



&nbsp;   public async Task<byte\[]> GenerateWordDocumentAsync(string templateName, object data)

&nbsp;   {

&nbsp;       using (FileStream templateStream = new FileStream(

&nbsp;           Path.Combine(\_templatePath, templateName), 

&nbsp;           FileMode.Open, 

&nbsp;           FileAccess.Read))

&nbsp;       {

&nbsp;           // Load template

&nbsp;           using (WordDocument document = new WordDocument(templateStream, FormatType.Docx))

&nbsp;           {

&nbsp;               // Mail merge or find and replace

&nbsp;               await Task.Run(() => ReplaceBookmarks(document, data));

&nbsp;               

&nbsp;               // Save to memory stream

&nbsp;               using (MemoryStream stream = new MemoryStream())

&nbsp;               {

&nbsp;                   document.Save(stream, FormatType.Docx);

&nbsp;                   return stream.ToArray();

&nbsp;               }

&nbsp;           }

&nbsp;       }

&nbsp;   }



&nbsp;   public async Task<byte\[]> GeneratePdfFromWordAsync(byte\[] wordDocument)

&nbsp;   {

&nbsp;       using (MemoryStream wordStream = new MemoryStream(wordDocument))

&nbsp;       using (WordDocument document = new WordDocument(wordStream, FormatType.Docx))

&nbsp;       {

&nbsp;           // Convert Word to PDF

&nbsp;           using (DocIORenderer renderer = new DocIORenderer())

&nbsp;           {

&nbsp;               PdfDocument pdfDocument = renderer.ConvertToPDF(document);

&nbsp;               

&nbsp;               using (MemoryStream pdfStream = new MemoryStream())

&nbsp;               {

&nbsp;                   pdfDocument.Save(pdfStream);

&nbsp;                   pdfDocument.Close();

&nbsp;                   return await Task.FromResult(pdfStream.ToArray());

&nbsp;               }

&nbsp;           }

&nbsp;       }

&nbsp;   }



&nbsp;   private void ReplaceBookmarks(WordDocument document, object data)

&nbsp;   {

&nbsp;       // Use reflection to get all properties

&nbsp;       var properties = data.GetType().GetProperties();

&nbsp;       

&nbsp;       foreach (var prop in properties)

&nbsp;       {

&nbsp;           string bookmarkName = prop.Name;

&nbsp;           string value = prop.GetValue(data)?.ToString() ?? "";

&nbsp;           

&nbsp;           // Find and replace text

&nbsp;           document.Replace($"{{{bookmarkName}}}", value, true, true);

&nbsp;       }

&nbsp;   }



&nbsp;   private async Task<object> GetDocumentDataAsync(long loanId, DocumentType documentType)

&nbsp;   {

&nbsp;       // Fetch from database based on document type

&nbsp;       switch (documentType)

&nbsp;       {

&nbsp;           case DocumentType.LoanProposal:

&nbsp;               return await GetLoanProposalDataAsync(loanId);

&nbsp;           case DocumentType.PromissoryNote:

&nbsp;               return await GetPromissoryNoteDataAsync(loanId);

&nbsp;           // Add other types

&nbsp;           default:

&nbsp;               throw new ArgumentException("Invalid document type");

&nbsp;       }

&nbsp;   }



&nbsp;   private async Task<LoanProposalData> GetLoanProposalDataAsync(long loanId)

&nbsp;   {

&nbsp;       // TODO: Fetch from your database

&nbsp;       // Example:

&nbsp;       // var loan = await \_dbContext.Loans

&nbsp;       //     .Include(l => l.Applicant)

&nbsp;       //     .Include(l => l.Branch)

&nbsp;       //     .FirstOrDefaultAsync(l => l.Id == loanId);

&nbsp;       

&nbsp;       return new LoanProposalData

&nbsp;       {

&nbsp;           ApplicantName = "John Doe",

&nbsp;           AccountNumber = "1234567890",

&nbsp;           LoanAmount = 500000,

&nbsp;           LoanPurpose = "Home Renovation",

&nbsp;           TenureMonths = 60,

&nbsp;           InterestRate = 8.5m,

&nbsp;           ApplicationDate = DateTime.Now,

&nbsp;           BranchName = "Main Branch",

&nbsp;           BankAddress = "123 Banking Street, City",

&nbsp;           MonthlyEMI = 10247.50m,

&nbsp;           GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),

&nbsp;           ReferenceNumber = $"LOAN/{DateTime.Now.Year}/{loanId}"

&nbsp;       };

&nbsp;   }



&nbsp;   private string GetTemplateName(DocumentType documentType)

&nbsp;   {

&nbsp;       return documentType switch

&nbsp;       {

&nbsp;           DocumentType.LoanProposal => "LoanProposal.docx",

&nbsp;           DocumentType.PromissoryNote => "PromissoryNote.docx",

&nbsp;           DocumentType.LetterOfContinuity => "LetterOfContinuity.docx",

&nbsp;           DocumentType.LetterOfRevival => "LetterOfRevival.docx",

&nbsp;           DocumentType.LetterOfArrangement => "LetterOfArrangement.docx",

&nbsp;           DocumentType.StandingOrder => "StandingOrder.docx",

&nbsp;           DocumentType.LetterOfIndemnity => "LetterOfIndemnity.docx",

&nbsp;           DocumentType.LetterOfLien => "LetterOfLien.docx",

&nbsp;           DocumentType.PersonalGuarantee => "PersonalGuarantee.docx",

&nbsp;           DocumentType.UDC => "UDC.docx",

&nbsp;           \_ => throw new ArgumentException("Invalid document type")

&nbsp;       };

&nbsp;   }

}



6\. Controller

csharp// Controllers/DocumentController.cs

\[ApiController]

\[Route("api/\[controller]")]

public class DocumentController : ControllerBase

{

&nbsp;   private readonly IDocumentService \_documentService;

&nbsp;   private readonly ILogger<DocumentController> \_logger;



&nbsp;   public DocumentController(

&nbsp;       IDocumentService documentService, 

&nbsp;       ILogger<DocumentController> logger)

&nbsp;   {

&nbsp;       \_documentService = documentService;

&nbsp;       \_logger = logger;

&nbsp;   }



&nbsp;   \[HttpPost("generate")]

&nbsp;   public async Task<IActionResult> GenerateDocument(\[FromBody] DocumentRequest request)

&nbsp;   {

&nbsp;       try

&nbsp;       {

&nbsp;           var documentBytes = await \_documentService.GenerateDocumentAsync(request);

&nbsp;           

&nbsp;           string fileName = $"{request.DocumentType}\_{request.LoanId}\_{DateTime.Now:yyyyMMddHHmmss}";

&nbsp;           string contentType = request.OutputFormat == OutputFormat.PDF 

&nbsp;               ? "application/pdf" 

&nbsp;               : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

&nbsp;           string extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";

&nbsp;           

&nbsp;           return File(documentBytes, contentType, $"{fileName}.{extension}");

&nbsp;       }

&nbsp;       catch (Exception ex)

&nbsp;       {

&nbsp;           \_logger.LogError(ex, "Error generating document");

&nbsp;           return StatusCode(500, new { error = "Failed to generate document" });

&nbsp;       }

&nbsp;   }



&nbsp;   \[HttpGet("download/{loanId}/{documentType}")]

&nbsp;   public async Task<IActionResult> DownloadDocument(

&nbsp;       long loanId, 

&nbsp;       DocumentType documentType,

&nbsp;       \[FromQuery] OutputFormat format = OutputFormat.PDF)

&nbsp;   {

&nbsp;       var request = new DocumentRequest

&nbsp;       {

&nbsp;           LoanId = loanId,

&nbsp;           DocumentType = documentType,

&nbsp;           OutputFormat = format

&nbsp;       };

&nbsp;       

&nbsp;       return await GenerateDocument(request);

&nbsp;   }



&nbsp;   \[HttpPost("batch-generate")]

&nbsp;   public async Task<IActionResult> GenerateBatchDocuments(\[FromBody] List<DocumentRequest> requests)

&nbsp;   {

&nbsp;       try

&nbsp;       {

&nbsp;           var tasks = requests.Select(r => \_documentService.GenerateDocumentAsync(r));

&nbsp;           var results = await Task.WhenAll(tasks);

&nbsp;           

&nbsp;           // Create ZIP file with all documents

&nbsp;           using (var memoryStream = new MemoryStream())

&nbsp;           {

&nbsp;               using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))

&nbsp;               {

&nbsp;                   for (int i = 0; i < requests.Count; i++)

&nbsp;                   {

&nbsp;                       var request = requests\[i];

&nbsp;                       var extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";

&nbsp;                       var fileName = $"{request.DocumentType}\_{request.LoanId}.{extension}";

&nbsp;                       

&nbsp;                       var entry = archive.CreateEntry(fileName);

&nbsp;                       using (var entryStream = entry.Open())

&nbsp;                       {

&nbsp;                           await entryStream.WriteAsync(results\[i], 0, results\[i].Length);

&nbsp;                       }

&nbsp;                   }

&nbsp;               }

&nbsp;               

&nbsp;               return File(memoryStream.ToArray(), 

&nbsp;                   "application/zip", 

&nbsp;                   $"Documents\_{DateTime.Now:yyyyMMddHHmmss}.zip");

&nbsp;           }

&nbsp;       }

&nbsp;       catch (Exception ex)

&nbsp;       {

&nbsp;           \_logger.LogError(ex, "Error generating batch documents");

&nbsp;           return StatusCode(500, new { error = "Failed to generate documents" });

&nbsp;       }

&nbsp;   }

}



7\. Program.cs / Startup Configuration

csharp// Program.cs (.NET 6+)

var builder = WebApplication.CreateBuilder(args);



// Add services

builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();

builder.Services.AddSwaggerGen();



// Register document service

builder.Services.AddScoped<IDocumentService, WordDocumentService>();



// Add CORS if needed

builder.Services.AddCors(options =>

{

&nbsp;   options.AddPolicy("AllowAll", builder =>

&nbsp;   {

&nbsp;       builder.AllowAnyOrigin()

&nbsp;              .AllowAnyMethod()

&nbsp;              .AllowAnyHeader();

&nbsp;   });

});



var app = builder.Build();



if (app.Environment.IsDevelopment())

{

&nbsp;   app.UseSwagger();

&nbsp;   app.UseSwaggerUI();

}



app.UseHttpsRedirection();

app.UseCors("AllowAll");

app.UseAuthorization();

app.MapControllers();



app.Run();

```



---



\### \*\*8. Template Creation (Word Document)\*\*



Create `LoanProposal.docx` template:

```

&nbsp;                   ABC BANK LTD

&nbsp;               {BankAddress}



&nbsp;               LOAN PROPOSAL



Reference No: {ReferenceNumber}

Date: {GeneratedDate}



Applicant Name: {ApplicantName}

Account Number: {AccountNumber}



Loan Details:

\- Loan Amount: Rs. {LoanAmount}

\- Purpose: {LoanPurpose}

\- Tenure: {TenureMonths} months

\- Interest Rate: {InterestRate}% p.a.

\- Monthly EMI: Rs. {MonthlyEMI}



Branch: {BranchName}

Application Date: {ApplicationDate}



Authorized Signature: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_



9\. API Usage Examples

Single Document:

httpPOST /api/document/generate

Content-Type: application/json



{

&nbsp;   "loanId": 12345,

&nbsp;   "documentType": "LoanProposal",

&nbsp;   "outputFormat": "PDF"

}

Download Direct:

httpGET /api/document/download/12345/PromissoryNote?format=PDF

Batch Generation:

httpPOST /api/document/batch-generate

Content-Type: application/json



\[

&nbsp;   {

&nbsp;       "loanId": 12345,

&nbsp;       "documentType": "LoanProposal",

&nbsp;       "outputFormat": "PDF"

&nbsp;   },

&nbsp;   {

&nbsp;       "loanId": 12345,

&nbsp;       "documentType": "PromissoryNote",

&nbsp;       "outputFormat": "DOCX"

&nbsp;   }

]



10\. Alternative: Using DocX (Free Library)

csharpusing Xceed.Words.NET;



public async Task<byte\[]> GenerateWithDocX(string templateName, object data)

{

&nbsp;   using (var document = DocX.Load(Path.Combine(\_templatePath, templateName)))

&nbsp;   {

&nbsp;       var properties = data.GetType().GetProperties();

&nbsp;       

&nbsp;       foreach (var prop in properties)

&nbsp;       {

&nbsp;           document.ReplaceText($"{{{prop.Name}}}", prop.GetValue(data)?.ToString() ?? "");

&nbsp;       }

&nbsp;       

&nbsp;       using (var stream = new MemoryStream())

&nbsp;       {

&nbsp;           document.SaveAs(stream);

&nbsp;           return stream.ToArray();

&nbsp;       }

&nbsp;   }

}



Recommendation

For production banking application, use:



Syncfusion (best balance of features and cost)

Aspose.Words (most robust, but expensive)



For prototype/testing:



DocX (free, limited features)



This solution gives you:

✅ Template-based generation

✅ Database + static data merge

✅ Word \& PDF output

✅ Batch processing

✅ RESTful API

✅ Easy maintenance

Would you like me to add database integration code or show how to create more complex templates with tables and conditional sections?Why need ├── Templates/ │ ├── LoanProposal.docx │ ├── PromissoryNote.docx │ └── LetterOfContinuity.docx1:45 PMGreat question! Let me explain why templates are essential and show you the alternative approaches.



Why Templates Are Needed

1\. Business-Controlled Formatting

Banking documents have strict legal and regulatory formats that:



Must match exact company letterheads

Require specific layouts, fonts, spacing

Need approval from legal/compliance teams

Change based on regulations



With Templates:

✅ Business team designs in MS Word

✅ Legal team approves the format

✅ Developers just plug in data

✅ Format changes = just update Word file (no code change)

Without Templates (hardcoded):

❌ Developers code every layout detail

❌ Every format change = code deployment

❌ Business can't preview/approve easily

❌ Maintenance nightmare



2\. Separation of Concerns

Template Approach:

├── Business Team → Maintains document format (Word files)

├── Legal Team → Approves templates

└── Dev Team → Handles data + generation logic



Hardcoded Approach:

└── Dev Team → Does EVERYTHING (format + logic + styling)



3\. Real Example

Loan Proposal Template (LoanProposal.docx):

&nbsp;                   \[Company Logo Here]

&nbsp;                   ABC BANK LIMITED

&nbsp;             123 Banking Street, Main Branch

&nbsp;             Tel: 123-456-7890 | www.abcbank.com



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━



&nbsp;                   LOAN PROPOSAL



Reference: {ReferenceNumber}                    Date: {GeneratedDate}



Dear {ApplicantName},



Subject: Loan Proposal for Rs. {LoanAmount}



We are pleased to present the following loan proposal:



┌──────────────────────────────────────┐

│ LOAN DETAILS                          │

├──────────────────────────────────────┤

│ Amount          : Rs. {LoanAmount}    │

│ Purpose         : {LoanPurpose}       │

│ Tenure          : {TenureMonths} mths │

│ Interest Rate   : {InterestRate}%     │

│ Monthly EMI     : Rs. {MonthlyEMI}    │

└──────────────────────────────────────┘



\[Legal terms and conditions - 2 pages]



Authorized Signatory                    Bank Seal

\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_                    \[SEAL]

If you hardcode this:

csharp// 500+ lines of code just for layout!

document.AddParagraph("ABC BANK LIMITED")

&nbsp;   .Font("Arial")

&nbsp;   .FontSize(16)

&nbsp;   .Bold()

&nbsp;   .Alignment(Alignment.center)

&nbsp;   .SpacingAfter(10);



document.AddParagraph("123 Banking Street...")

&nbsp;   .Font("Arial")

&nbsp;   .FontSize(10)

&nbsp;   .Alignment(Alignment.center);



// Add logo

document.AddImage(...)

&nbsp;   .Width(100)

&nbsp;   .Height(50);



// ... 400 more lines for one document!



Alternatives to Templates

Option 1: Hardcoded Document Generation ❌

Don't recommend for banking

csharppublic byte\[] GenerateLoanProposal(LoanData data)

{

&nbsp;   using (WordDocument document = new WordDocument())

&nbsp;   {

&nbsp;       IWSection section = document.AddSection();

&nbsp;       IWParagraph para = section.AddParagraph();

&nbsp;       

&nbsp;       // Hardcode EVERYTHING

&nbsp;       para.AppendText("ABC BANK LIMITED");

&nbsp;       para.ApplyStyle(BuiltinStyle.Heading1);

&nbsp;       

&nbsp;       para = section.AddParagraph();

&nbsp;       para.AppendText($"Applicant: {data.ApplicantName}");

&nbsp;       

&nbsp;       // ... 100s of lines per document type

&nbsp;       // ... nightmare to maintain

&nbsp;       

&nbsp;       using (MemoryStream stream = new MemoryStream())

&nbsp;       {

&nbsp;           document.Save(stream, FormatType.Docx);

&nbsp;           return stream.ToArray();

&nbsp;       }

&nbsp;   }

}

Problems:



❌ 500+ lines of code per document

❌ Format changes require code deployment

❌ Can't match exact business format easily

❌ Legal team can't approve format before development





Option 2: HTML to PDF Conversion ⚠️

Possible, but limited

csharppublic byte\[] GenerateFromHtml(LoanData data)

{

&nbsp;   string html = $@"

&nbsp;       <html>

&nbsp;       <head>

&nbsp;           <style>

&nbsp;               .header {{ text-align: center; font-size: 20px; }}

&nbsp;               .content {{ margin: 20px; }}

&nbsp;           </style>

&nbsp;       </head>

&nbsp;       <body>

&nbsp;           <div class='header'>ABC BANK LIMITED</div>

&nbsp;           <div class='content'>

&nbsp;               <p>Applicant: {data.ApplicantName}</p>

&nbsp;               <p>Amount: {data.LoanAmount}</p>

&nbsp;           </div>

&nbsp;       </body>

&nbsp;       </html>

&nbsp;   ";

&nbsp;   

&nbsp;   // Convert HTML to PDF

&nbsp;   return ConvertHtmlToPdf(html);

}

Problems:



⚠️ HTML/CSS doesn't match Word/PDF exactly

⚠️ Hard to maintain complex layouts

⚠️ Banking documents need precise formatting (letterheads, signatures, seals)





Option 3: Database-Stored Templates ✅

Better than hardcoding, but more complex

Instead of file templates, store template structure in database:

sqlCREATE TABLE DocumentTemplates (

&nbsp;   Id INT,

&nbsp;   DocumentType VARCHAR(50),

&nbsp;   TemplateJson TEXT,  -- JSON structure of document

&nbsp;   Version INT

)



-- JSON structure:

{

&nbsp;   "sections": \[

&nbsp;       {

&nbsp;           "type": "header",

&nbsp;           "content": "ABC BANK LIMITED",

&nbsp;           "style": { "fontSize": 16, "alignment": "center" }

&nbsp;       },

&nbsp;       {

&nbsp;           "type": "field",

&nbsp;           "placeholder": "{ApplicantName}",

&nbsp;           "label": "Applicant Name"

&nbsp;       }

&nbsp;   ]

}

csharp// Then parse JSON and generate document

var template = await \_db.DocumentTemplates

&nbsp;   .FirstOrDefaultAsync(t => t.DocumentType == "LoanProposal");

&nbsp;   

var structure = JsonConvert.DeserializeObject<TemplateStructure>(template.TemplateJson);

// Build document from structure

```



\*\*Pros:\*\*

\- ✅ Templates stored centrally

\- ✅ Version control in database

\- ✅ Can build UI for template editing



\*\*Cons:\*\*

\- ⚠️ Complex to build template editor

\- ⚠️ Business team can't use familiar Word

\- ⚠️ More development effort



---



\### \*\*Option 4: Template Files (RECOMMENDED)\*\* ✅✅✅

```

Templates/

├── LoanProposal.docx        ← Business team edits in Word

├── PromissoryNote.docx      ← Legal team approves

└── LetterOfContinuity.docx  ← Just upload new version to update

Code is simple:

csharp// Load template

var template = LoadTemplate("LoanProposal.docx");



// Replace placeholders

template.Replace("{ApplicantName}", data.ApplicantName);

template.Replace("{LoanAmount}", data.LoanAmount);



// Save

return template.ToBytes();

```



\*\*Benefits:\*\*

\- ✅ \*\*5 lines of code\*\* instead of 500

\- ✅ Business team controls format

\- ✅ No code changes for format updates

\- ✅ Easy legal approval process

\- ✅ Exact company branding maintained

\- ✅ Version control (Git tracks template files)



---



\## \*\*Template Workflow in Banking\*\*

```

1\. Business Team

&nbsp;  ↓ Creates LoanProposal.docx in MS Word

&nbsp;  ↓ Adds placeholders: {ApplicantName}, {LoanAmount}

&nbsp;  

2\. Legal Team

&nbsp;  ↓ Reviews and approves format

&nbsp;  

3\. Dev Team

&nbsp;  ↓ Saves template to Templates/ folder

&nbsp;  ↓ Writes code to replace placeholders

&nbsp;  

4\. Format Change Needed?

&nbsp;  ↓ Business updates Word file

&nbsp;  ↓ Legal approves

&nbsp;  ↓ Replace old template file

&nbsp;  ↓ ZERO code changes needed!



What If You Really Don't Want Template Files?

Hybrid Approach: Templates in Database as Binary

csharp// Store template file content in database

CREATE TABLE DocumentTemplates (

&nbsp;   DocumentType VARCHAR(50),

&nbsp;   TemplateContent VARBINARY(MAX),  -- Store .docx file

&nbsp;   Version INT,

&nbsp;   UploadedDate DATETIME

)



// Load from database

public async Task<byte\[]> GenerateDocument(string docType, object data)

{

&nbsp;   var template = await \_db.DocumentTemplates

&nbsp;       .Where(t => t.DocumentType == docType)

&nbsp;       .OrderByDescending(t => t.Version)

&nbsp;       .FirstOrDefaultAsync();

&nbsp;   

&nbsp;   // Load from binary

&nbsp;   using (var stream = new MemoryStream(template.TemplateContent))

&nbsp;   using (var doc = new WordDocument(stream))

&nbsp;   {

&nbsp;       ReplaceBookmarks(doc, data);

&nbsp;       return doc.ToBytes();

&nbsp;   }

}

This gives you:



✅ Centralized storage

✅ Business still uses Word

✅ Version control

⚠️ But adds database complexity





Final Recommendation

For Banking Applications:

Use file-based templates stored in:



Development: Templates/ folder in project

Production: Azure Blob Storage / AWS S3 / Network share



Why:



✅ Regulatory compliance (exact format control)

✅ Business team independence

✅ Legal team can approve easily

✅ Minimal code maintenance

✅ Fast format updates without deployment



Storage Strategy:

csharp// Load from cloud storage

public async Task<byte\[]> LoadTemplate(string templateName)

{

&nbsp;   // Development: Load from local folder

&nbsp;   if (\_env.IsDevelopment())

&nbsp;       return File.ReadAllBytes($"Templates/{templateName}");

&nbsp;   

&nbsp;   // Production: Load from Azure Blob

&nbsp;   var blob = \_blobClient.GetBlobClient(templateName);

&nbsp;   return await blob.DownloadContentAsync();

}



Bottom line: Templates separate WHAT the document says (business logic) from HOW it looks (business format). This is critical in banking where format compliance is as important as data accuracy.

Does this clarify why templates are essential? Would you like to see how to implement template versioning or cloud storage integration?I like Templatehow to implement it1:46 PMExcellent choice! Let me give you a complete step-by-step implementation for template-based document generation.



Complete Implementation Guide



STEP 1: Create .NET Web API Project

bash# Create new Web API project

dotnet new webapi -n BankingDocumentAPI

cd BankingDocumentAPI



\# Install required packages

dotnet add package Syncfusion.DocIO.Net.Core

dotnet add package Syncfusion.DocIORenderer.Net.Core

dotnet add package Syncfusion.Pdf.Net.Core



\# For database (Entity Framework)

dotnet add package Microsoft.EntityFrameworkCore.SqlServer

dotnet add package Microsoft.EntityFrameworkCore.Tools

```



---



\### \*\*STEP 2: Project Structure\*\*

```

BankingDocumentAPI/

├── Controllers/

│   └── DocumentController.cs

├── Services/

│   ├── IDocumentGenerationService.cs

│   └── DocumentGenerationService.cs

├── Data/

│   ├── ApplicationDbContext.cs

│   └── Entities/

│       └── Loan.cs

├── Models/

│   ├── DTOs/

│   │   ├── DocumentRequest.cs

│   │   └── LoanProposalData.cs

│   └── Enums/

│       ├── DocumentType.cs

│       └── OutputFormat.cs

├── Templates/

│   ├── LoanProposal.docx

│   ├── PromissoryNote.docx

│   ├── LetterOfContinuity.docx

│   ├── LetterOfRevival.docx

│   ├── LetterOfArrangement.docx

│   ├── StandingOrder.docx

│   ├── LetterOfIndemnity.docx

│   ├── LetterOfLien.docx

│   ├── PersonalGuarantee.docx

│   └── UDC.docx

├── Helpers/

│   └── DocumentHelper.cs

├── appsettings.json

└── Program.cs



STEP 3: Create Models

Models/Enums/DocumentType.cs

csharpnamespace BankingDocumentAPI.Models.Enums

{

&nbsp;   public enum DocumentType

&nbsp;   {

&nbsp;       LoanProposal = 1,

&nbsp;       PromissoryNote = 2,

&nbsp;       LetterOfContinuity = 3,

&nbsp;       LetterOfRevival = 4,

&nbsp;       LetterOfArrangement = 5,

&nbsp;       StandingOrder = 6,

&nbsp;       LetterOfIndemnity = 7,

&nbsp;       LetterOfLien = 8,

&nbsp;       PersonalGuarantee = 9,

&nbsp;       UDC = 10

&nbsp;   }

}

Models/Enums/OutputFormat.cs

csharpnamespace BankingDocumentAPI.Models.Enums

{

&nbsp;   public enum OutputFormat

&nbsp;   {

&nbsp;       DOCX = 1,

&nbsp;       PDF = 2

&nbsp;   }

}

Models/DTOs/DocumentRequest.cs

csharpnamespace BankingDocumentAPI.Models.DTOs

{

&nbsp;   public class DocumentRequest

&nbsp;   {

&nbsp;       public long LoanId { get; set; }

&nbsp;       public DocumentType DocumentType { get; set; }

&nbsp;       public OutputFormat OutputFormat { get; set; } = OutputFormat.PDF;

&nbsp;   }

}

Models/DTOs/LoanProposalData.cs

csharpnamespace BankingDocumentAPI.Models.DTOs

{

&nbsp;   public class LoanProposalData

&nbsp;   {

&nbsp;       // Database fields

&nbsp;       public string ApplicantName { get; set; }

&nbsp;       public string FatherName { get; set; }

&nbsp;       public string Address { get; set; }

&nbsp;       public string AccountNumber { get; set; }

&nbsp;       public string LoanAmount { get; set; }

&nbsp;       public string LoanAmountInWords { get; set; }

&nbsp;       public string LoanPurpose { get; set; }

&nbsp;       public string TenureMonths { get; set; }

&nbsp;       public string InterestRate { get; set; }

&nbsp;       public string MonthlyEMI { get; set; }

&nbsp;       public string ApplicationDate { get; set; }

&nbsp;       public string BranchName { get; set; }

&nbsp;       public string BranchCode { get; set; }

&nbsp;       public string ContactNumber { get; set; }

&nbsp;       public string Email { get; set; }

&nbsp;       

&nbsp;       // Static fields

&nbsp;       public string BankName { get; set; }

&nbsp;       public string BankAddress { get; set; }

&nbsp;       public string BankPhone { get; set; }

&nbsp;       public string BankEmail { get; set; }

&nbsp;       public string BankWebsite { get; set; }

&nbsp;       

&nbsp;       // Generated fields

&nbsp;       public string ReferenceNumber { get; set; }

&nbsp;       public string GeneratedDate { get; set; }

&nbsp;       public string GeneratedBy { get; set; }

&nbsp;   }

}



STEP 4: Database Entities (Example)

Data/Entities/Loan.cs

csharpnamespace BankingDocumentAPI.Data.Entities

{

&nbsp;   public class Loan

&nbsp;   {

&nbsp;       public long Id { get; set; }

&nbsp;       public string ApplicantName { get; set; }

&nbsp;       public string FatherName { get; set; }

&nbsp;       public string Address { get; set; }

&nbsp;       public string AccountNumber { get; set; }

&nbsp;       public decimal LoanAmount { get; set; }

&nbsp;       public string LoanPurpose { get; set; }

&nbsp;       public int TenureMonths { get; set; }

&nbsp;       public decimal InterestRate { get; set; }

&nbsp;       public DateTime ApplicationDate { get; set; }

&nbsp;       public string BranchName { get; set; }

&nbsp;       public string BranchCode { get; set; }

&nbsp;       public string ContactNumber { get; set; }

&nbsp;       public string Email { get; set; }

&nbsp;       public DateTime CreatedDate { get; set; }

&nbsp;       public string CreatedBy { get; set; }

&nbsp;   }

}

Data/ApplicationDbContext.cs

csharpusing Microsoft.EntityFrameworkCore;

using BankingDocumentAPI.Data.Entities;



namespace BankingDocumentAPI.Data

{

&nbsp;   public class ApplicationDbContext : DbContext

&nbsp;   {

&nbsp;       public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)

&nbsp;           : base(options)

&nbsp;       {

&nbsp;       }



&nbsp;       public DbSet<Loan> Loans { get; set; }

&nbsp;   }

}



STEP 5: Document Generation Service

Services/IDocumentGenerationService.cs

csharpnamespace BankingDocumentAPI.Services

{

&nbsp;   public interface IDocumentGenerationService

&nbsp;   {

&nbsp;       Task<byte\[]> GenerateDocumentAsync(DocumentRequest request);

&nbsp;   }

}

Services/DocumentGenerationService.cs

csharpusing Syncfusion.DocIO;

using Syncfusion.DocIO.DLS;

using Syncfusion.DocIORenderer;

using Syncfusion.Pdf;

using BankingDocumentAPI.Data;

using BankingDocumentAPI.Models.DTOs;

using BankingDocumentAPI.Models.Enums;

using Microsoft.EntityFrameworkCore;

using System.Globalization;



namespace BankingDocumentAPI.Services

{

&nbsp;   public class DocumentGenerationService : IDocumentGenerationService

&nbsp;   {

&nbsp;       private readonly ApplicationDbContext \_context;

&nbsp;       private readonly ILogger<DocumentGenerationService> \_logger;

&nbsp;       private readonly IConfiguration \_configuration;

&nbsp;       private readonly string \_templatePath;



&nbsp;       public DocumentGenerationService(

&nbsp;           ApplicationDbContext context,

&nbsp;           ILogger<DocumentGenerationService> logger,

&nbsp;           IConfiguration configuration)

&nbsp;       {

&nbsp;           \_context = context;

&nbsp;           \_logger = logger;

&nbsp;           \_configuration = configuration;

&nbsp;           \_templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");

&nbsp;       }



&nbsp;       public async Task<byte\[]> GenerateDocumentAsync(DocumentRequest request)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               \_logger.LogInformation($"Generating {request.DocumentType} for Loan ID: {request.LoanId}");



&nbsp;               // Get data from database

&nbsp;               var documentData = await GetDocumentDataAsync(request.LoanId, request.DocumentType);



&nbsp;               // Get template file name

&nbsp;               string templateFileName = GetTemplateFileName(request.DocumentType);

&nbsp;               string templateFullPath = Path.Combine(\_templatePath, templateFileName);



&nbsp;               if (!File.Exists(templateFullPath))

&nbsp;               {

&nbsp;                   throw new FileNotFoundException($"Template not found: {templateFileName}");

&nbsp;               }



&nbsp;               // Generate Word document

&nbsp;               byte\[] wordDocument = await GenerateWordDocumentAsync(templateFullPath, documentData);



&nbsp;               // Convert to PDF if requested

&nbsp;               if (request.OutputFormat == OutputFormat.PDF)

&nbsp;               {

&nbsp;                   return await ConvertWordToPdfAsync(wordDocument);

&nbsp;               }



&nbsp;               return wordDocument;

&nbsp;           }

&nbsp;           catch (Exception ex)

&nbsp;           {

&nbsp;               \_logger.LogError(ex, $"Error generating document for Loan ID: {request.LoanId}");

&nbsp;               throw;

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<byte\[]> GenerateWordDocumentAsync(string templatePath, object data)

&nbsp;       {

&nbsp;           using (FileStream templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))

&nbsp;           {

&nbsp;               // Load the template

&nbsp;               using (WordDocument document = new WordDocument(templateStream, FormatType.Docx))

&nbsp;               {

&nbsp;                   // Replace all placeholders with actual data

&nbsp;                   ReplaceBookmarks(document, data);



&nbsp;                   // Save to memory stream

&nbsp;                   using (MemoryStream outputStream = new MemoryStream())

&nbsp;                   {

&nbsp;                       document.Save(outputStream, FormatType.Docx);

&nbsp;                       return await Task.FromResult(outputStream.ToArray());

&nbsp;                   }

&nbsp;               }

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<byte\[]> ConvertWordToPdfAsync(byte\[] wordDocument)

&nbsp;       {

&nbsp;           using (MemoryStream wordStream = new MemoryStream(wordDocument))

&nbsp;           using (WordDocument document = new WordDocument(wordStream, FormatType.Docx))

&nbsp;           {

&nbsp;               // Convert Word to PDF using DocIORenderer

&nbsp;               using (DocIORenderer renderer = new DocIORenderer())

&nbsp;               {

&nbsp;                   PdfDocument pdfDocument = renderer.ConvertToPDF(document);



&nbsp;                   using (MemoryStream pdfStream = new MemoryStream())

&nbsp;                   {

&nbsp;                       pdfDocument.Save(pdfStream);

&nbsp;                       pdfDocument.Close();

&nbsp;                       return await Task.FromResult(pdfStream.ToArray());

&nbsp;                   }

&nbsp;               }

&nbsp;           }

&nbsp;       }



&nbsp;       private void ReplaceBookmarks(WordDocument document, object data)

&nbsp;       {

&nbsp;           if (data == null) return;



&nbsp;           // Get all properties using reflection

&nbsp;           var properties = data.GetType().GetProperties();



&nbsp;           foreach (var property in properties)

&nbsp;           {

&nbsp;               string placeholder = $"{{{property.Name}}}";

&nbsp;               string value = property.GetValue(data)?.ToString() ?? "";



&nbsp;               // Find and replace in the entire document

&nbsp;               document.Replace(placeholder, value, true, true);

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<object> GetDocumentDataAsync(long loanId, DocumentType documentType)

&nbsp;       {

&nbsp;           switch (documentType)

&nbsp;           {

&nbsp;               case DocumentType.LoanProposal:

&nbsp;                   return await GetLoanProposalDataAsync(loanId);

&nbsp;               

&nbsp;               case DocumentType.PromissoryNote:

&nbsp;                   return await GetPromissoryNoteDataAsync(loanId);

&nbsp;               

&nbsp;               case DocumentType.LetterOfContinuity:

&nbsp;                   return await GetLetterOfContinuityDataAsync(loanId);

&nbsp;               

&nbsp;               // Add other document types

&nbsp;               

&nbsp;               default:

&nbsp;                   throw new ArgumentException($"Unsupported document type: {documentType}");

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<LoanProposalData> GetLoanProposalDataAsync(long loanId)

&nbsp;       {

&nbsp;           var loan = await \_context.Loans

&nbsp;               .FirstOrDefaultAsync(l => l.Id == loanId);



&nbsp;           if (loan == null)

&nbsp;           {

&nbsp;               throw new NotFoundException($"Loan not found with ID: {loanId}");

&nbsp;           }



&nbsp;           // Calculate EMI

&nbsp;           decimal monthlyEMI = CalculateEMI(loan.LoanAmount, loan.InterestRate, loan.TenureMonths);



&nbsp;           // Static bank information (can be from config or database)

&nbsp;           string bankName = \_configuration\["BankInfo:Name"] ?? "ABC Bank Limited";

&nbsp;           string bankAddress = \_configuration\["BankInfo:Address"] ?? "123 Banking Street, Financial District";

&nbsp;           string bankPhone = \_configuration\["BankInfo:Phone"] ?? "+880-123-456789";

&nbsp;           string bankEmail = \_configuration\["BankInfo:Email"] ?? "info@abcbank.com";

&nbsp;           string bankWebsite = \_configuration\["BankInfo:Website"] ?? "www.abcbank.com";



&nbsp;           return new LoanProposalData

&nbsp;           {

&nbsp;               // Database fields

&nbsp;               ApplicantName = loan.ApplicantName,

&nbsp;               FatherName = loan.FatherName,

&nbsp;               Address = loan.Address,

&nbsp;               AccountNumber = loan.AccountNumber,

&nbsp;               LoanAmount = loan.LoanAmount.ToString("N2"),

&nbsp;               LoanAmountInWords = ConvertAmountToWords(loan.LoanAmount),

&nbsp;               LoanPurpose = loan.LoanPurpose,

&nbsp;               TenureMonths = loan.TenureMonths.ToString(),

&nbsp;               InterestRate = loan.InterestRate.ToString("N2"),

&nbsp;               MonthlyEMI = monthlyEMI.ToString("N2"),

&nbsp;               ApplicationDate = loan.ApplicationDate.ToString("dd-MMM-yyyy"),

&nbsp;               BranchName = loan.BranchName,

&nbsp;               BranchCode = loan.BranchCode,

&nbsp;               ContactNumber = loan.ContactNumber,

&nbsp;               Email = loan.Email,



&nbsp;               // Static fields

&nbsp;               BankName = bankName,

&nbsp;               BankAddress = bankAddress,

&nbsp;               BankPhone = bankPhone,

&nbsp;               BankEmail = bankEmail,

&nbsp;               BankWebsite = bankWebsite,



&nbsp;               // Generated fields

&nbsp;               ReferenceNumber = $"LOAN/{DateTime.Now.Year}/{loanId:D6}",

&nbsp;               GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),

&nbsp;               GeneratedBy = "System" // Or current user

&nbsp;           };

&nbsp;       }



&nbsp;       private async Task<object> GetPromissoryNoteDataAsync(long loanId)

&nbsp;       {

&nbsp;           // Similar to LoanProposal but with different fields

&nbsp;           var loan = await \_context.Loans.FirstOrDefaultAsync(l => l.Id == loanId);

&nbsp;           

&nbsp;           if (loan == null)

&nbsp;               throw new NotFoundException($"Loan not found with ID: {loanId}");



&nbsp;           // Return appropriate DTO for Promissory Note

&nbsp;           return new

&nbsp;           {

&nbsp;               ApplicantName = loan.ApplicantName,

&nbsp;               LoanAmount = loan.LoanAmount.ToString("N2"),

&nbsp;               LoanAmountInWords = ConvertAmountToWords(loan.LoanAmount),

&nbsp;               InterestRate = loan.InterestRate.ToString("N2"),

&nbsp;               RepaymentDate = loan.ApplicationDate.AddMonths(loan.TenureMonths).ToString("dd-MMM-yyyy"),

&nbsp;               ReferenceNumber = $"PN/{DateTime.Now.Year}/{loanId:D6}",

&nbsp;               GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),

&nbsp;               BankName = \_configuration\["BankInfo:Name"]

&nbsp;           };

&nbsp;       }



&nbsp;       private async Task<object> GetLetterOfContinuityDataAsync(long loanId)

&nbsp;       {

&nbsp;           // Implement for Letter of Continuity

&nbsp;           var loan = await \_context.Loans.FirstOrDefaultAsync(l => l.Id == loanId);

&nbsp;           

&nbsp;           if (loan == null)

&nbsp;               throw new NotFoundException($"Loan not found with ID: {loanId}");



&nbsp;           return new

&nbsp;           {

&nbsp;               ApplicantName = loan.ApplicantName,

&nbsp;               AccountNumber = loan.AccountNumber,

&nbsp;               LoanAmount = loan.LoanAmount.ToString("N2"),

&nbsp;               ReferenceNumber = $"LOC/{DateTime.Now.Year}/{loanId:D6}",

&nbsp;               GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),

&nbsp;               BankName = \_configuration\["BankInfo:Name"]

&nbsp;           };

&nbsp;       }



&nbsp;       private string GetTemplateFileName(DocumentType documentType)

&nbsp;       {

&nbsp;           return documentType switch

&nbsp;           {

&nbsp;               DocumentType.LoanProposal => "LoanProposal.docx",

&nbsp;               DocumentType.PromissoryNote => "PromissoryNote.docx",

&nbsp;               DocumentType.LetterOfContinuity => "LetterOfContinuity.docx",

&nbsp;               DocumentType.LetterOfRevival => "LetterOfRevival.docx",

&nbsp;               DocumentType.LetterOfArrangement => "LetterOfArrangement.docx",

&nbsp;               DocumentType.StandingOrder => "StandingOrder.docx",

&nbsp;               DocumentType.LetterOfIndemnity => "LetterOfIndemnity.docx",

&nbsp;               DocumentType.LetterOfLien => "LetterOfLien.docx",

&nbsp;               DocumentType.PersonalGuarantee => "PersonalGuarantee.docx",

&nbsp;               DocumentType.UDC => "UDC.docx",

&nbsp;               \_ => throw new ArgumentException($"Unknown document type: {documentType}")

&nbsp;           };

&nbsp;       }



&nbsp;       private decimal CalculateEMI(decimal principal, decimal annualRate, int tenureMonths)

&nbsp;       {

&nbsp;           if (annualRate == 0) return principal / tenureMonths;



&nbsp;           decimal monthlyRate = annualRate / (12 \* 100);

&nbsp;           decimal emi = principal \* monthlyRate \* 

&nbsp;                        (decimal)Math.Pow((double)(1 + monthlyRate), tenureMonths) /

&nbsp;                        ((decimal)Math.Pow((double)(1 + monthlyRate), tenureMonths) - 1);



&nbsp;           return Math.Round(emi, 2);

&nbsp;       }



&nbsp;       private string ConvertAmountToWords(decimal amount)

&nbsp;       {

&nbsp;           // Implement number to words conversion

&nbsp;           // For simplicity, using a basic implementation

&nbsp;           // You can use a library like Humanizer for better results

&nbsp;           

&nbsp;           var culture = new CultureInfo("en-IN");

&nbsp;           return $"{amount:N2} Taka Only"; // Simplified version

&nbsp;       }

&nbsp;   }



&nbsp;   public class NotFoundException : Exception

&nbsp;   {

&nbsp;       public NotFoundException(string message) : base(message) { }

&nbsp;   }

}



STEP 6: Controller

Controllers/DocumentController.cs

csharpusing Microsoft.AspNetCore.Mvc;

using BankingDocumentAPI.Models.DTOs;

using BankingDocumentAPI.Models.Enums;

using BankingDocumentAPI.Services;



namespace BankingDocumentAPI.Controllers

{

&nbsp;   \[ApiController]

&nbsp;   \[Route("api/\[controller]")]

&nbsp;   public class DocumentController : ControllerBase

&nbsp;   {

&nbsp;       private readonly IDocumentGenerationService \_documentService;

&nbsp;       private readonly ILogger<DocumentController> \_logger;



&nbsp;       public DocumentController(

&nbsp;           IDocumentGenerationService documentService,

&nbsp;           ILogger<DocumentController> logger)

&nbsp;       {

&nbsp;           \_documentService = documentService;

&nbsp;           \_logger = logger;

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Generate a single document

&nbsp;       /// </summary>

&nbsp;       \[HttpPost("generate")]

&nbsp;       \[ProducesResponseType(typeof(FileContentResult), 200)]

&nbsp;       \[ProducesResponseType(400)]

&nbsp;       \[ProducesResponseType(404)]

&nbsp;       \[ProducesResponseType(500)]

&nbsp;       public async Task<IActionResult> GenerateDocument(\[FromBody] DocumentRequest request)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               if (request.LoanId <= 0)

&nbsp;               {

&nbsp;                   return BadRequest("Invalid Loan ID");

&nbsp;               }



&nbsp;               var documentBytes = await \_documentService.GenerateDocumentAsync(request);



&nbsp;               string fileName = $"{request.DocumentType}\_{request.LoanId}\_{DateTime.Now:yyyyMMddHHmmss}";

&nbsp;               string contentType = request.OutputFormat == OutputFormat.PDF

&nbsp;                   ? "application/pdf"

&nbsp;                   : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

&nbsp;               string extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";



&nbsp;               return File(documentBytes, contentType, $"{fileName}.{extension}");

&nbsp;           }

&nbsp;           catch (NotFoundException ex)

&nbsp;           {

&nbsp;               \_logger.LogWarning(ex.Message);

&nbsp;               return NotFound(new { error = ex.Message });

&nbsp;           }

&nbsp;           catch (Exception ex)

&nbsp;           {

&nbsp;               \_logger.LogError(ex, "Error generating document");

&nbsp;               return StatusCode(500, new { error = "Failed to generate document. Please try again." });

&nbsp;           }

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Generate document by direct URL parameters

&nbsp;       /// </summary>

&nbsp;       \[HttpGet("download/{loanId}/{documentType}")]

&nbsp;       public async Task<IActionResult> DownloadDocument(

&nbsp;           long loanId,

&nbsp;           DocumentType documentType,

&nbsp;           \[FromQuery] OutputFormat format = OutputFormat.PDF)

&nbsp;       {

&nbsp;           var request = new DocumentRequest

&nbsp;           {

&nbsp;               LoanId = loanId,

&nbsp;               DocumentType = documentType,

&nbsp;               OutputFormat = format

&nbsp;           };



&nbsp;           return await GenerateDocument(request);

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Generate multiple documents in a ZIP file

&nbsp;       /// </summary>

&nbsp;       \[HttpPost("batch-generate")]

&nbsp;       public async Task<IActionResult> GenerateBatchDocuments(\[FromBody] List<DocumentRequest> requests)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               if (requests == null || !requests.Any())

&nbsp;               {

&nbsp;                   return BadRequest("No documents requested");

&nbsp;               }



&nbsp;               using (var memoryStream = new MemoryStream())

&nbsp;               {

&nbsp;                   using (var archive = new System.IO.Compression.ZipArchive(memoryStream, System.IO.Compression.ZipArchiveMode.Create, true))

&nbsp;                   {

&nbsp;                       foreach (var request in requests)

&nbsp;                       {

&nbsp;                           var documentBytes = await \_documentService.GenerateDocumentAsync(request);

&nbsp;                           var extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";

&nbsp;                           var fileName = $"{request.DocumentType}\_{request.LoanId}.{extension}";



&nbsp;                           var entry = archive.CreateEntry(fileName);

&nbsp;                           using (var entryStream = entry.Open())

&nbsp;                           {

&nbsp;                               await entryStream.WriteAsync(documentBytes, 0, documentBytes.Length);

&nbsp;                           }

&nbsp;                       }

&nbsp;                   }



&nbsp;                   memoryStream.Position = 0;

&nbsp;                   return File(memoryStream.ToArray(),

&nbsp;                       "application/zip",

&nbsp;                       $"BankDocuments\_{DateTime.Now:yyyyMMddHHmmss}.zip");

&nbsp;               }

&nbsp;           }

&nbsp;           catch (Exception ex)

&nbsp;           {

&nbsp;               \_logger.LogError(ex, "Error generating batch documents");

&nbsp;               return StatusCode(500, new { error = "Failed to generate documents" });

&nbsp;           }

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Get list of available document types

&nbsp;       /// </summary>

&nbsp;       \[HttpGet("types")]

&nbsp;       public IActionResult GetDocumentTypes()

&nbsp;       {

&nbsp;           var types = Enum.GetValues(typeof(DocumentType))

&nbsp;               .Cast<DocumentType>()

&nbsp;               .Select(t => new

&nbsp;               {

&nbsp;                   Id = (int)t,

&nbsp;                   Name = t.ToString()

&nbsp;               });



&nbsp;           return Ok(types);

&nbsp;       }

&nbsp;   }

}



STEP 7: Configuration

appsettings.json

json{

&nbsp; "Logging": {

&nbsp;   "LogLevel": {

&nbsp;     "Default": "Information",

&nbsp;     "Microsoft.AspNetCore": "Warning"

&nbsp;   }

&nbsp; },

&nbsp; "AllowedHosts": "\*",

&nbsp; "ConnectionStrings": {

&nbsp;   "DefaultConnection": "Server=localhost;Database=BankingDB;User Id=sa;Password=YourPassword;TrustServerCertificate=True;"

&nbsp; },

&nbsp; "BankInfo": {

&nbsp;   "Name": "ABC Bank Limited",

&nbsp;   "Address": "123 Banking Street, Financial District, Dhaka-1000",

&nbsp;   "Phone": "+880-2-12345678",

&nbsp;   "Email": "info@abcbank.com",

&nbsp;   "Website": "www.abcbank.com"

&nbsp; }

}

Program.cs

csharpusing BankingDocumentAPI.Data;

using BankingDocumentAPI.Services;

using Microsoft.EntityFrameworkCore;



var builder = WebApplication.CreateBuilder(args);



// Add services to the container

builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();

builder.Services.AddSwaggerGen();



// Database

builder.Services.AddDbContext<ApplicationDbContext>(options =>

&nbsp;   options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));



// Document Service

builder.Services.AddScoped<IDocumentGenerationService, DocumentGenerationService>();



// CORS

builder.Services.AddCors(options =>

{

&nbsp;   options.AddPolicy("AllowAll",

&nbsp;       builder =>

&nbsp;       {

&nbsp;           builder.AllowAnyOrigin()

&nbsp;                  .AllowAnyMethod()

&nbsp;                  .AllowAnyHeader();

&nbsp;       });

});



var app = builder.Build();



// Configure the HTTP request pipeline

if (app.Environment.IsDevelopment())

{

&nbsp;   app.UseSwagger();

&nbsp;   app.UseSwaggerUI();

}



app.UseHttpsRedirection();

app.UseCors("AllowAll");

app.UseAuthorization();

app.MapControllers();



app.Run();

```



---



\### \*\*STEP 8: Create Word Templates\*\*



\*\*Create `Templates/LoanProposal.docx`:\*\*



Open Microsoft Word and create:

```

&nbsp;                       \[Insert Logo Image Here]

&nbsp;                       

&nbsp;                       {BankName}

&nbsp;                   {BankAddress}

&nbsp;           Phone: {BankPhone} | Email: {BankEmail}

&nbsp;                   Website: {BankWebsite}



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━



&nbsp;                   LOAN PROPOSAL



Reference No: {ReferenceNumber}

Date: {GeneratedDate}



Dear Mr./Ms. {ApplicantName},



Subject: Loan Proposal - Rs. {LoanAmount}



We are pleased to present the following loan proposal based on your application:



APPLICANT DETAILS:

Name                : {ApplicantName}

Father's Name       : {FatherName}

Address             : {Address}

Account Number      : {AccountNumber}

Contact Number      : {ContactNumber}

Email               : {Email}



LOAN DETAILS:

Loan Amount         : Rs. {LoanAmount} ({LoanAmountInWords})

Purpose             : {LoanPurpose}

Tenure              : {TenureMonths} months

Interest Rate       : {InterestRate}% per annum

Monthly EMI         : Rs. {MonthlyEMI}

Application Date    : {ApplicationDate}



BRANCH DETAILS:

Branch Name         : {BranchName}

Branch Code         : {BranchCode}



\[Add your terms and conditions here]



This proposal is valid for 30 days from the date of issue.



For any queries, please contact our branch.





Authorized Signatory                        Bank Seal

\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_                       \[SEAL]



Generated by: {GeneratedBy}

Generation Date: {GeneratedDate}



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Save as: Templates/LoanProposal.docx

Create similar templates for:



PromissoryNote.docx

LetterOfContinuity.docx

LetterOfRevival.docx

etc.





STEP 9: Test the API

Run the application:

bashdotnet run

Test with Swagger: Navigate to https://localhost:7xxx/swagger

Test with Postman:

httpPOST https://localhost:7xxx/api/document/generate

Content-Type: application/json



{

&nbsp;   "loanId": 1,

&nbsp;   "documentType": 1,

&nbsp;   "outputFormat": 2

}

```



\*\*Or direct download:\*\*

```

GET https://localhost:7xxx/api/document/download/1/LoanProposal?format=PDF



STEP 10: Add Sample Data (Optional)

csharp// Create a migration to add sample data

public class SeedData

{

&nbsp;   public static void Initialize(ApplicationDbContext context)

&nbsp;   {

&nbsp;       if (!context.Loans.Any())

&nbsp;       {

&nbsp;           context.Loans.AddRange(

&nbsp;               new Loan

&nbsp;               {

&nbsp;                   ApplicantName = "John Doe",

&nbsp;                   FatherName = "Robert Doe",

&nbsp;                   Address = "456 Customer Street, Dhaka",

&nbsp;                   AccountNumber = "1234567890",

&nbsp;                   LoanAmount = 500000,

&nbsp;                   LoanPurpose = "Home Renovation",

&nbsp;                   TenureMonths = 60,

&nbsp;                   InterestRate = 8.5m,

&nbsp;                   ApplicationDate = DateTime.Now,

&nbsp;                   BranchName = "Main Branch",

&nbsp;                   BranchCode = "001",

&nbsp;                   ContactNumber = "+880-1712345678",

&nbsp;                   Email = "john.doe@email.com",

&nbsp;                   CreatedDate = DateTime.Now,

&nbsp;                   CreatedBy = "Admin"

&nbsp;               }

&nbsp;           );



&nbsp;           context.SaveChanges();

&nbsp;       }

&nbsp;   }

}



// Call in Program.cs after app.Build()

using (var scope = app.Services.CreateScope())

{

&nbsp;   var context = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();

&nbsp;   context.Database.Migrate();

&nbsp;   SeedData.Initialize(context);

}



Complete! 🎉

Your template-based document generation system is ready!

Next steps:



Create all Word templates with placeholders

Test each document type

Add authentication/authorization

Add audit logging

Deploy to production



Would you like me to show you how to:



Add more complex templates with tables?

Implement template versioning?

Add digital signatures?

Store templates in cloud storage?

give a sample project to do that with one demo .doc and one .pdf using mock api.. leter i will connect real api1:49 PMPerfect! I'll create a complete working sample project with mock data, one Word template, and one PDF template.



Complete Sample Project with Mock API

STEP 1: Create Project

bash# Create project

dotnet new webapi -n BankDocumentDemo

cd BankDocumentDemo



\# Install packages

dotnet add package Syncfusion.DocIO.Net.Core

dotnet add package Syncfusion.DocIORenderer.Net.Core

dotnet add package Syncfusion.Pdf.Net.Core

```



---



\### \*\*STEP 2: Project Structure\*\*

```

BankDocumentDemo/

├── Controllers/

│   └── DocumentController.cs

├── Services/

│   ├── IDocumentService.cs

│   └── DocumentService.cs

├── Models/

│   ├── DocumentRequest.cs

│   ├── LoanData.cs

│   └── Enums.cs

├── MockData/

│   └── LoanMockData.cs

├── Templates/

│   ├── LoanProposal.docx       ← You'll create this

│   └── PromissoryNote.docx     ← You'll create this

├── appsettings.json

├── Program.cs

└── BankDocumentDemo.csproj



STEP 3: Create Models

Models/Enums.cs

csharpnamespace BankDocumentDemo.Models

{

&nbsp;   public enum DocumentType

&nbsp;   {

&nbsp;       LoanProposal = 1,

&nbsp;       PromissoryNote = 2

&nbsp;   }



&nbsp;   public enum OutputFormat

&nbsp;   {

&nbsp;       DOCX = 1,

&nbsp;       PDF = 2

&nbsp;   }

}

Models/DocumentRequest.cs

csharpnamespace BankDocumentDemo.Models

{

&nbsp;   public class DocumentRequest

&nbsp;   {

&nbsp;       public int LoanId { get; set; }

&nbsp;       public DocumentType DocumentType { get; set; }

&nbsp;       public OutputFormat OutputFormat { get; set; } = OutputFormat.PDF;

&nbsp;   }

}

Models/LoanData.cs

csharpnamespace BankDocumentDemo.Models

{

&nbsp;   public class LoanData

&nbsp;   {

&nbsp;       // Applicant Information

&nbsp;       public string ApplicantName { get; set; }

&nbsp;       public string FatherName { get; set; }

&nbsp;       public string Address { get; set; }

&nbsp;       public string ContactNumber { get; set; }

&nbsp;       public string Email { get; set; }

&nbsp;       public string NationalID { get; set; }



&nbsp;       // Loan Information

&nbsp;       public string AccountNumber { get; set; }

&nbsp;       public string LoanAmount { get; set; }

&nbsp;       public string LoanAmountInWords { get; set; }

&nbsp;       public string LoanPurpose { get; set; }

&nbsp;       public string TenureMonths { get; set; }

&nbsp;       public string InterestRate { get; set; }

&nbsp;       public string MonthlyEMI { get; set; }

&nbsp;       public string TotalPayable { get; set; }

&nbsp;       public string ProcessingFee { get; set; }



&nbsp;       // Dates

&nbsp;       public string ApplicationDate { get; set; }

&nbsp;       public string DisbursementDate { get; set; }

&nbsp;       public string MaturityDate { get; set; }



&nbsp;       // Bank Information

&nbsp;       public string BankName { get; set; }

&nbsp;       public string BankAddress { get; set; }

&nbsp;       public string BankPhone { get; set; }

&nbsp;       public string BankEmail { get; set; }

&nbsp;       public string BankWebsite { get; set; }



&nbsp;       // Branch Information

&nbsp;       public string BranchName { get; set; }

&nbsp;       public string BranchCode { get; set; }

&nbsp;       public string BranchAddress { get; set; }

&nbsp;       public string BranchManager { get; set; }



&nbsp;       // Document Information

&nbsp;       public string ReferenceNumber { get; set; }

&nbsp;       public string GeneratedDate { get; set; }

&nbsp;       public string GeneratedTime { get; set; }

&nbsp;       public string GeneratedBy { get; set; }



&nbsp;       // Additional Fields for Promissory Note

&nbsp;       public string WitnessName1 { get; set; }

&nbsp;       public string WitnessAddress1 { get; set; }

&nbsp;       public string WitnessName2 { get; set; }

&nbsp;       public string WitnessAddress2 { get; set; }

&nbsp;   }

}



STEP 4: Create Mock Data

MockData/LoanMockData.cs

csharpnamespace BankDocumentDemo.MockData

{

&nbsp;   public static class LoanMockData

&nbsp;   {

&nbsp;       private static readonly Dictionary<int, LoanDataModel> \_mockLoans = new()

&nbsp;       {

&nbsp;           {

&nbsp;               1, new LoanDataModel

&nbsp;               {

&nbsp;                   LoanId = 1,

&nbsp;                   ApplicantName = "Muhammad Rahman",

&nbsp;                   FatherName = "Abdul Karim",

&nbsp;                   Address = "House# 45, Road# 12, Dhanmondi, Dhaka-1209",

&nbsp;                   ContactNumber = "+880-1712-345678",

&nbsp;                   Email = "m.rahman@email.com",

&nbsp;                   NationalID = "1234567890123",

&nbsp;                   AccountNumber = "ACC-2024-001234",

&nbsp;                   LoanAmount = 500000m,

&nbsp;                   LoanPurpose = "Business Expansion",

&nbsp;                   TenureMonths = 36,

&nbsp;                   InterestRate = 9.5m,

&nbsp;                   ProcessingFee = 5000m,

&nbsp;                   BranchName = "Dhanmondi Branch",

&nbsp;                   BranchCode = "DHK-001"

&nbsp;               }

&nbsp;           },

&nbsp;           {

&nbsp;               2, new LoanDataModel

&nbsp;               {

&nbsp;                   LoanId = 2,

&nbsp;                   ApplicantName = "Fatima Begum",

&nbsp;                   FatherName = "Hassan Ali",

&nbsp;                   Address = "Flat# 3B, Green View Apartment, Gulshan-2, Dhaka-1212",

&nbsp;                   ContactNumber = "+880-1987-654321",

&nbsp;                   Email = "fatima.begum@email.com",

&nbsp;                   NationalID = "9876543210987",

&nbsp;                   AccountNumber = "ACC-2024-005678",

&nbsp;                   LoanAmount = 1000000m,

&nbsp;                   LoanPurpose = "Home Renovation",

&nbsp;                   TenureMonths = 60,

&nbsp;                   InterestRate = 8.75m,

&nbsp;                   ProcessingFee = 10000m,

&nbsp;                   BranchName = "Gulshan Branch",

&nbsp;                   BranchCode = "DHK-002"

&nbsp;               }

&nbsp;           },

&nbsp;           {

&nbsp;               3, new LoanDataModel

&nbsp;               {

&nbsp;                   LoanId = 3,

&nbsp;                   ApplicantName = "Kamal Hossain",

&nbsp;                   FatherName = "Jamal Uddin",

&nbsp;                   Address = "Village: Rampur, Post: Savar, Dhaka-1340",

&nbsp;                   ContactNumber = "+880-1555-123456",

&nbsp;                   Email = "kamal.hossain@email.com",

&nbsp;                   NationalID = "5555666677778888",

&nbsp;                   AccountNumber = "ACC-2024-009012",

&nbsp;                   LoanAmount = 250000m,

&nbsp;                   LoanPurpose = "Agriculture Equipment",

&nbsp;                   TenureMonths = 24,

&nbsp;                   InterestRate = 7.5m,

&nbsp;                   ProcessingFee = 2500m,

&nbsp;                   BranchName = "Savar Branch",

&nbsp;                   BranchCode = "DHK-003"

&nbsp;               }

&nbsp;           }

&nbsp;       };



&nbsp;       public static LoanData GetLoanData(int loanId)

&nbsp;       {

&nbsp;           if (!\_mockLoans.ContainsKey(loanId))

&nbsp;           {

&nbsp;               throw new KeyNotFoundException($"Loan with ID {loanId} not found");

&nbsp;           }



&nbsp;           var loan = \_mockLoans\[loanId];

&nbsp;           

&nbsp;           // Calculate derived fields

&nbsp;           decimal monthlyEMI = CalculateEMI(loan.LoanAmount, loan.InterestRate, loan.TenureMonths);

&nbsp;           decimal totalPayable = monthlyEMI \* loan.TenureMonths;



&nbsp;           DateTime applicationDate = DateTime.Now.AddDays(-10);

&nbsp;           DateTime disbursementDate = DateTime.Now;

&nbsp;           DateTime maturityDate = disbursementDate.AddMonths(loan.TenureMonths);



&nbsp;           return new LoanData

&nbsp;           {

&nbsp;               // Applicant Information

&nbsp;               ApplicantName = loan.ApplicantName,

&nbsp;               FatherName = loan.FatherName,

&nbsp;               Address = loan.Address,

&nbsp;               ContactNumber = loan.ContactNumber,

&nbsp;               Email = loan.Email,

&nbsp;               NationalID = loan.NationalID,



&nbsp;               // Loan Information

&nbsp;               AccountNumber = loan.AccountNumber,

&nbsp;               LoanAmount = $"{loan.LoanAmount:N2}",

&nbsp;               LoanAmountInWords = ConvertToWords(loan.LoanAmount),

&nbsp;               LoanPurpose = loan.LoanPurpose,

&nbsp;               TenureMonths = loan.TenureMonths.ToString(),

&nbsp;               InterestRate = $"{loan.InterestRate:N2}",

&nbsp;               MonthlyEMI = $"{monthlyEMI:N2}",

&nbsp;               TotalPayable = $"{totalPayable:N2}",

&nbsp;               ProcessingFee = $"{loan.ProcessingFee:N2}",



&nbsp;               // Dates

&nbsp;               ApplicationDate = applicationDate.ToString("dd MMMM yyyy"),

&nbsp;               DisbursementDate = disbursementDate.ToString("dd MMMM yyyy"),

&nbsp;               MaturityDate = maturityDate.ToString("dd MMMM yyyy"),



&nbsp;               // Bank Information (Static)

&nbsp;               BankName = "ABC Bank Limited",

&nbsp;               BankAddress = "Head Office: 123 Motijheel C/A, Dhaka-1000, Bangladesh",

&nbsp;               BankPhone = "+880-2-9559191",

&nbsp;               BankEmail = "info@abcbank.com.bd",

&nbsp;               BankWebsite = "www.abcbank.com.bd",



&nbsp;               // Branch Information

&nbsp;               BranchName = loan.BranchName,

&nbsp;               BranchCode = loan.BranchCode,

&nbsp;               BranchAddress = GetBranchAddress(loan.BranchCode),

&nbsp;               BranchManager = GetBranchManager(loan.BranchCode),



&nbsp;               // Document Information

&nbsp;               ReferenceNumber = $"LOAN/{DateTime.Now.Year}/{loanId:D6}",

&nbsp;               GeneratedDate = DateTime.Now.ToString("dd MMMM yyyy"),

&nbsp;               GeneratedTime = DateTime.Now.ToString("hh:mm tt"),

&nbsp;               GeneratedBy = "System Administrator",



&nbsp;               // Witness Information (for Promissory Note)

&nbsp;               WitnessName1 = "Ahmed Khan",

&nbsp;               WitnessAddress1 = "House# 23, Road# 5, Banani, Dhaka",

&nbsp;               WitnessName2 = "Nasrin Akter",

&nbsp;               WitnessAddress2 = "Flat# 2A, Holding# 67, Mohammadpur, Dhaka"

&nbsp;           };

&nbsp;       }



&nbsp;       public static List<LoanSummary> GetAllLoans()

&nbsp;       {

&nbsp;           return \_mockLoans.Values.Select(l => new LoanSummary

&nbsp;           {

&nbsp;               LoanId = l.LoanId,

&nbsp;               ApplicantName = l.ApplicantName,

&nbsp;               AccountNumber = l.AccountNumber,

&nbsp;               LoanAmount = l.LoanAmount,

&nbsp;               Status = "Active"

&nbsp;           }).ToList();

&nbsp;       }



&nbsp;       private static decimal CalculateEMI(decimal principal, decimal annualRate, int months)

&nbsp;       {

&nbsp;           if (annualRate == 0) return principal / months;



&nbsp;           decimal monthlyRate = annualRate / (12 \* 100);

&nbsp;           decimal emi = principal \* monthlyRate \*

&nbsp;                        (decimal)Math.Pow((double)(1 + monthlyRate), months) /

&nbsp;                        ((decimal)Math.Pow((double)(1 + monthlyRate), months) - 1);



&nbsp;           return Math.Round(emi, 2);

&nbsp;       }



&nbsp;       private static string ConvertToWords(decimal amount)

&nbsp;       {

&nbsp;           // Simplified version - you can enhance this

&nbsp;           string\[] ones = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine" };

&nbsp;           string\[] teens = { "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };

&nbsp;           string\[] tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };



&nbsp;           int intAmount = (int)amount;

&nbsp;           

&nbsp;           if (intAmount == 0) return "Zero Taka Only";



&nbsp;           string words = "";

&nbsp;           

&nbsp;           // Lakhs

&nbsp;           if (intAmount >= 100000)

&nbsp;           {

&nbsp;               words += ConvertToWordsHelper(intAmount / 100000, ones, teens, tens) + " Lakh ";

&nbsp;               intAmount %= 100000;

&nbsp;           }



&nbsp;           // Thousands

&nbsp;           if (intAmount >= 1000)

&nbsp;           {

&nbsp;               words += ConvertToWordsHelper(intAmount / 1000, ones, teens, tens) + " Thousand ";

&nbsp;               intAmount %= 1000;

&nbsp;           }



&nbsp;           // Hundreds

&nbsp;           if (intAmount >= 100)

&nbsp;           {

&nbsp;               words += ones\[intAmount / 100] + " Hundred ";

&nbsp;               intAmount %= 100;

&nbsp;           }



&nbsp;           // Remaining

&nbsp;           if (intAmount > 0)

&nbsp;           {

&nbsp;               words += ConvertToWordsHelper(intAmount, ones, teens, tens);

&nbsp;           }



&nbsp;           return words.Trim() + " Taka Only";

&nbsp;       }



&nbsp;       private static string ConvertToWordsHelper(int number, string\[] ones, string\[] teens, string\[] tens)

&nbsp;       {

&nbsp;           if (number == 0) return "";

&nbsp;           if (number < 10) return ones\[number];

&nbsp;           if (number < 20) return teens\[number - 10];

&nbsp;           

&nbsp;           return tens\[number / 10] + (number % 10 > 0 ? " " + ones\[number % 10] : "");

&nbsp;       }



&nbsp;       private static string GetBranchAddress(string branchCode)

&nbsp;       {

&nbsp;           return branchCode switch

&nbsp;           {

&nbsp;               "DHK-001" => "House# 15, Road# 27, Dhanmondi R/A, Dhaka-1209",

&nbsp;               "DHK-002" => "Plot# 23, Road# 11, Gulshan-2, Dhaka-1212",

&nbsp;               "DHK-003" => "Main Road, Savar Bazar, Savar, Dhaka-1340",

&nbsp;               \_ => "Branch Address"

&nbsp;           };

&nbsp;       }



&nbsp;       private static string GetBranchManager(string branchCode)

&nbsp;       {

&nbsp;           return branchCode switch

&nbsp;           {

&nbsp;               "DHK-001" => "Mr. Khalid Rahman",

&nbsp;               "DHK-002" => "Ms. Aisha Siddique",

&nbsp;               "DHK-003" => "Mr. Rafiq Ahmed",

&nbsp;               \_ => "Branch Manager"

&nbsp;           };

&nbsp;       }



&nbsp;       public class LoanDataModel

&nbsp;       {

&nbsp;           public int LoanId { get; set; }

&nbsp;           public string ApplicantName { get; set; }

&nbsp;           public string FatherName { get; set; }

&nbsp;           public string Address { get; set; }

&nbsp;           public string ContactNumber { get; set; }

&nbsp;           public string Email { get; set; }

&nbsp;           public string NationalID { get; set; }

&nbsp;           public string AccountNumber { get; set; }

&nbsp;           public decimal LoanAmount { get; set; }

&nbsp;           public string LoanPurpose { get; set; }

&nbsp;           public int TenureMonths { get; set; }

&nbsp;           public decimal InterestRate { get; set; }

&nbsp;           public decimal ProcessingFee { get; set; }

&nbsp;           public string BranchName { get; set; }

&nbsp;           public string BranchCode { get; set; }

&nbsp;       }



&nbsp;       public class LoanSummary

&nbsp;       {

&nbsp;           public int LoanId { get; set; }

&nbsp;           public string ApplicantName { get; set; }

&nbsp;           public string AccountNumber { get; set; }

&nbsp;           public decimal LoanAmount { get; set; }

&nbsp;           public string Status { get; set; }

&nbsp;       }

&nbsp;   }

}



STEP 5: Create Document Service

Services/IDocumentService.cs

csharpusing BankDocumentDemo.Models;



namespace BankDocumentDemo.Services

{

&nbsp;   public interface IDocumentService

&nbsp;   {

&nbsp;       Task<byte\[]> GenerateDocumentAsync(DocumentRequest request);

&nbsp;   }

}

Services/DocumentService.cs

csharpusing Syncfusion.DocIO;

using Syncfusion.DocIO.DLS;

using Syncfusion.DocIORenderer;

using Syncfusion.Pdf;

using BankDocumentDemo.Models;

using BankDocumentDemo.MockData;



namespace BankDocumentDemo.Services

{

&nbsp;   public class DocumentService : IDocumentService

&nbsp;   {

&nbsp;       private readonly ILogger<DocumentService> \_logger;

&nbsp;       private readonly string \_templatePath;



&nbsp;       public DocumentService(ILogger<DocumentService> logger)

&nbsp;       {

&nbsp;           \_logger = logger;

&nbsp;           \_templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");

&nbsp;       }



&nbsp;       public async Task<byte\[]> GenerateDocumentAsync(DocumentRequest request)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               \_logger.LogInformation($"Generating {request.DocumentType} for Loan ID: {request.LoanId}");



&nbsp;               // Get mock data

&nbsp;               var loanData = LoanMockData.GetLoanData(request.LoanId);



&nbsp;               // Get template file name

&nbsp;               string templateFileName = GetTemplateFileName(request.DocumentType);

&nbsp;               string templateFullPath = Path.Combine(\_templatePath, templateFileName);



&nbsp;               if (!File.Exists(templateFullPath))

&nbsp;               {

&nbsp;                   throw new FileNotFoundException($"Template not found: {templateFileName} at {templateFullPath}");

&nbsp;               }



&nbsp;               // Generate Word document

&nbsp;               byte\[] wordDocument = await GenerateWordDocumentAsync(templateFullPath, loanData);



&nbsp;               // Convert to PDF if requested

&nbsp;               if (request.OutputFormat == OutputFormat.PDF)

&nbsp;               {

&nbsp;                   return await ConvertWordToPdfAsync(wordDocument);

&nbsp;               }



&nbsp;               return wordDocument;

&nbsp;           }

&nbsp;           catch (Exception ex)

&nbsp;           {

&nbsp;               \_logger.LogError(ex, $"Error generating document for Loan ID: {request.LoanId}");

&nbsp;               throw;

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<byte\[]> GenerateWordDocumentAsync(string templatePath, LoanData data)

&nbsp;       {

&nbsp;           using (FileStream templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))

&nbsp;           {

&nbsp;               // Load the template

&nbsp;               using (WordDocument document = new WordDocument(templateStream, FormatType.Docx))

&nbsp;               {

&nbsp;                   // Replace all placeholders

&nbsp;                   ReplaceBookmarks(document, data);



&nbsp;                   // Save to memory stream

&nbsp;                   using (MemoryStream outputStream = new MemoryStream())

&nbsp;                   {

&nbsp;                       document.Save(outputStream, FormatType.Docx);

&nbsp;                       return await Task.FromResult(outputStream.ToArray());

&nbsp;                   }

&nbsp;               }

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<byte\[]> ConvertWordToPdfAsync(byte\[] wordDocument)

&nbsp;       {

&nbsp;           using (MemoryStream wordStream = new MemoryStream(wordDocument))

&nbsp;           using (WordDocument document = new WordDocument(wordStream, FormatType.Docx))

&nbsp;           {

&nbsp;               // Convert Word to PDF

&nbsp;               using (DocIORenderer renderer = new DocIORenderer())

&nbsp;               {

&nbsp;                   PdfDocument pdfDocument = renderer.ConvertToPDF(document);



&nbsp;                   using (MemoryStream pdfStream = new MemoryStream())

&nbsp;                   {

&nbsp;                       pdfDocument.Save(pdfStream);

&nbsp;                       pdfDocument.Close();

&nbsp;                       return await Task.FromResult(pdfStream.ToArray());

&nbsp;                   }

&nbsp;               }

&nbsp;           }

&nbsp;       }



&nbsp;       private void ReplaceBookmarks(WordDocument document, LoanData data)

&nbsp;       {

&nbsp;           // Get all properties using reflection

&nbsp;           var properties = data.GetType().GetProperties();



&nbsp;           foreach (var property in properties)

&nbsp;           {

&nbsp;               string placeholder = $"{{{property.Name}}}";

&nbsp;               string value = property.GetValue(data)?.ToString() ?? "";



&nbsp;               // Find and replace in the entire document

&nbsp;               document.Replace(placeholder, value, true, true);

&nbsp;           }



&nbsp;           \_logger.LogInformation($"Replaced {properties.Length} placeholders in document");

&nbsp;       }



&nbsp;       private string GetTemplateFileName(DocumentType documentType)

&nbsp;       {

&nbsp;           return documentType switch

&nbsp;           {

&nbsp;               DocumentType.LoanProposal => "LoanProposal.docx",

&nbsp;               DocumentType.PromissoryNote => "PromissoryNote.docx",

&nbsp;               \_ => throw new ArgumentException($"Unknown document type: {documentType}")

&nbsp;           };

&nbsp;       }

&nbsp;   }

}



STEP 6: Create Controller

Controllers/DocumentController.cs

csharpusing Microsoft.AspNetCore.Mvc;

using BankDocumentDemo.Models;

using BankDocumentDemo.Services;

using BankDocumentDemo.MockData;



namespace BankDocumentDemo.Controllers

{

&nbsp;   \[ApiController]

&nbsp;   \[Route("api/\[controller]")]

&nbsp;   public class DocumentController : ControllerBase

&nbsp;   {

&nbsp;       private readonly IDocumentService \_documentService;

&nbsp;       private readonly ILogger<DocumentController> \_logger;



&nbsp;       public DocumentController(

&nbsp;           IDocumentService documentService,

&nbsp;           ILogger<DocumentController> logger)

&nbsp;       {

&nbsp;           \_documentService = documentService;

&nbsp;           \_logger = logger;

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Get all available loans (Mock Data)

&nbsp;       /// </summary>

&nbsp;       \[HttpGet("loans")]

&nbsp;       public IActionResult GetAllLoans()

&nbsp;       {

&nbsp;           var loans = LoanMockData.GetAllLoans();

&nbsp;           return Ok(loans);

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Get loan details by ID (Mock Data)

&nbsp;       /// </summary>

&nbsp;       \[HttpGet("loans/{loanId}")]

&nbsp;       public IActionResult GetLoanById(int loanId)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               var loanData = LoanMockData.GetLoanData(loanId);

&nbsp;               return Ok(loanData);

&nbsp;           }

&nbsp;           catch (KeyNotFoundException ex)

&nbsp;           {

&nbsp;               return NotFound(new { error = ex.Message });

&nbsp;           }

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Generate document (POST method)

&nbsp;       /// </summary>

&nbsp;       \[HttpPost("generate")]

&nbsp;       \[ProducesResponseType(typeof(FileContentResult), 200)]

&nbsp;       \[ProducesResponseType(400)]

&nbsp;       \[ProducesResponseType(404)]

&nbsp;       public async Task<IActionResult> GenerateDocument(\[FromBody] DocumentRequest request)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               var documentBytes = await \_documentService.GenerateDocumentAsync(request);



&nbsp;               string fileName = $"{request.DocumentType}\_{request.LoanId}\_{DateTime.Now:yyyyMMddHHmmss}";

&nbsp;               string contentType = request.OutputFormat == OutputFormat.PDF

&nbsp;                   ? "application/pdf"

&nbsp;                   : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

&nbsp;               string extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";



&nbsp;               return File(documentBytes, contentType, $"{fileName}.{extension}");

&nbsp;           }

&nbsp;           catch (KeyNotFoundException ex)

&nbsp;           {

&nbsp;               return NotFound(new { error = ex.Message });

&nbsp;           }

&nbsp;           catch (FileNotFoundException ex)

&nbsp;           {

&nbsp;               return NotFound(new { error = ex.Message });

&nbsp;           }

&nbsp;           catch (Exception ex)

&nbsp;           {

&nbsp;               \_logger.LogError(ex, "Error generating document");

&nbsp;               return StatusCode(500, new { error = "Failed to generate document", details = ex.Message });

&nbsp;           }

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Generate document (GET method - easier for testing)

&nbsp;       /// </summary>

&nbsp;       \[HttpGet("download/{loanId}/{documentType}")]

&nbsp;       public async Task<IActionResult> DownloadDocument(

&nbsp;           int loanId,

&nbsp;           DocumentType documentType,

&nbsp;           \[FromQuery] OutputFormat format = OutputFormat.PDF)

&nbsp;       {

&nbsp;           var request = new DocumentRequest

&nbsp;           {

&nbsp;               LoanId = loanId,

&nbsp;               DocumentType = documentType,

&nbsp;               OutputFormat = format

&nbsp;           };



&nbsp;           return await GenerateDocument(request);

&nbsp;       }



&nbsp;       /// <summary>

&nbsp;       /// Get available document types

&nbsp;       /// </summary>

&nbsp;       \[HttpGet("types")]

&nbsp;       public IActionResult GetDocumentTypes()

&nbsp;       {

&nbsp;           var types = Enum.GetValues(typeof(DocumentType))

&nbsp;               .Cast<DocumentType>()

&nbsp;               .Select(t => new

&nbsp;               {

&nbsp;                   Id = (int)t,

&nbsp;                   Name = t.ToString()

&nbsp;               });



&nbsp;           return Ok(types);

&nbsp;       }

&nbsp;   }

}



STEP 7: Configure Program.cs

Program.cs

csharpusing BankDocumentDemo.Services;



var builder = WebApplication.CreateBuilder(args);



// Add services to the container

builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();

builder.Services.AddSwaggerGen(c =>

{

&nbsp;   c.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo

&nbsp;   {

&nbsp;       Title = "Bank Document Generation API",

&nbsp;       Version = "v1",

&nbsp;       Description = "Demo API for generating banking documents using templates"

&nbsp;   });

});



// Register Document Service

builder.Services.AddScoped<IDocumentService, DocumentService>();



// Add CORS

builder.Services.AddCors(options =>

{

&nbsp;   options.AddPolicy("AllowAll", builder =>

&nbsp;   {

&nbsp;       builder.AllowAnyOrigin()

&nbsp;              .AllowAnyMethod()

&nbsp;              .AllowAnyHeader();

&nbsp;   });

});



var app = builder.Build();



// Configure the HTTP request pipeline

app.UseSwagger();

app.UseSwaggerUI(c =>

{

&nbsp;   c.SwaggerEndpoint("/swagger/v1/swagger.json", "Bank Document API v1");

&nbsp;   c.RoutePrefix = string.Empty; // Set Swagger UI at root

});



app.UseHttpsRedirection();

app.UseCors("AllowAll");

app.UseAuthorization();

app.MapControllers();



Console.WriteLine("=================================================");

Console.WriteLine("  Bank Document Generation API is running!");

Console.WriteLine("=================================================");

Console.WriteLine("  Swagger UI: https://localhost:7xxx");

Console.WriteLine("  Available Endpoints:");

Console.WriteLine("  - GET  /api/document/loans");

Console.WriteLine("  - GET  /api/document/loans/{id}");

Console.WriteLine("  - POST /api/document/generate");

Console.WriteLine("  - GET  /api/document/download/{loanId}/{type}");

Console.WriteLine("=================================================");



app.Run();

```



---



\### \*\*STEP 8: Create Word Templates\*\*



Create a \*\*Templates\*\* folder in your project root, then create these templates:



\#### \*\*Templates/LoanProposal.docx\*\*



Open Microsoft Word and create this document:

```

&nbsp;                       ABC BANK LIMITED

&nbsp;           Head Office: 123 Motijheel C/A, Dhaka-1000

&nbsp;       Phone: +880-2-9559191 | Email: info@abcbank.com.bd

&nbsp;                   www.abcbank.com.bd



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━



&nbsp;                       LOAN PROPOSAL



Reference No: {ReferenceNumber}

Date: {GeneratedDate}



{BranchName} ({BranchCode})

{BranchAddress}

Branch Manager: {BranchManager}



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━



Dear Mr./Ms. {ApplicantName},



Subject: Loan Proposal for BDT {LoanAmount}



We are pleased to submit the following loan proposal based on your 

application dated {ApplicationDate}.



APPLICANT DETAILS

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Full Name          : {ApplicantName}

Father's Name      : {FatherName}

Present Address    : {Address}

National ID        : {NationalID}

Contact Number     : {ContactNumber}

Email Address      : {Email}

Account Number     : {AccountNumber}



LOAN DETAILS

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Loan Amount        : BDT {LoanAmount}

Amount in Words    : {LoanAmountInWords}

Loan Purpose       : {LoanPurpose}

Loan Tenure        : {TenureMonths} months

Interest Rate      : {InterestRate}% per annum (Reducing Balance)

Monthly EMI        : BDT {MonthlyEMI}

Total Payable      : BDT {TotalPayable}

Processing Fee     : BDT {ProcessingFee}



IMPORTANT DATES

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Application Date   : {ApplicationDate}

Disbursement Date  : {DisbursementDate}

Maturity Date      : {MaturityDate}



TERMS AND CONDITIONS:



1\. The loan will be disbursed upon submission of all required documents.

2\. Monthly EMI should be paid by the 5th of each month.

3\. Late payment will attract penalty charges as per bank policy.

4\. The bank reserves the right to recall the loan in case of default.

5\. All applicable taxes and charges will be borne by the borrower.



This proposal is valid for 30 days from the date of issue.



For any queries, please contact the branch during banking hours.





Best Regards,





\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_                    \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Branch Manager                            Authorized Officer

{BranchManager}                          {BankName}





Generated by: {GeneratedBy}

Date \& Time: {GeneratedDate} at {GeneratedTime}



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

&nbsp;               This is a computer-generated document

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

```



Save as: \*\*Templates/LoanProposal.docx\*\*



---



\#### \*\*Templates/PromissoryNote.docx\*\*

```

&nbsp;                       ABC BANK LIMITED

&nbsp;           Head Office: 123 Motijheel C/A, Dhaka-1000

&nbsp;       Phone: +880-2-9559191 | Email: info@abcbank.com.bd



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━



&nbsp;                   DEMAND PROMISSORY NOTE



Reference No: {ReferenceNumber}

Date: {GeneratedDate}



━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━





On demand, I/We {ApplicantName}, son/daughter of {FatherName}, residing 

at {Address}, do hereby promise to pay {BankName} or order the sum of:



&nbsp;   BDT {LoanAmount} ({LoanAmountInWords})



together with interest at the rate of {InterestRate}% per annum or at 

such other rate as the Bank may determine from time to time.



This amount is payable on or before {MaturityDate}.



LOAN DETAILS:

\- Account Number: {AccountNumber}

\- Loan Amount: BDT {LoanAmount}

\- Monthly EMI: BDT {MonthlyEMI}

\- Tenure: {TenureMonths} months

\- Purpose: {LoanPurpose}



I/We hereby agree to abide by all terms and conditions set forth by the 

Bank and acknowledge that this is a legally binding document.





BORROWER DECLARATION:



I/We confirm that all information provided is true and correct to the 

best of my/our knowledge.





\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_                  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Signature of Borrower                        Date

{ApplicantName}

National ID: {NationalID}





WITNESS 1:                                   WITNESS 2:



\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_                  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Name: {WitnessName1}                        Name: {WitnessName2}

Address: {WitnessAddress1}                  Address: {WitnessAddress2}

Signature: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_                 Signature: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_                       Date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_





FOR BANK USE ONLY:



Verified and Approved by:



\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Authorized Bank Officer

{BranchName}

Date: {GeneratedDate}



\[BANK SEAL]





━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

&nbsp;   This is a legally binding document. Please read carefully.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Save as: Templates/PromissoryNote.docx



STEP 9: Update .csproj

Make sure templates are copied to output:

BankDocumentDemo.csproj

xml<Project Sdk="Microsoft.NET.Sdk.Web">



&nbsp; <PropertyGroup>

&nbsp;   <TargetFramework>net8.0</TargetFramework>

&nbsp;   <Nullable>enable</Nullable>

&nbsp;   <ImplicitUsings>enable</ImplicitUsings>

&nbsp; </PropertyGroup>



&nbsp; <ItemGroup>

&nbsp;   <PackageReference Include="Syncfusion.DocIO.Net.Core" Version="27.1.48" />

&nbsp;   <PackageReference Include="Syncfusion.DocIORenderer.Net.Core" Version="27.1.48" />

&nbsp;   <PackageReference Include="Syncfusion.Pdf.Net.Core" Version="27.1.48" />

&nbsp;   <PackageReference Include="Swashbuckle.AspNetCore" Version="6.4.0" />

&nbsp; </ItemGroup>



&nbsp; <ItemGroup>

&nbsp;   <None Update="Templates\\\*\*\\\*">

&nbsp;     <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>

&nbsp;   </None>

&nbsp; </ItemGroup>



</Project>



STEP 10: Run and Test

bash# Run the application

dotnet run

```



Open browser: `https://localhost:7xxx` (Swagger UI will open)



---



\### \*\*STEP 11: Test the API\*\*



\#### \*\*Test 1: Get All Loans\*\*

```

GET https://localhost:7xxx/api/document/loans

Response:

json\[

&nbsp; {

&nbsp;   "loanId": 1,

&nbsp;   "applicantName": "Muhammad Rahman",

&nbsp;   "accountNumber": "ACC-2024-001234",

&nbsp;   "loanAmount": 500000,

&nbsp;   "status": "Active"

&nbsp; },

&nbsp; ...

]

```



\#### \*\*Test 2: Get Loan Details\*\*

```

GET https://localhost:7xxx/api/document/loans/1

```



\#### \*\*Test 3: Generate Loan Proposal (PDF)\*\*

```

GET https://localhost:7xxx/api/document/download/1/LoanProposal?format=2

```

This will download a PDF file!



\#### \*\*Test 4: Generate Promissory Note (Word)\*\*

```

GET https://localhost:7xxx/api/document/download/1/PromissoryNote?format=1

```

This will download a .docx file!



\#### \*\*Test 5: Using POST method\*\*

```

POST https://localhost:7xxx/api/document/generate

Content-Type: application/json



{

&nbsp; "loanId": 2,

&nbsp; "documentType": 1,

&nbsp; "outputFormat": 2

}



STEP 12: Testing with Postman/Browser

Easy Browser Test:



Open: https://localhost:7xxx

Try endpoint: GET /api/document/download/1/LoanProposal?format=2

Click "Try it out"

Click "Execute"

Download the PDF!





Project Complete! 🎉

What you have:

✅ Mock API with 3 sample loans

✅ 2 Word templates (Loan Proposal \& Promissory Note)

✅ Generate both DOCX and PDF

✅ Full Swagger documentation

✅ Ready to connect to real database later

To connect real database later:

Just replace LoanMockData.GetLoanData() with your actual database calls in DocumentService.cs!



Download the complete project as ZIP? Let me know if you want me to create a downloadable package or explain any part in more detail!yes.Download the complete project as ZIP1:52 PMI cannot directly create and provide ZIP files for download. However, I can help you create this project step-by-step on your machine, or I can provide you with a comprehensive guide to set it up quickly.

Let me give you the fastest way to create this complete project:



Quick Setup Script (PowerShell - Windows)

Save this as CreateBankDocumentDemo.ps1 and run it:

powershell# Create project

Write-Host "Creating Bank Document Demo Project..." -ForegroundColor Green



\# Create main project

dotnet new webapi -n BankDocumentDemo

Set-Location BankDocumentDemo



\# Install packages

Write-Host "Installing NuGet packages..." -ForegroundColor Yellow

dotnet add package Syncfusion.DocIO.Net.Core

dotnet add package Syncfusion.DocIORenderer.Net.Core

dotnet add package Syncfusion.Pdf.Net.Core



\# Create folder structure

Write-Host "Creating folder structure..." -ForegroundColor Yellow

New-Item -ItemType Directory -Path "Models" -Force

New-Item -ItemType Directory -Path "Services" -Force

New-Item -ItemType Directory -Path "MockData" -Force

New-Item -ItemType Directory -Path "Templates" -Force



Write-Host "Project structure created successfully!" -ForegroundColor Green

Write-Host "Next steps:" -ForegroundColor Cyan

Write-Host "1. Copy the provided code files into the project" -ForegroundColor White

Write-Host "2. Create the Word templates in the Templates folder" -ForegroundColor White

Write-Host "3. Run: dotnet run" -ForegroundColor White



Alternative: Manual Step-by-Step (5 Minutes)

Step 1: Create Project

bashdotnet new webapi -n BankDocumentDemo

cd BankDocumentDemo

Step 2: Install Packages

bashdotnet add package Syncfusion.DocIO.Net.Core

dotnet add package Syncfusion.DocIORenderer.Net.Core

dotnet add package Syncfusion.Pdf.Net.Core

Step 3: Create Folders

bashmkdir Models Services MockData Templates

Step 4: Create Files

Now I'll give you ALL the code in ready-to-copy format:



FILE 1: Models/Enums.cs

csharpnamespace BankDocumentDemo.Models

{

&nbsp;   public enum DocumentType

&nbsp;   {

&nbsp;       LoanProposal = 1,

&nbsp;       PromissoryNote = 2

&nbsp;   }



&nbsp;   public enum OutputFormat

&nbsp;   {

&nbsp;       DOCX = 1,

&nbsp;       PDF = 2

&nbsp;   }

}



FILE 2: Models/DocumentRequest.cs

csharpnamespace BankDocumentDemo.Models

{

&nbsp;   public class DocumentRequest

&nbsp;   {

&nbsp;       public int LoanId { get; set; }

&nbsp;       public DocumentType DocumentType { get; set; }

&nbsp;       public OutputFormat OutputFormat { get; set; } = OutputFormat.PDF;

&nbsp;   }

}



FILE 3: Models/LoanData.cs

csharpnamespace BankDocumentDemo.Models

{

&nbsp;   public class LoanData

&nbsp;   {

&nbsp;       public string ApplicantName { get; set; }

&nbsp;       public string FatherName { get; set; }

&nbsp;       public string Address { get; set; }

&nbsp;       public string ContactNumber { get; set; }

&nbsp;       public string Email { get; set; }

&nbsp;       public string NationalID { get; set; }

&nbsp;       public string AccountNumber { get; set; }

&nbsp;       public string LoanAmount { get; set; }

&nbsp;       public string LoanAmountInWords { get; set; }

&nbsp;       public string LoanPurpose { get; set; }

&nbsp;       public string TenureMonths { get; set; }

&nbsp;       public string InterestRate { get; set; }

&nbsp;       public string MonthlyEMI { get; set; }

&nbsp;       public string TotalPayable { get; set; }

&nbsp;       public string ProcessingFee { get; set; }

&nbsp;       public string ApplicationDate { get; set; }

&nbsp;       public string DisbursementDate { get; set; }

&nbsp;       public string MaturityDate { get; set; }

&nbsp;       public string BankName { get; set; }

&nbsp;       public string BankAddress { get; set; }

&nbsp;       public string BankPhone { get; set; }

&nbsp;       public string BankEmail { get; set; }

&nbsp;       public string BankWebsite { get; set; }

&nbsp;       public string BranchName { get; set; }

&nbsp;       public string BranchCode { get; set; }

&nbsp;       public string BranchAddress { get; set; }

&nbsp;       public string BranchManager { get; set; }

&nbsp;       public string ReferenceNumber { get; set; }

&nbsp;       public string GeneratedDate { get; set; }

&nbsp;       public string GeneratedTime { get; set; }

&nbsp;       public string GeneratedBy { get; set; }

&nbsp;       public string WitnessName1 { get; set; }

&nbsp;       public string WitnessAddress1 { get; set; }

&nbsp;       public string WitnessName2 { get; set; }

&nbsp;       public string WitnessAddress2 { get; set; }

&nbsp;   }

}



FILE 4: MockData/LoanMockData.cs

<details>

<summary>Click to expand (Large file - 300+ lines)</summary>

```csharp

namespace BankDocumentDemo.MockData

{

&nbsp;   public static class LoanMockData

&nbsp;   {

&nbsp;       private static readonly Dictionary<int, LoanDataModel> \_mockLoans = new()

&nbsp;       {

&nbsp;           { 1, new LoanDataModel

&nbsp;               {

&nbsp;                   LoanId = 1,

&nbsp;                   ApplicantName = "Muhammad Rahman",

&nbsp;                   FatherName = "Abdul Karim",

&nbsp;                   Address = "House# 45, Road# 12, Dhanmondi, Dhaka-1209",

&nbsp;                   ContactNumber = "+880-1712-345678",

&nbsp;                   Email = "m.rahman@email.com",

&nbsp;                   NationalID = "1234567890123",

&nbsp;                   AccountNumber = "ACC-2024-001234",

&nbsp;                   LoanAmount = 500000m,

&nbsp;                   LoanPurpose = "Business Expansion",

&nbsp;                   TenureMonths = 36,

&nbsp;                   InterestRate = 9.5m,

&nbsp;                   ProcessingFee = 5000m,

&nbsp;                   BranchName = "Dhanmondi Branch",

&nbsp;                   BranchCode = "DHK-001"

&nbsp;               }

&nbsp;           },

&nbsp;           { 2, new LoanDataModel

&nbsp;               {

&nbsp;                   LoanId = 2,

&nbsp;                   ApplicantName = "Fatima Begum",

&nbsp;                   FatherName = "Hassan Ali",

&nbsp;                   Address = "Flat# 3B, Green View Apartment, Gulshan-2, Dhaka-1212",

&nbsp;                   ContactNumber = "+880-1987-654321",

&nbsp;                   Email = "fatima.begum@email.com",

&nbsp;                   NationalID = "9876543210987",

&nbsp;                   AccountNumber = "ACC-2024-005678",

&nbsp;                   LoanAmount = 1000000m,

&nbsp;                   LoanPurpose = "Home Renovation",

&nbsp;                   TenureMonths = 60,

&nbsp;                   InterestRate = 8.75m,

&nbsp;                   ProcessingFee = 10000m,

&nbsp;                   BranchName = "Gulshan Branch",

&nbsp;                   BranchCode = "DHK-002"

&nbsp;               }

&nbsp;           },

&nbsp;           { 3, new LoanDataModel

&nbsp;               {

&nbsp;                   LoanId = 3,

&nbsp;                   ApplicantName = "Kamal Hossain",

&nbsp;                   FatherName = "Jamal Uddin",

&nbsp;                   Address = "Village: Rampur, Post: Savar, Dhaka-1340",

&nbsp;                   ContactNumber = "+880-1555-123456",

&nbsp;                   Email = "kamal.hossain@email.com",

&nbsp;                   NationalID = "5555666677778888",

&nbsp;                   AccountNumber = "ACC-2024-009012",

&nbsp;                   LoanAmount = 250000m,

&nbsp;                   LoanPurpose = "Agriculture Equipment",

&nbsp;                   TenureMonths = 24,

&nbsp;                   InterestRate = 7.5m,

&nbsp;                   ProcessingFee = 2500m,

&nbsp;                   BranchName = "Savar Branch",

&nbsp;                   BranchCode = "DHK-003"

&nbsp;               }

&nbsp;           }

&nbsp;       };

&nbsp;   public static Models.LoanData GetLoanData(int loanId)

&nbsp;   {

&nbsp;       if (!\_mockLoans.ContainsKey(loanId))

&nbsp;           throw new KeyNotFoundException($"Loan with ID {loanId} not found");



&nbsp;       var loan = \_mockLoans\[loanId];

&nbsp;       decimal monthlyEMI = CalculateEMI(loan.LoanAmount, loan.InterestRate, loan.TenureMonths);

&nbsp;       decimal totalPayable = monthlyEMI \* loan.TenureMonths;



&nbsp;       DateTime applicationDate = DateTime.Now.AddDays(-10);

&nbsp;       DateTime disbursementDate = DateTime.Now;

&nbsp;       DateTime maturityDate = disbursementDate.AddMonths(loan.TenureMonths);



&nbsp;       return new Models.LoanData

&nbsp;       {

&nbsp;           ApplicantName = loan.ApplicantName,

&nbsp;           FatherName = loan.FatherName,

&nbsp;           Address = loan.Address,

&nbsp;           ContactNumber = loan.ContactNumber,

&nbsp;           Email = loan.Email,

&nbsp;           NationalID = loan.NationalID,

&nbsp;           AccountNumber = loan.AccountNumber,

&nbsp;           LoanAmount = $"{loan.LoanAmount:N2}",

&nbsp;           LoanAmountInWords = ConvertToWords(loan.LoanAmount),

&nbsp;           LoanPurpose = loan.LoanPurpose,

&nbsp;           TenureMonths = loan.TenureMonths.ToString(),

&nbsp;           InterestRate = $"{loan.InterestRate:N2}",

&nbsp;           MonthlyEMI = $"{monthlyEMI:N2}",

&nbsp;           TotalPayable = $"{totalPayable:N2}",

&nbsp;           ProcessingFee = $"{loan.ProcessingFee:N2}",

&nbsp;           ApplicationDate = applicationDate.ToString("dd MMMM yyyy"),

&nbsp;           DisbursementDate = disbursementDate.ToString("dd MMMM yyyy"),

&nbsp;           MaturityDate = maturityDate.ToString("dd MMMM yyyy"),

&nbsp;           BankName = "ABC Bank Limited",

&nbsp;           BankAddress = "Head Office: 123 Motijheel C/A, Dhaka-1000, Bangladesh",

&nbsp;           BankPhone = "+880-2-9559191",

&nbsp;           BankEmail = "info@abcbank.com.bd",

&nbsp;           BankWebsite = "www.abcbank.com.bd",

&nbsp;           BranchName = loan.BranchName,

&nbsp;           BranchCode = loan.BranchCode,

&nbsp;           BranchAddress = GetBranchAddress(loan.BranchCode),

&nbsp;           BranchManager = GetBranchManager(loan.BranchCode),

&nbsp;           ReferenceNumber = $"LOAN/{DateTime.Now.Year}/{loanId:D6}",

&nbsp;           GeneratedDate = DateTime.Now.ToString("dd MMMM yyyy"),

&nbsp;           GeneratedTime = DateTime.Now.ToString("hh:mm tt"),

&nbsp;           GeneratedBy = "System Administrator",

&nbsp;           WitnessName1 = "Ahmed Khan",

&nbsp;           WitnessAddress1 = "House# 23, Road# 5, Banani, Dhaka",

&nbsp;           WitnessName2 = "Nasrin Akter",

&nbsp;           WitnessAddress2 = "Flat# 2A, Holding# 67, Mohammadpur, Dhaka"

&nbsp;       };

&nbsp;   }



&nbsp;   public static List<LoanSummary> GetAllLoans()

&nbsp;   {

&nbsp;       return \_mockLoans.Values.Select(l => new LoanSummary

&nbsp;       {

&nbsp;           LoanId = l.LoanId,

&nbsp;           ApplicantName = l.ApplicantName,

&nbsp;           AccountNumber = l.AccountNumber,

&nbsp;           LoanAmount = l.LoanAmount,

&nbsp;           Status = "Active"

&nbsp;       }).ToList();

&nbsp;   }



&nbsp;   private static decimal CalculateEMI(decimal principal, decimal annualRate, int months)

&nbsp;   {

&nbsp;       if (annualRate == 0) return principal / months;

&nbsp;       decimal monthlyRate = annualRate / (12 \* 100);

&nbsp;       decimal emi = principal \* monthlyRate \* (decimal)Math.Pow((double)(1 + monthlyRate), months) / ((decimal)Math.Pow((double)(1 + monthlyRate), months) - 1);

&nbsp;       return Math.Round(emi, 2);

&nbsp;   }



&nbsp;   private static string ConvertToWords(decimal amount)

&nbsp;   {

&nbsp;       string\[] ones = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine" };

&nbsp;       string\[] teens = { "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };

&nbsp;       string\[] tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };



&nbsp;       int intAmount = (int)amount;

&nbsp;       if (intAmount == 0) return "Zero Taka Only";



&nbsp;       string words = "";

&nbsp;       if (intAmount >= 100000)

&nbsp;       {

&nbsp;           words += ConvertToWordsHelper(intAmount / 100000, ones, teens, tens) + " Lakh ";

&nbsp;           intAmount %= 100000;

&nbsp;       }

&nbsp;       if (intAmount >= 1000)

&nbsp;       {

&nbsp;           words += ConvertToWordsHelper(intAmount / 1000, ones, teens, tens) + " Thousand ";

&nbsp;           intAmount %= 1000;

&nbsp;       }

&nbsp;       if (intAmount >= 100)

&nbsp;       {

&nbsp;           words += ones\[intAmount / 100] + " Hundred ";

&nbsp;           intAmount %= 100;

&nbsp;       }

&nbsp;       if (intAmount > 0)

&nbsp;       {

&nbsp;           words += ConvertToWordsHelper(intAmount, ones, teens, tens);

&nbsp;       }

&nbsp;       return words.Trim() + " Taka Only";

&nbsp;   }



&nbsp;   private static string ConvertToWordsHelper(int number, string\[] ones, string\[] teens, string\[] tens)

&nbsp;   {

&nbsp;       if (number == 0) return "";

&nbsp;       if (number < 10) return ones\[number];

&nbsp;       if (number < 20) return teens\[number - 10];

&nbsp;       return tens\[number / 10] + (number % 10 > 0 ? " " + ones\[number % 10] : "");

&nbsp;   }



&nbsp;   private static string GetBranchAddress(string branchCode) => branchCode switch

&nbsp;   {

&nbsp;       "DHK-001" => "House# 15, Road# 27, Dhanmondi R/A, Dhaka-1209",

&nbsp;       "DHK-002" => "Plot# 23, Road# 11, Gulshan-2, Dhaka-1212",

&nbsp;       "DHK-003" => "Main Road, Savar Bazar, Savar, Dhaka-1340",

&nbsp;       \_ => "Branch Address"

&nbsp;   };



&nbsp;   private static string GetBranchManager(string branchCode) => branchCode switch

&nbsp;   {

&nbsp;       "DHK-001" => "Mr. Khalid Rahman",

&nbsp;       "DHK-002" => "Ms. Aisha Siddique",

&nbsp;       "DHK-003" => "Mr. Rafiq Ahmed",

&nbsp;       \_ => "Branch Manager"

&nbsp;   };



&nbsp;   public class LoanDataModel

&nbsp;   {

&nbsp;       public int LoanId { get; set; }

&nbsp;       public string ApplicantName { get; set; }

&nbsp;       public string FatherName { get; set; }

&nbsp;       public string Address { get; set; }

&nbsp;       public string ContactNumber { get; set; }

&nbsp;       public string Email { get; set; }

&nbsp;       public string NationalID { get; set; }

&nbsp;       public string AccountNumber { get; set; }

&nbsp;       public decimal LoanAmount { get; set; }

&nbsp;       public string LoanPurpose { get; set; }

&nbsp;       public int TenureMonths { get; set; }

&nbsp;       public decimal InterestRate { get; set; }

&nbsp;       public decimal ProcessingFee { get; set; }

&nbsp;       public string BranchName { get; set; }

&nbsp;       public string BranchCode { get; set; }

&nbsp;   }



&nbsp;   public class LoanSummary

&nbsp;   {

&nbsp;       public int LoanId { get; set; }

&nbsp;       public string ApplicantName { get; set; }

&nbsp;       public string AccountNumber { get; set; }

&nbsp;       public decimal LoanAmount { get; set; }

&nbsp;       public string Status { get; set; }

&nbsp;   }

}

}



</details>



---



\## \*\*FILE 5: Services/IDocumentService.cs\*\*

```csharp

using BankDocumentDemo.Models;



namespace BankDocumentDemo.Services

{

&nbsp;   public interface IDocumentService

&nbsp;   {

&nbsp;       Task<byte\[]> GenerateDocumentAsync(DocumentRequest request);

&nbsp;   }

}



FILE 6: Services/DocumentService.cs

csharpusing Syncfusion.DocIO;

using Syncfusion.DocIO.DLS;

using Syncfusion.DocIORenderer;

using Syncfusion.Pdf;

using BankDocumentDemo.Models;

using BankDocumentDemo.MockData;



namespace BankDocumentDemo.Services

{

&nbsp;   public class DocumentService : IDocumentService

&nbsp;   {

&nbsp;       private readonly ILogger<DocumentService> \_logger;

&nbsp;       private readonly string \_templatePath;



&nbsp;       public DocumentService(ILogger<DocumentService> logger)

&nbsp;       {

&nbsp;           \_logger = logger;

&nbsp;           \_templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");

&nbsp;       }



&nbsp;       public async Task<byte\[]> GenerateDocumentAsync(DocumentRequest request)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               \_logger.LogInformation($"Generating {request.DocumentType} for Loan ID: {request.LoanId}");

&nbsp;               var loanData = LoanMockData.GetLoanData(request.LoanId);

&nbsp;               string templateFileName = GetTemplateFileName(request.DocumentType);

&nbsp;               string templateFullPath = Path.Combine(\_templatePath, templateFileName);



&nbsp;               if (!File.Exists(templateFullPath))

&nbsp;                   throw new FileNotFoundException($"Template not found: {templateFileName}");



&nbsp;               byte\[] wordDocument = await GenerateWordDocumentAsync(templateFullPath, loanData);



&nbsp;               if (request.OutputFormat == OutputFormat.PDF)

&nbsp;                   return await ConvertWordToPdfAsync(wordDocument);



&nbsp;               return wordDocument;

&nbsp;           }

&nbsp;           catch (Exception ex)

&nbsp;           {

&nbsp;               \_logger.LogError(ex, $"Error generating document");

&nbsp;               throw;

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<byte\[]> GenerateWordDocumentAsync(string templatePath, LoanData data)

&nbsp;       {

&nbsp;           using (FileStream templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))

&nbsp;           using (WordDocument document = new WordDocument(templateStream, FormatType.Docx))

&nbsp;           {

&nbsp;               ReplaceBookmarks(document, data);

&nbsp;               using (MemoryStream outputStream = new MemoryStream())

&nbsp;               {

&nbsp;                   document.Save(outputStream, FormatType.Docx);

&nbsp;                   return await Task.FromResult(outputStream.ToArray());

&nbsp;               }

&nbsp;           }

&nbsp;       }



&nbsp;       private async Task<byte\[]> ConvertWordToPdfAsync(byte\[] wordDocument)

&nbsp;       {

&nbsp;           using (MemoryStream wordStream = new MemoryStream(wordDocument))

&nbsp;           using (WordDocument document = new WordDocument(wordStream, FormatType.Docx))

&nbsp;           using (DocIORenderer renderer = new DocIORenderer())

&nbsp;           {

&nbsp;               PdfDocument pdfDocument = renderer.ConvertToPDF(document);

&nbsp;               using (MemoryStream pdfStream = new MemoryStream())

&nbsp;               {

&nbsp;                   pdfDocument.Save(pdfStream);

&nbsp;                   pdfDocument.Close();

&nbsp;                   return await Task.FromResult(pdfStream.ToArray());

&nbsp;               }

&nbsp;           }

&nbsp;       }



&nbsp;       private void ReplaceBookmarks(WordDocument document, LoanData data)

&nbsp;       {

&nbsp;           var properties = data.GetType().GetProperties();

&nbsp;           foreach (var property in properties)

&nbsp;           {

&nbsp;               string placeholder = $"{{{property.Name}}}";

&nbsp;               string value = property.GetValue(data)?.ToString() ?? "";

&nbsp;               document.Replace(placeholder, value, true, true);

&nbsp;           }

&nbsp;       }



&nbsp;       private string GetTemplateFileName(DocumentType documentType) => documentType switch

&nbsp;       {

&nbsp;           DocumentType.LoanProposal => "LoanProposal.docx",

&nbsp;           DocumentType.PromissoryNote => "PromissoryNote.docx",

&nbsp;           \_ => throw new ArgumentException($"Unknown document type: {documentType}")

&nbsp;       };

&nbsp;   }

}



FILE 7: Controllers/DocumentController.cs

csharpusing Microsoft.AspNetCore.Mvc;

using BankDocumentDemo.Models;

using BankDocumentDemo.Services;

using BankDocumentDemo.MockData;



namespace BankDocumentDemo.Controllers

{

&nbsp;   \[ApiController]

&nbsp;   \[Route("api/\[controller]")]

&nbsp;   public class DocumentController : ControllerBase

&nbsp;   {

&nbsp;       private readonly IDocumentService \_documentService;

&nbsp;       private readonly ILogger<DocumentController> \_logger;



&nbsp;       public DocumentController(IDocumentService documentService, ILogger<DocumentController> logger)

&nbsp;       {

&nbsp;           \_documentService = documentService;

&nbsp;           \_logger = logger;

&nbsp;       }



&nbsp;       \[HttpGet("loans")]

&nbsp;       public IActionResult GetAllLoans() => Ok(LoanMockData.GetAllLoans());



&nbsp;       \[HttpGet("loans/{loanId}")]

&nbsp;       public IActionResult GetLoanById(int loanId)

&nbsp;       {

&nbsp;           try { return Ok(LoanMockData.GetLoanData(loanId)); }

&nbsp;           catch (KeyNotFoundException ex) { return NotFound(new { error = ex.Message }); }

&nbsp;       }



&nbsp;       \[HttpPost("generate")]

&nbsp;       public async Task<IActionResult> GenerateDocument(\[FromBody] DocumentRequest request)

&nbsp;       {

&nbsp;           try

&nbsp;           {

&nbsp;               var documentBytes = await \_documentService.GenerateDocumentAsync(request);

&nbsp;               string fileName = $"{request.DocumentType}\_{request.LoanId}\_{DateTime.Now:yyyyMMddHHmmss}";

&nbsp;               string contentType = request.OutputFormat == OutputFormat.PDF ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

&nbsp;               string extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";

&nbsp;               return File(documentBytes, contentType, $"{fileName}.{extension}");

&nbsp;           }

&nbsp;           catch (Exception ex)

&nbsp;           {

&nbsp;               \_logger.LogError(ex, "Error generating document");

&nbsp;               return StatusCode(500, new { error = ex.Message });

&nbsp;           }

&nbsp;       }



&nbsp;       \[HttpGet("download/{loanId}/{documentType}")]

&nbsp;       public async Task<IActionResult> DownloadDocument(int loanId, DocumentType documentType, \[FromQuery] OutputFormat format = OutputFormat.PDF)

&nbsp;       {

&nbsp;           return await GenerateDocument(new DocumentRequest { LoanId = loanId, DocumentType = documentType, OutputFormat = format });

&nbsp;       }



&nbsp;       \[HttpGet("types")]

&nbsp;       public IActionResult GetDocumentTypes()

&nbsp;       {

&nbsp;           return Ok(Enum.GetValues(typeof(DocumentType)).Cast<DocumentType>().Select(t => new { Id = (int)t, Name = t.ToString() }));

&nbsp;       }

&nbsp;   }

}



FILE 8: Program.cs (Replace existing)

csharpusing BankDocumentDemo.Services;



var builder = WebApplication.CreateBuilder(args);



builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();

builder.Services.AddSwaggerGen();

builder.Services.AddScoped<IDocumentService, DocumentService>();

builder.Services.AddCors(options => options.AddPolicy("AllowAll", b => b.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader()));



var app = builder.Build();



app.UseSwagger();

app.UseSwaggerUI(c => c.RoutePrefix = string.Empty);

app.UseHttpsRedirection();

app.UseCors("AllowAll");

app.UseAuthorization();

app.MapControllers();



Console.WriteLine("=== Bank Document API Running ===");

Console.WriteLine("Swagger: https://localhost:7xxx");

Console.WriteLine("=================================");



app.Run();



FILE 9: BankDocumentDemo.csproj (Replace existing)

xml<Project Sdk="Microsoft.NET.Sdk.Web">

&nbsp; <PropertyGroup>

&nbsp;   <TargetFramework>net8.0</TargetFramework>

&nbsp;   <Nullable>enable</Nullable>

&nbsp;   <ImplicitUsings>enable</ImplicitUsings>

&nbsp; </PropertyGroup>

&nbsp; <ItemGroup>

&nbsp;   <PackageReference Include="Syncfusion.DocIO.Net.Core" Version="27.1.48" />

&nbsp;   <PackageReference Include="Syncfusion.DocIORenderer.Net.Core" Version="27.1.48" />

&nbsp;   <PackageReference Include="Syncfusion.Pdf.Net.Core" Version="27.1.48" />

&nbsp;   <PackageReference Include="Swashbuckle.AspNetCore" Version="6.4.0" />

&nbsp; </ItemGroup>

&nbsp; <ItemGroup>

&nbsp;   <None Update="Templates\\\*\*\\\*">

&nbsp;     <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>

&nbsp;   </None>

&nbsp; </ItemGroup>

</Project>



STEP 5: Create Word Templates

You need to create TWO Word documents manually in the Templates folder.

I've provided you with the template text in my previous message. Here's a quick summary:

Templates/LoanProposal.docx



Open Microsoft Word

Copy the Loan Proposal template text from my previous response

Replace dynamic fields with placeholders like {ApplicantName}, {LoanAmount}, etc.

Save as LoanProposal.docx in the Templates folder



Templates/PromissoryNote.docx



Open Microsoft Word

Copy the Promissory Note template text

Add placeholders

Save as PromissoryNote.docx





STEP 6: Run the Project

bashdotnet restore

dotnet run



Alternative: GitHub Repository

I recommend you:



Create a GitHub repository and I'll guide you through pushing this code

Use a code-sharing service like:



GitHub Gist (for individual files)

Pastebin

CodePen









