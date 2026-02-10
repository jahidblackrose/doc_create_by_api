using BankingDocumentAPI.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using DocumentFormat.OpenXml;
using PdfDocument = QuestPDF.Fluent.Document;
using DocumentType = BankingDocumentAPI.Models.DocumentType;

namespace BankingDocumentAPI.Services
{
    public class WordDocumentService : IDocumentService
    {
        private readonly ILogger<WordDocumentService> _logger;
        private readonly WordToPdfConverterService _pdfConverterService;

        public WordDocumentService(ILogger<WordDocumentService> logger, WordToPdfConverterService pdfConverterService)
        {
            _logger = logger;
            _pdfConverterService = pdfConverterService;
            QuestPDF.Settings.License = LicenseType.Community;
        }

        public async Task<byte[]> GenerateDocumentAsync(DocumentRequest request)
        {
            var data = await GetDocumentDataAsync(request.LoanId, request.DocumentType);

            if (request.OutputFormat == OutputFormat.PDF)
            {
                // Generate Word from template first, then convert to PDF
                var wordBytes = GenerateSimpleWordFile(request.DocumentType, data);
                return await ConvertWordToPdfAsync(wordBytes);
            }
            else
            {
                // For DOCX, use template-based generation
                return GenerateSimpleWordFile(request.DocumentType, data);
            }
        }

        public async Task<byte[]> GenerateWordDocumentAsync(string templateName, object data)
        {
            await Task.CompletedTask;
            return GenerateSimpleWordFile(DocumentType.LoanProposal, data);
        }

        public async Task<byte[]> GeneratePdfFromWordAsync(byte[] wordDocument)
        {
            return await ConvertWordToPdfAsync(wordDocument);
        }

        private async Task<byte[]> ConvertWordToPdfAsync(byte[] wordBytes)
        {
            // Use the new WordToPdfConverterService for template-based PDF generation
            // This parses the Word document structure and recreates it in QuestPDF
            return await Task.Run(() => _pdfConverterService.ConvertWordBytesToPdf(wordBytes));
        }

        private byte[] GeneratePdf(DocumentType documentType, object data)
        {
            return documentType switch
            {
                DocumentType.LoanProposal => GenerateLoanProposalPdf((LoanProposalData)data),
                DocumentType.PromissoryNote => GeneratePromissoryNotePdf((PromissoryNoteData)data),
                DocumentType.LetterOfContinuity => GenerateGenericPdf("LETTER OF CONTINUITY", data),
                DocumentType.StandingOrder => GenerateGenericPdf("STANDING ORDER REQUEST", data),
                DocumentType.LetterOfIndemnity => GenerateGenericPdf("LETTER OF INDEMNITY", data),
                DocumentType.PersonalGuarantee => GenerateGenericPdf("PERSONAL GUARANTEE", data),
                DocumentType.UDC => GenerateGenericPdf("UNDERTAKING OF DEBT CREATION", data),
                _ => GenerateGenericPdf(documentType.ToString(), data)
            };
        }

        private byte[] GenerateSimpleWordFile(DocumentType docType, object data)
        {
            // Use template-based generation
            var templateName = GetTemplateName(docType);
            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates", templateName);

            // Check if template exists
            if (!File.Exists(templatePath))
            {
                _logger.LogWarning($"Template not found: {templatePath}. Using programmatic generation.");
                return GenerateWordProgrammatically(docType, data);
            }

            try
            {
                return GenerateFromTemplate(templatePath, data);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error generating from template: {templatePath}. Using programmatic generation.");
                return GenerateWordProgrammatically(docType, data);
            }
        }

        private string GetTemplateName(DocumentType docType)
        {
            return docType switch
            {
                DocumentType.LoanProposal => "LoanProposal.docx",
                DocumentType.PromissoryNote => "PromissoryNote.docx",
                DocumentType.LetterOfContinuity => "LetterOfContinuity.docx",
                DocumentType.LetterOfRevival => "LetterOfRevival.docx",
                DocumentType.LetterOfArrangement => "LetterOfArrangement.docx",
                DocumentType.StandingOrder => "StandingOrder.docx",
                DocumentType.LetterOfIndemnity => "LetterOfIndemnity.docx",
                DocumentType.LetterOfLien => "LetterOfLien.docx",
                DocumentType.PersonalGuarantee => "PersonalGuarantee.docx",
                DocumentType.UDC => "UDC.docx",
                _ => "LoanProposal.docx"
            };
        }

        private byte[] GenerateFromTemplate(string templatePath, object data)
        {
            // Copy template to memory stream
            var memoryStream = new MemoryStream();
            using (var templateStream = File.OpenRead(templatePath))
            {
                templateStream.CopyTo(memoryStream);
            }

            // Open the copied document
            memoryStream.Position = 0;
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, true))
            {
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart == null)
                {
                    throw new InvalidOperationException("Template file is not a valid Word document.");
                }

                // Get all text elements in the document
                var body = mainPart.Document.Body;
                if (body == null) return memoryStream.ToArray();

                // Replace placeholders in all text elements
                ReplacePlaceholdersInParagraphs(body, data);
                ReplacePlaceholdersInTables(body, data);
                ReplacePlaceholdersInHeaders(mainPart, data);
                ReplacePlaceholdersInFooters(mainPart, data);

                mainPart.Document.Save();
            }

            return memoryStream.ToArray();
        }

        private void ReplacePlaceholdersInParagraphs(Body body, object data)
        {
            var properties = data.GetType().GetProperties()
                .ToDictionary(p => p.Name, p => p.GetValue(data)?.ToString() ?? "");

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        if (text.Text != null)
                        {
                            text.Text = ReplacePlaceholders(text.Text, properties);
                        }
                    }
                }
            }
        }

        private void ReplacePlaceholdersInTables(Body body, object data)
        {
            var properties = data.GetType().GetProperties()
                .ToDictionary(p => p.Name, p => p.GetValue(data)?.ToString() ?? "");

            foreach (var table in body.Elements<Table>())
            {
                foreach (var row in table.Elements<TableRow>())
                {
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        foreach (var paragraph in cell.Elements<Paragraph>())
                        {
                            foreach (var run in paragraph.Elements<Run>())
                            {
                                foreach (var text in run.Elements<Text>())
                                {
                                    if (text.Text != null)
                                    {
                                        text.Text = ReplacePlaceholders(text.Text, properties);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ReplacePlaceholdersInHeaders(MainDocumentPart mainPart, object data)
        {
            var properties = data.GetType().GetProperties()
                .ToDictionary(p => p.Name, p => p.GetValue(data)?.ToString() ?? "");

            foreach (var headerPart in mainPart.HeaderParts)
            {
                foreach (var paragraph in headerPart.Header.Elements<Paragraph>())
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            if (text.Text != null)
                            {
                                text.Text = ReplacePlaceholders(text.Text, properties);
                            }
                        }
                    }
                }
            }
        }

        private void ReplacePlaceholdersInFooters(MainDocumentPart mainPart, object data)
        {
            var properties = data.GetType().GetProperties()
                .ToDictionary(p => p.Name, p => p.GetValue(data)?.ToString() ?? "");

            foreach (var footerPart in mainPart.FooterParts)
            {
                foreach (var paragraph in footerPart.Footer.Elements<Paragraph>())
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            if (text.Text != null)
                            {
                                text.Text = ReplacePlaceholders(text.Text, properties);
                            }
                        }
                    }
                }
            }
        }

        private string ReplacePlaceholders(string text, Dictionary<string, string> properties)
        {
            // Match {PropertyName} pattern
            var result = text;
            foreach (var prop in properties)
            {
                result = result.Replace($"{{{prop.Key}}}", prop.Value);
            }
            return result;
        }

        private byte[] GenerateWordProgrammatically(DocumentType docType, object data)
        {
            // Fallback: Generate DOCX programmatically (original method)
            var memoryStream = new MemoryStream();

            using (var wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                var body = mainPart.Document.AppendChild(new Body());

                // Add title
                var titlePara = body.AppendChild(new Paragraph());
                var titleRun = titlePara.AppendChild(new Run());
                titleRun.AppendChild(new Text(docType.ToString().ToUpper())
                {
                    Space = SpaceProcessingModeValues.Preserve
                });

                titlePara.ParagraphProperties = new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center },
                    new SpacingBetweenLines() { Before = "200", After = "200" }
                );

                titleRun.RunProperties = new RunProperties(
                    new Bold(),
                    new FontSize() { Val = "32" }
                );

                // Add separator
                var sepPara = body.AppendChild(new Paragraph());
                var sepRun = sepPara.AppendChild(new Run());
                sepRun.AppendChild(new Text(new string('=', 50))
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
                sepPara.ParagraphProperties = new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                );

                // Get properties
                var properties = data.GetType().GetProperties();
                var bankNameProp = data.GetType().GetProperty("BankName");
                var referenceProp = data.GetType().GetProperty("ReferenceNumber");
                var generatedDateProp = data.GetType().GetProperty("GeneratedDate");

                // Add bank header
                if (bankNameProp != null)
                {
                    var bankName = bankNameProp.GetValue(data)?.ToString() ?? "ABC Bank Ltd.";
                    var bankPara = body.AppendChild(new Paragraph());
                    var bankRun = bankPara.AppendChild(new Run());
                    bankRun.AppendChild(new Text(bankName)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
                    bankPara.ParagraphProperties = new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                    );
                    bankRun.RunProperties = new RunProperties(new Bold());
                }

                // Add reference and date
                if (referenceProp != null)
                {
                    var refPara = body.AppendChild(new Paragraph());
                    var refRun = refPara.AppendChild(new Run());
                    refRun.AppendChild(new Text($"Reference No: {referenceProp.GetValue(data)}")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
                    refRun.RunProperties = new RunProperties(new Bold());
                }

                if (generatedDateProp != null)
                {
                    var datePara = body.AppendChild(new Paragraph());
                    var dateRun = datePara.AppendChild(new Run());
                    dateRun.AppendChild(new Text($"Date: {generatedDateProp.GetValue(data)}")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
                    dateRun.RunProperties = new RunProperties(new Bold());
                }

                body.AppendChild(new Paragraph());

                // Add all properties
                foreach (var prop in properties)
                {
                    if (prop.Name == "BankName" || prop.Name == "BranchName" ||
                        prop.Name == "ReferenceNumber" || prop.Name == "GeneratedDate")
                        continue;

                    var para = body.AppendChild(new Paragraph());
                    var run = para.AppendChild(new Run());

                    var label = prop.Name;
                    var value = prop.GetValue(data)?.ToString() ?? "";

                    run.AppendChild(new Text($"{label}: ")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
                    run.RunProperties = new RunProperties(new Bold());

                    var valueRun = para.AppendChild(new Run());
                    valueRun.AppendChild(new Text(value)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
                }

                body.AppendChild(new Paragraph());
                body.AppendChild(new Paragraph());

                // Add signature line
                var sigPara = body.AppendChild(new Paragraph());
                var sigRun = sigPara.AppendChild(new Run());
                sigRun.AppendChild(new Text("Authorized Signature: _______________")
                {
                    Space = SpaceProcessingModeValues.Preserve
                });

                mainPart.Document.Save();
            }

            return memoryStream.ToArray();
        }

        private byte[] GenerateLoanProposalPdf(LoanProposalData data)
        {
            return PdfDocument.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(2, Unit.Centimetre);

                    page.Header().Element(header =>
                    {
                        header.Column(column =>
                        {
                            column.Item().AlignCenter().Text(data.BankName).Bold().FontSize(16);
                            column.Item().AlignCenter().Text(data.BankAddress).FontSize(10);
                            column.Item().PaddingTop(1).PaddingBottom(1).LineHorizontal(1);
                        });
                    });

                    page.Content().Element(content =>
                    {
                        content.Column(column =>
                        {
                            column.Item().AlignCenter().Text("LOAN PROPOSAL").Bold().FontSize(18);
                            column.Item().PaddingBottom(1);

                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Reference No:").Bold();
                                grid.Item().Text(data.ReferenceNumber);
                                grid.Item().Text("Date:").Bold();
                                grid.Item().Text(data.GeneratedDate);
                            });

                            column.Item().PaddingTop(1).PaddingBottom(1);
                            column.Item().Text("APPLICANT DETAILS").Bold().FontSize(14);
                            column.Item().PaddingBottom(0.5f);

                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Applicant Name:").Bold();
                                grid.Item().Text(data.ApplicantName);
                                grid.Item().Text("Account Number:").Bold();
                                grid.Item().Text(data.AccountNumber);
                            });

                            column.Item().PaddingTop(1).PaddingBottom(1);
                            column.Item().Text("LOAN DETAILS").Bold().FontSize(14);
                            column.Item().PaddingBottom(0.5f);

                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Loan Amount:").Bold();
                                grid.Item().Text($"Rs. {data.LoanAmount:N0}");
                                grid.Item().Text("Purpose:").Bold();
                                grid.Item().Text(data.LoanPurpose);
                                grid.Item().Text("Tenure:").Bold();
                                grid.Item().Text($"{data.TenureMonths} months");
                                grid.Item().Text("Interest Rate:").Bold();
                                grid.Item().Text($"{data.InterestRate}% p.a.");
                                grid.Item().Text("Monthly EMI:").Bold();
                                grid.Item().Text($"Rs. {data.MonthlyEMI:N2}");
                            });

                            column.Item().PaddingTop(1);
                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Branch:").Bold();
                                grid.Item().Text(data.BranchName);
                                grid.Item().Text("Application Date:").Bold();
                                grid.Item().Text(data.ApplicationDate.ToString("dd-MMM-yyyy"));
                            });

                            column.Item().PaddingTop(3);
                            column.Item().Text("Authorized Signature: _______________");
                        });
                    });

                    page.Footer().AlignCenter().Text(x =>
                    {
                        x.Span("Page ");
                        x.CurrentPageNumber();
                    });
                });
            }).GeneratePdf();
        }

        private byte[] GeneratePromissoryNotePdf(PromissoryNoteData data)
        {
            return PdfDocument.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(2, Unit.Centimetre);

                    page.Header().Element(header =>
                    {
                        header.Column(column =>
                        {
                            column.Item().AlignCenter().Text("PROMISSORY NOTE").Bold().FontSize(18);
                        });
                    });

                    page.Content().Element(content =>
                    {
                        content.Column(column =>
                        {
                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Reference No:").Bold();
                                grid.Item().Text(data.ReferenceNumber);
                                grid.Item().Text("Date:").Bold();
                                grid.Item().Text(data.GeneratedDate);
                            });

                            column.Item().PaddingTop(1).PaddingBottom(1);
                            column.Item().Text("BORROWER DETAILS").Bold().FontSize(14);

                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Borrower Name:").Bold();
                                grid.Item().Text(data.BorrowerName);
                                grid.Item().Text("CNIC:").Bold();
                                grid.Item().Text(data.BorrowerCNIC);
                                grid.Item().Text("Address:").Bold();
                                grid.Item().Text(data.BorrowerAddress);
                                grid.Item().Text("Account Number:").Bold();
                                grid.Item().Text(data.AccountNumber);
                            });

                            column.Item().PaddingTop(1).PaddingBottom(1);
                            column.Item().Text("LOAN DETAILS").Bold().FontSize(14);

                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Principal Amount:").Bold();
                                grid.Item().Text($"Rs. {data.PrincipalAmount:N0} ({data.AmountInWords})");
                                grid.Item().Text("Interest Rate:").Bold();
                                grid.Item().Text($"{data.InterestRate}% p.a.");
                                grid.Item().Text("Loan Start Date:").Bold();
                                grid.Item().Text(data.LoanStartDate.ToString("dd-MMM-yyyy"));
                                grid.Item().Text("Loan Tenure:").Bold();
                                grid.Item().Text($"{data.LoanTenureMonths} months");
                                grid.Item().Text("Maturity Date:").Bold();
                                grid.Item().Text(data.MaturityDate.ToString("dd-MMM-yyyy"));
                                grid.Item().Text("Monthly Installment:").Bold();
                                grid.Item().Text($"Rs. {data.MonthlyInstallment:N2}");
                                grid.Item().Text("Payment Mode:").Bold();
                                grid.Item().Text(data.PaymentMode);
                            });

                            column.Item().PaddingTop(1);
                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Branch:").Bold();
                                grid.Item().Text(data.BranchName);
                                grid.Item().Text("Bank:").Bold();
                                grid.Item().Text(data.BankName);
                            });

                            column.Item().PaddingTop(2).PaddingBottom(1);
                            column.Item().Text("WITNESS DETAILS").Bold().FontSize(14);

                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Witness Name:").Bold();
                                grid.Item().Text(data.WitnessName);
                                grid.Item().Text("Witness CNIC:").Bold();
                                grid.Item().Text(data.WitnessCNIC);
                            });

                            column.Item().PaddingTop(2);
                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Borrower Signature: _______________");
                                grid.Item().Text("Witness Signature: _______________");
                            });
                        });
                    });
                });
            }).GeneratePdf();
        }

        private byte[] GenerateGenericPdf(string title, object data)
        {
            var properties = data.GetType().GetProperties();

            string bankName = "ABC Bank Ltd.";
            string branchName = "Main Branch";
            string referenceNumber = "";
            string generatedDate = DateTime.Now.ToString("dd-MMM-yyyy");

            var bankNameProp = data.GetType().GetProperty("BankName");
            var branchNameProp = data.GetType().GetProperty("BranchName");
            var referenceProp = data.GetType().GetProperty("ReferenceNumber");
            var generatedDateProp = data.GetType().GetProperty("GeneratedDate");

            if (bankNameProp != null) bankName = bankNameProp.GetValue(data)?.ToString() ?? bankName;
            if (branchNameProp != null) branchName = branchNameProp.GetValue(data)?.ToString() ?? branchName;
            if (referenceProp != null) referenceNumber = referenceProp.GetValue(data)?.ToString() ?? "";
            if (generatedDateProp != null) generatedDate = generatedDateProp.GetValue(data)?.ToString() ?? generatedDate;

            return PdfDocument.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(2, Unit.Centimetre);

                    page.Header().Element(header =>
                    {
                        header.Column(column =>
                        {
                            column.Item().AlignCenter().Text(bankName).Bold().FontSize(16);
                            column.Item().AlignCenter().Text(branchName).FontSize(10);
                            column.Item().PaddingTop(1).PaddingBottom(1).LineHorizontal(1);
                        });
                    });

                    page.Content().Element(content =>
                    {
                        content.Column(column =>
                        {
                            column.Item().AlignCenter().Text(title).Bold().FontSize(16);
                            column.Item().PaddingBottom(1);

                            column.Item().Grid(grid =>
                            {
                                grid.Columns(2);
                                grid.Item().Text("Reference No:").Bold();
                                grid.Item().Text(referenceNumber);
                                grid.Item().Text("Date:").Bold();
                                grid.Item().Text(generatedDate);
                            });

                            column.Item().PaddingTop(1);

                            foreach (var prop in properties)
                            {
                                if (prop.Name != "BankName" && prop.Name != "BranchName" &&
                                    prop.Name != "ReferenceNumber" && prop.Name != "GeneratedDate")
                                {
                                    column.Item().Grid(grid =>
                                    {
                                        grid.Columns(2);
                                        grid.Item().Text($"{prop.Name}:").Bold();
                                        grid.Item().Text(prop.GetValue(data)?.ToString() ?? "");
                                    });
                                }
                            }

                            column.Item().PaddingTop(3);
                            column.Item().Text("Authorized Signature: _______________");
                        });
                    });

                    page.Footer().AlignCenter().Text(x =>
                    {
                        x.Span("Page ");
                        x.CurrentPageNumber();
                    });
                });
            }).GeneratePdf();
        }

        private async Task<object> GetDocumentDataAsync(long loanId, DocumentType documentType)
        {
            await Task.Delay(1);

            return documentType switch
            {
                DocumentType.LoanProposal => GetLoanProposalMockData(loanId),
                DocumentType.PromissoryNote => GetPromissoryNoteMockData(loanId),
                DocumentType.LetterOfContinuity => GetLetterOfContinuityMockData(loanId),
                DocumentType.StandingOrder => GetStandingOrderMockData(loanId),
                DocumentType.LetterOfIndemnity => GetLetterOfIndemnityMockData(loanId),
                DocumentType.PersonalGuarantee => GetPersonalGuaranteeMockData(loanId),
                DocumentType.UDC => GetUDCMockData(loanId),
                _ => GetLoanProposalMockData(loanId)
            };
        }

        // Mock Data Methods
        private LoanProposalData GetLoanProposalMockData(long loanId)
        {
            return new LoanProposalData
            {
                ApplicantName = "John Doe",
                AccountNumber = "1234567890",
                LoanAmount = 500000,
                LoanPurpose = "Home Renovation",
                TenureMonths = 60,
                InterestRate = 8.5m,
                ApplicationDate = DateTime.Now,
                BranchName = "Main Branch",
                BankAddress = "123 Banking Street, Main Branch, City",
                MonthlyEMI = 10247.50m,
                GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),
                ReferenceNumber = $"LOAN/{DateTime.Now.Year}/{loanId}"
            };
        }

        private PromissoryNoteData GetPromissoryNoteMockData(long loanId)
        {
            return new PromissoryNoteData
            {
                BorrowerName = "John Doe",
                BorrowerAddress = "456 Residential Area, City",
                BorrowerCNIC = "12345-6789012-3",
                AccountNumber = "1234567890",
                PrincipalAmount = 500000,
                AmountInWords = "Five Hundred Thousand Only",
                InterestRate = 8.5m,
                LoanStartDate = DateTime.Now,
                LoanTenureMonths = 60,
                MaturityDate = DateTime.Now.AddMonths(60),
                MonthlyInstallment = 10247.50m,
                PaymentMode = "Monthly Post-Dated Cheques",
                BranchName = "Main Branch",
                BankAddress = "123 Banking Street, Main Branch, City",
                BankName = "ABC Bank Ltd.",
                GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),
                ReferenceNumber = $"PN/{DateTime.Now.Year}/{loanId}",
                WitnessName = "Jane Smith",
                WitnessCNIC = "54321-8765432-1"
            };
        }

        private LetterOfContinuityData GetLetterOfContinuityMockData(long loanId)
        {
            return new LetterOfContinuityData
            {
                ApplicantName = "John Doe",
                AccountNumber = "1234567890",
                LoanAccountNumber = "LN-2024-001234",
                OutstandingAmount = 350000,
                FacilityType = "Term Loan",
                SanctionReference = "SANCTION/2024/12345",
                SanctionDate = DateTime.Now.AddMonths(-6),
                BranchName = "Main Branch",
                BankName = "ABC Bank Ltd.",
                GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),
                ReferenceNumber = $"LOC/{DateTime.Now.Year}/{loanId}"
            };
        }

        private StandingOrderData GetStandingOrderMockData(long loanId)
        {
            return new StandingOrderData
            {
                AccountHolderName = "John Doe",
                AccountNumber = "1234567890",
                AccountType = "Current Account",
                Amount = 15000,
                AmountInWords = "Fifteen Thousand Only",
                Frequency = "Monthly",
                StartDate = DateTime.Now.AddDays(15),
                EndDate = DateTime.Now.AddYears(2),
                BeneficiaryName = "ABC Construction Company",
                BeneficiaryAccountNumber = "9876543210",
                BeneficiaryBank = "XYZ Bank",
                BranchName = "Main Branch",
                BankName = "ABC Bank Ltd.",
                GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),
                ReferenceNumber = $"SO/{DateTime.Now.Year}/{loanId}"
            };
        }

        private LetterOfIndemnityData GetLetterOfIndemnityMockData(long loanId)
        {
            return new LetterOfIndemnityData
            {
                ApplicantName = "John Doe",
                AccountNumber = "1234567890",
                RequestType = "Lost Cheque",
                ReferenceDetails = "Cheque #001234 dated 15-Jan-2024 for Rs. 25,000",
                IndemnityAmount = 25000,
                BranchName = "Main Branch",
                BankName = "ABC Bank Ltd.",
                GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),
                ReferenceNumber = $"LOI/{DateTime.Now.Year}/{loanId}"
            };
        }

        private PersonalGuaranteeData GetPersonalGuaranteeMockData(long loanId)
        {
            return new PersonalGuaranteeData
            {
                GuarantorName = "Jane Smith",
                GuarantorCNIC = "54321-8765432-1",
                GuarantorAddress = "789 Business District, City",
                BorrowerName = "John Doe",
                BorrowerCNIC = "12345-6789012-3",
                GuaranteedAmount = 500000,
                AmountInWords = "Five Hundred Thousand Only",
                LoanAccountNumber = "LN-2024-001234",
                BranchName = "Main Branch",
                BankName = "ABC Bank Ltd.",
                GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),
                ReferenceNumber = $"PG/{DateTime.Now.Year}/{loanId}"
            };
        }

        private UDCData GetUDCMockData(long loanId)
        {
            return new UDCData
            {
                ApplicantName = "John Doe",
                AccountNumber = "1234567890",
                FacilityType = "Cash Credit Facility",
                CreditLimit = 1000000,
                SecurityDetails = "Property Mortgage - Plot #123, Sector A",
                SanctionDate = DateTime.Now.AddMonths(-3),
                SanctionReference = "UDC/2024/67890",
                BranchName = "Main Branch",
                BankName = "ABC Bank Ltd.",
                GeneratedDate = DateTime.Now.ToString("dd-MMM-yyyy"),
                ReferenceNumber = $"UDC/{DateTime.Now.Year}/{loanId}"
            };
        }
    }
}
