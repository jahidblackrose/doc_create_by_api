using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

class Program
{
    static void Main(string[] args)
    {
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
        Console.WriteLine("Generating Word document templates...");
        Console.WriteLine($"Output folder: {templatePath}");
        Directory.CreateDirectory(templatePath);

        GenerateLoanProposalTemplate(templatePath);
        GeneratePromissoryNoteTemplate(templatePath);
        GenerateLetterOfContinuityTemplate(templatePath);
        GenerateLetterOfRevivalTemplate(templatePath);
        GenerateLetterOfArrangementTemplate(templatePath);
        GenerateStandingOrderTemplate(templatePath);
        GenerateLetterOfIndemnityTemplate(templatePath);
        GenerateLetterOfLienTemplate(templatePath);
        GeneratePersonalGuaranteeTemplate(templatePath);
        GenerateUDCTemplate(templatePath);

        Console.WriteLine($"\nAll templates generated successfully in: {templatePath}");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    static void AddCenteredHeader(IWSection section, string text, BuiltinStyle style)
    {
        IWParagraph paragraph = section.AddParagraph();
        paragraph.AppendText(text);
        paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
        paragraph.ApplyStyle(style);
    }

    static void AddField(IWParagraph paragraph, string label, string placeholder)
    {
        paragraph.AppendText(label + ": ");
        paragraph.AppendText("{" + placeholder + "}");
        paragraph.AppendText("\n");
    }

    static void SaveTemplate(WordDocument document, string path, string fileName)
    {
        string filePath = Path.Combine(path, fileName);
        using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        {
            document.Save(stream, FormatType.Docx);
        }
    }

    static void GenerateLoanProposalTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "ABC BANK LTD.", BuiltinStyle.Heading1);
            p = section.AddParagraph();
            p.AppendText("{BankAddress}\n");
            p.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

            AddCenteredHeader(section, "LOAN PROPOSAL", BuiltinStyle.Heading2);

            p = section.AddParagraph();
            p.AppendText("Reference No: {ReferenceNumber}                    Date: {GeneratedDate}\n");

            p = section.AddParagraph();
            p.AppendText("Applicant Details\n");
            p.ApplyStyle(BuiltinStyle.Heading3);

            p = section.AddParagraph();
            AddField(p, "Applicant Name", "ApplicantName");
            AddField(p, "Account Number", "AccountNumber");

            p = section.AddParagraph();
            p.AppendText("Loan Details\n");
            p.ApplyStyle(BuiltinStyle.Heading3);

            p = section.AddParagraph();
            p.AppendText("Loan Amount: Rs. {LoanAmount}\n");
            p.AppendText("Purpose: {LoanPurpose}\n");
            p.AppendText("Tenure: {TenureMonths} months\n");
            p.AppendText("Interest Rate: {InterestRate}% p.a.\n");
            p.AppendText("Monthly EMI: Rs. {MonthlyEMI}\n");

            p = section.AddParagraph();
            p.AppendText("Branch: {BranchName}\n");
            p.AppendText("Application Date: {ApplicationDate}\n");

            p = section.AddParagraph();
            p.AppendText("Authorized Signature: _______________\n");

            SaveTemplate(document, path, "LoanProposal.docx");
            Console.WriteLine("✓ Generated: LoanProposal.docx");
        }
    }

    static void GeneratePromissoryNoteTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "PROMISSORY NOTE", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference No: {ReferenceNumber}                    Date: {GeneratedDate}\n");

            p = section.AddParagraph();
            p.AppendText("Borrower Details\n");
            p.ApplyStyle(BuiltinStyle.Heading3);

            p = section.AddParagraph();
            AddField(p, "Borrower Name", "BorrowerName");
            AddField(p, "CNIC", "BorrowerCNIC");
            AddField(p, "Address", "BorrowerAddress");
            AddField(p, "Account Number", "AccountNumber");

            p = section.AddParagraph();
            p.AppendText("Loan Details\n");
            p.ApplyStyle(BuiltinStyle.Heading3);

            p = section.AddParagraph();
            p.AppendText("Principal Amount: Rs. {PrincipalAmount} ({AmountInWords})\n");
            p.AppendText("Interest Rate: {InterestRate}% p.a.\n");
            p.AppendText("Loan Start Date: {LoanStartDate}\n");
            p.AppendText("Loan Tenure: {LoanTenureMonths} months\n");
            p.AppendText("Maturity Date: {MaturityDate}\n");
            p.AppendText("Monthly Installment: Rs. {MonthlyInstallment}\n");
            p.AppendText("Payment Mode: {PaymentMode}\n");

            p = section.AddParagraph();
            AddField(p, "Branch", "BranchName");
            AddField(p, "Bank", "BankName");
            AddField(p, "Bank Address", "BankAddress");

            p = section.AddParagraph();
            p.AppendText("Witness Details\n");
            p.ApplyStyle(BuiltinStyle.Heading3);

            p = section.AddParagraph();
            AddField(p, "Witness Name", "WitnessName");
            AddField(p, "Witness CNIC", "WitnessCNIC");

            p = section.AddParagraph();
            p.AppendText("Borrower Signature: _______________                    Witness Signature: _______________\n");

            SaveTemplate(document, path, "PromissoryNote.docx");
            Console.WriteLine("✓ Generated: PromissoryNote.docx");
        }
    }

    static void GenerateLetterOfContinuityTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "{BankName}", BuiltinStyle.Heading1);
            p = section.AddParagraph();
            p.AppendText("{BranchName}\n");
            p.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

            AddCenteredHeader(section, "LETTER OF CONTINUITY", BuiltinStyle.Heading2);

            p = section.AddParagraph();
            p.AppendText("Reference No: {ReferenceNumber}                    Date: {GeneratedDate}\n");

            p = section.AddParagraph();
            p.AppendText("Dear {ApplicantName},\n");

            p = section.AddParagraph();
            AddField(p, "Account Number", "AccountNumber");
            AddField(p, "Loan Account Number", "LoanAccountNumber");

            p = section.AddParagraph();
            p.AppendText("Facility Details\n");
            p.ApplyStyle(BuiltinStyle.Heading3);

            p = section.AddParagraph();
            AddField(p, "Facility Type", "FacilityType");
            p.AppendText("Outstanding Amount: Rs. {OutstandingAmount}\n");
            AddField(p, "Sanction Reference", "SanctionReference");
            AddField(p, "Sanction Date", "SanctionDate");

            p = section.AddParagraph();
            p.AppendText("Yours faithfully,\n");
            p.AppendText("For {BankName}\n");
            p.AppendText("Authorized Signature: _______________\n");

            SaveTemplate(document, path, "LetterOfContinuity.docx");
            Console.WriteLine("✓ Generated: LetterOfContinuity.docx");
        }
    }

    static void GenerateLetterOfRevivalTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "LETTER OF REVIVAL", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");
            p.AppendText("Dear {ApplicantName},\n");
            p.AppendText("We acknowledge the revival of your facility with the following details:\n");
            AddField(p, "Account Number", "AccountNumber");
            AddField(p, "Facility Type", "FacilityType");
            p.AppendText("Outstanding Amount: Rs. {OutstandingAmount}\n");
            AddField(p, "Revival Date", "RevivalDate");
            p.AppendText("For {BankName}\n");
            p.AppendText("Authorized Signature: _______________\n");

            SaveTemplate(document, path, "LetterOfRevival.docx");
            Console.WriteLine("✓ Generated: LetterOfRevival.docx");
        }
    }

    static void GenerateLetterOfArrangementTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "LETTER OF ARRANGEMENT", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");
            p.AppendText("Dear {ApplicantName},\n");
            p.AppendText("We are pleased to confirm the following arrangement:\n");
            AddField(p, "Account Number", "AccountNumber");
            AddField(p, "Arrangement Type", "ArrangementType");
            p.AppendText("Amount: Rs. {Amount}\n");
            AddField(p, "Effective Date", "EffectiveDate");
            p.AppendText("For {BankName}\n");
            p.AppendText("Authorized Signature: _______________\n");

            SaveTemplate(document, path, "LetterOfArrangement.docx");
            Console.WriteLine("✓ Generated: LetterOfArrangement.docx");
        }
    }

    static void GenerateStandingOrderTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "STANDING ORDER REQUEST", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");
            p.AppendText("Account Holder Details:\n");
            AddField(p, "Name", "AccountHolderName");
            AddField(p, "Account Number", "AccountNumber");
            AddField(p, "Account Type", "AccountType");
            p.AppendText("Payment Instructions:\n");
            p.AppendText("Amount: Rs. {Amount} ({AmountInWords})\n");
            AddField(p, "Frequency", "Frequency");
            AddField(p, "Start Date", "StartDate");
            AddField(p, "End Date", "EndDate");
            p.AppendText("Beneficiary Details:\n");
            AddField(p, "Name", "BeneficiaryName");
            AddField(p, "Account Number", "BeneficiaryAccountNumber");
            AddField(p, "Bank", "BeneficiaryBank");
            AddField(p, "Branch", "BranchName");
            p.AppendText("For {BankName}\n");
            p.AppendText("Authorized Signature: _______________\n");

            SaveTemplate(document, path, "StandingOrder.docx");
            Console.WriteLine("✓ Generated: StandingOrder.docx");
        }
    }

    static void GenerateLetterOfIndemnityTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "LETTER OF INDEMNITY", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");
            p.AppendText("To,\n");
            p.AppendText("{BankName}\n");
            p.AppendText("{BranchName}\n");
            p.AppendText("Dear Sir/Madam,\n");
            p.AppendText("We, {ApplicantName}, holding Account Number: {AccountNumber}, hereby submit this letter of indemnity regarding:\n");
            AddField(p, "Request Type", "RequestType");
            AddField(p, "Reference Details", "ReferenceDetails");
            p.AppendText("Indemnity Amount: Rs. {IndemnityAmount}\n");
            p.AppendText("We hereby indemnify the bank against any claims arising from this request.\n");
            p.AppendText("Yours faithfully,\n");
            p.AppendText("{ApplicantName}\n");
            p.AppendText("Signature: _______________\n");

            SaveTemplate(document, path, "LetterOfIndemnity.docx");
            Console.WriteLine("✓ Generated: LetterOfIndemnity.docx");
        }
    }

    static void GenerateLetterOfLienTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "LETTER OF LIEN", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");
            p.AppendText("Dear {ApplicantName},\n");
            p.AppendText("This is to confirm that the bank has a lien on the following account:\n");
            AddField(p, "Account Number", "AccountNumber");
            p.AppendText("Lien Amount: Rs. {LienAmount}\n");
            AddField(p, "Lien Reason", "LienReason");
            AddField(p, "Lien Start Date", "LienStartDate");
            AddField(p, "Branch", "BranchName");
            p.AppendText("For {BankName}\n");
            p.AppendText("Authorized Signature: _______________\n");

            SaveTemplate(document, path, "LetterOfLien.docx");
            Console.WriteLine("✓ Generated: LetterOfLien.docx");
        }
    }

    static void GeneratePersonalGuaranteeTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "PERSONAL GUARANTEE", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");
            p.AppendText("I, {GuarantorName}, holding CNIC: {GuarantorCNIC}, residing at:\n");
            p.AppendText("{GuarantorAddress}\n");
            p.AppendText("Hereby guarantee the obligations of:\n");
            AddField(p, "Borrower Name", "BorrowerName");
            AddField(p, "Borrower CNIC", "BorrowerCNIC");
            AddField(p, "Loan Account Number", "LoanAccountNumber");
            p.AppendText("Guarantee Details:\n");
            p.AppendText("Guaranteed Amount: Rs. {GuaranteedAmount} ({AmountInWords})\n");
            p.AppendText("I hereby undertake to pay the guaranteed amount upon demand by the bank.\n");
            AddField(p, "Branch", "BranchName");
            p.AppendText("Guarantor Signature: _______________\n");
            p.AppendText("Name: {GuarantorName}\n");
            p.AppendText("Date: {GeneratedDate}\n");

            SaveTemplate(document, path, "PersonalGuarantee.docx");
            Console.WriteLine("✓ Generated: PersonalGuarantee.docx");
        }
    }

    static void GenerateUDCTemplate(string path)
    {
        using (WordDocument document = new WordDocument())
        {
            IWSection section = document.AddSection();
            IWParagraph p;

            AddCenteredHeader(section, "UNDERTAKING OF DEBT CREATION (UDC)", BuiltinStyle.Heading1);

            p = section.AddParagraph();
            p.AppendText("Reference: {ReferenceNumber}                    Date: {GeneratedDate}\n");
            p.AppendText("To,\n");
            p.AppendText("{BankName}\n");
            p.AppendText("{BranchName}\n");
            p.AppendText("Dear Sir/Madam,\n");
            p.AppendText("We, {ApplicantName}, holding Account Number: {AccountNumber}, hereby request the creation of the following debt facility:\n");
            AddField(p, "Facility Type", "FacilityType");
            p.AppendText("Credit Limit: Rs. {CreditLimit}\n");
            p.AppendText("Security Details:\n");
            p.AppendText("{SecurityDetails}\n");
            p.AppendText("Sanction Details:\n");
            AddField(p, "Sanction Reference", "SanctionReference");
            AddField(p, "Sanction Date", "SanctionDate");
            p.AppendText("We hereby agree to abide by all terms and conditions of the facility.\n");
            p.AppendText("Yours faithfully,\n");
            p.AppendText("{ApplicantName}\n");
            p.AppendText("Signature: _______________\n");

            SaveTemplate(document, path, "UDC.docx");
            Console.WriteLine("✓ Generated: UDC.docx");
        }
    }
}
