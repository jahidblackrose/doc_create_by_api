namespace BankingDocumentAPI.Models
{
    public class DocumentRequest
    {
        public long LoanId { get; set; }
        public DocumentType DocumentType { get; set; }
        public OutputFormat OutputFormat { get; set; } = OutputFormat.PDF;
    }

    public enum DocumentType
    {
        LoanProposal,
        PromissoryNote,
        LetterOfContinuity,
        LetterOfRevival,
        LetterOfArrangement,
        StandingOrder,
        LetterOfIndemnity,
        LetterOfLien,
        PersonalGuarantee,
        UDC
    }

    public enum OutputFormat
    {
        DOCX,
        PDF
    }
}
