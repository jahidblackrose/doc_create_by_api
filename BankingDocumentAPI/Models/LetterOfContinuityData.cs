namespace BankingDocumentAPI.Models
{
    public class LetterOfContinuityData
    {
        public string ApplicantName { get; set; } = string.Empty;
        public string AccountNumber { get; set; } = string.Empty;
        public string LoanAccountNumber { get; set; } = string.Empty;
        public decimal OutstandingAmount { get; set; }
        public string FacilityType { get; set; } = string.Empty;
        public string SanctionReference { get; set; } = string.Empty;
        public DateTime SanctionDate { get; set; }

        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BranchName { get; set; } = string.Empty;
        public string GeneratedDate { get; set; } = string.Empty;
        public string ReferenceNumber { get; set; } = string.Empty;
    }
}
