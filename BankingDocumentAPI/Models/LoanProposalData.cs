namespace BankingDocumentAPI.Models
{
    public class LoanProposalData
    {
        // From Database
        public string ApplicantName { get; set; } = string.Empty;
        public string AccountNumber { get; set; } = string.Empty;
        public decimal LoanAmount { get; set; }
        public string LoanPurpose { get; set; } = string.Empty;
        public int TenureMonths { get; set; }
        public decimal InterestRate { get; set; }
        public DateTime ApplicationDate { get; set; }
        public string BranchName { get; set; } = string.Empty;

        // Static/Calculated
        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BankAddress { get; set; } = "123 Banking Street, Main Branch, City";
        public decimal MonthlyEMI { get; set; }
        public string GeneratedDate { get; set; } = string.Empty;
        public string ReferenceNumber { get; set; } = string.Empty;
    }
}
