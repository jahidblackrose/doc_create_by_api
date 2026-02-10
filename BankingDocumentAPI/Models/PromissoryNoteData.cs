namespace BankingDocumentAPI.Models
{
    public class PromissoryNoteData
    {
        // Borrower Details
        public string BorrowerName { get; set; } = string.Empty;
        public string BorrowerAddress { get; set; } = string.Empty;
        public string BorrowerCNIC { get; set; } = string.Empty;
        public string AccountNumber { get; set; } = string.Empty;

        // Loan Details
        public decimal PrincipalAmount { get; set; }
        public string AmountInWords { get; set; } = string.Empty;
        public decimal InterestRate { get; set; }
        public DateTime LoanStartDate { get; set; }
        public int LoanTenureMonths { get; set; }
        public DateTime MaturityDate { get; set; }

        // Repayment Details
        public decimal MonthlyInstallment { get; set; }
        public string PaymentMode { get; set; } = "Monthly Post-Dated Cheques";
        public string BranchName { get; set; } = string.Empty;

        // Bank Details
        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BankAddress { get; set; } = "123 Banking Street, Main Branch, City";

        // Document Details
        public string GeneratedDate { get; set; } = string.Empty;
        public string ReferenceNumber { get; set; } = string.Empty;
        public string WitnessName { get; set; } = string.Empty;
        public string WitnessCNIC { get; set; } = string.Empty;
    }
}
