namespace BankingDocumentAPI.Models
{
    public class StandingOrderData
    {
        public string AccountHolderName { get; set; } = string.Empty;
        public string AccountNumber { get; set; } = string.Empty;
        public string AccountType { get; set; } = string.Empty;
        public decimal Amount { get; set; }
        public string AmountInWords { get; set; } = string.Empty;
        public string Frequency { get; set; } = "Monthly";
        public DateTime StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string BeneficiaryName { get; set; } = string.Empty;
        public string BeneficiaryAccountNumber { get; set; } = string.Empty;
        public string BeneficiaryBank { get; set; } = string.Empty;

        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BranchName { get; set; } = string.Empty;
        public string GeneratedDate { get; set; } = string.Empty;
        public string ReferenceNumber { get; set; } = string.Empty;
    }

    public class LetterOfIndemnityData
    {
        public string ApplicantName { get; set; } = string.Empty;
        public string AccountNumber { get; set; } = string.Empty;
        public string RequestType { get; set; } = string.Empty; // e.g., "Lost Cheque", "Lost Document"
        public string ReferenceDetails { get; set; } = string.Empty;
        public decimal IndemnityAmount { get; set; }

        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BranchName { get; set; } = string.Empty;
        public string GeneratedDate { get; set; } = string.Empty;
        public string ReferenceNumber { get; set; } = string.Empty;
    }

    public class PersonalGuaranteeData
    {
        public string GuarantorName { get; set; } = string.Empty;
        public string GuarantorCNIC { get; set; } = string.Empty;
        public string GuarantorAddress { get; set; } = string.Empty;
        public string BorrowerName { get; set; } = string.Empty;
        public string BorrowerCNIC { get; set; } = string.Empty;
        public decimal GuaranteedAmount { get; set; }
        public string AmountInWords { get; set; } = string.Empty;
        public string LoanAccountNumber { get; set; } = string.Empty;

        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BranchName { get; set; } = string.Empty;
        public string GeneratedDate { get; set; } = string.Empty;
        public string ReferenceNumber { get; set; } = string.Empty;
    }

    public class UDCData
    {
        public string ApplicantName { get; set; } = string.Empty;
        public string AccountNumber { get; set; } = string.Empty;
        public string FacilityType { get; set; } = string.Empty;
        public decimal CreditLimit { get; set; }
        public string SecurityDetails { get; set; } = string.Empty;
        public DateTime SanctionDate { get; set; }
        public string SanctionReference { get; set; } = string.Empty;

        public string BankName { get; set; } = "ABC Bank Ltd.";
        public string BranchName { get; set; } = string.Empty;
        public string GeneratedDate { get; set; } = string.Empty;
        public string ReferenceNumber { get; set; } = string.Empty;
    }
}
