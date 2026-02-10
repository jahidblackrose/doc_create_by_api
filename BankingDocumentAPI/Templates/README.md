# Templates Folder

Place your Word document templates in this folder.

## Required Templates:

1. **LoanProposal.docx** - Loan Proposal Document
2. **PromissoryNote.docx** - Promissory Note Document
3. **LetterOfContinuity.docx** - Letter of Continuity Document
4. **LetterOfRevival.docx** - Letter of Revival Document
5. **LetterOfArrangement.docx** - Letter of Arrangement Document
6. **StandingOrder.docx** - Standing Order Document
7. **LetterOfIndemnity.docx** - Letter of Indemnity Document
8. **LetterOfLien.docx** - Letter of Lien Document
9. **PersonalGuarantee.docx** - Personal Guarantee Document
10. **UDC.docx** - UDC (Undertaking of Debt Creation) Document

## Template Format:

Templates should be Word documents (.docx) with placeholders in the format `{PropertyName}`.

For example, a LoanProposal.docx template might contain:

```
                    {BankName}
                {BankAddress}

                    LOAN PROPOSAL

Reference No: {ReferenceNumber}
Date: {GeneratedDate}

Applicant Name: {ApplicantName}
Account Number: {AccountNumber}

Loan Details:
- Loan Amount: Rs. {LoanAmount}
- Purpose: {LoanPurpose}
- Tenure: {TenureMonths} months
- Interest Rate: {InterestRate}% p.a.
- Monthly EMI: Rs. {MonthlyEMI}

Branch: {BranchName}
Application Date: {ApplicationDate}
```

## Available Properties for Each Template:

### LoanProposal.docx:
- BankName, BankAddress
- ApplicantName, AccountNumber
- LoanAmount, LoanPurpose, TenureMonths
- InterestRate, MonthlyEMI
- BranchName, ApplicationDate
- GeneratedDate, ReferenceNumber

### PromissoryNote.docx:
- BankName, BankAddress
- BorrowerName, BorrowerAddress, BorrowerCNIC
- AccountNumber
- PrincipalAmount, AmountInWords
- InterestRate, LoanStartDate, LoanTenureMonths
- MaturityDate, MonthlyInstallment
- PaymentMode, BranchName
- GeneratedDate, ReferenceNumber
- WitnessName, WitnessCNIC

### LetterOfContinuity.docx:
- BankName, BranchName
- ApplicantName, AccountNumber, LoanAccountNumber
- OutstandingAmount, FacilityType
- SanctionReference, SanctionDate
- GeneratedDate, ReferenceNumber

### StandingOrder.docx:
- BankName, BranchName
- AccountHolderName, AccountNumber, AccountType
- Amount, AmountInWords, Frequency
- StartDate, EndDate
- BeneficiaryName, BeneficiaryAccountNumber, BeneficiaryBank
- GeneratedDate, ReferenceNumber

### LetterOfIndemnity.docx:
- BankName, BranchName
- ApplicantName, AccountNumber
- RequestType, ReferenceDetails
- IndemnityAmount
- GeneratedDate, ReferenceNumber

### PersonalGuarantee.docx:
- BankName, BranchName
- GuarantorName, GuarantorCNIC, GuarantorAddress
- BorrowerName, BorrowerCNIC
- GuaranteedAmount, AmountInWords
- LoanAccountNumber
- GeneratedDate, ReferenceNumber

### UDC.docx:
- BankName, BranchName
- ApplicantName, AccountNumber
- FacilityType, CreditLimit
- SecurityDetails, SanctionDate, SanctionReference
- GeneratedDate, ReferenceNumber

## Note:

If templates are not found, the API will automatically generate basic documents with all available data.
