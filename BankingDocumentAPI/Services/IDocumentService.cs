using BankingDocumentAPI.Models;

namespace BankingDocumentAPI.Services
{
    public interface IDocumentService
    {
        Task<byte[]> GenerateDocumentAsync(DocumentRequest request);
        Task<byte[]> GenerateWordDocumentAsync(string templateName, object data);
        Task<byte[]> GeneratePdfFromWordAsync(byte[] wordDocument);
    }
}
