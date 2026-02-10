using BankingDocumentAPI.Models;
using BankingDocumentAPI.Services;
using Microsoft.AspNetCore.Mvc;

namespace BankingDocumentAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentController : ControllerBase
    {
        private readonly IDocumentService _documentService;
        private readonly ILogger<DocumentController> _logger;

        public DocumentController(
            IDocumentService documentService,
            ILogger<DocumentController> logger)
        {
            _documentService = documentService;
            _logger = logger;
        }

        /// <summary>
        /// Generate a single document
        /// </summary>
        /// <param name="request">Document generation request</param>
        /// <returns>Generated document file</returns>
        [HttpPost("generate")]
        [ProducesResponseType(typeof(FileContentResult), 200)]
        [ProducesResponseType(500)]
        public async Task<IActionResult> GenerateDocument([FromBody] DocumentRequest request)
        {
            try
            {
                var documentBytes = await _documentService.GenerateDocumentAsync(request);

                string fileName = $"{request.DocumentType}_{request.LoanId}_{DateTime.Now:yyyyMMddHHmmss}";
                string contentType = request.OutputFormat == OutputFormat.PDF
                    ? "application/pdf"
                    : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                string extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";

                return File(documentBytes, contentType, $"{fileName}.{extension}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating document");
                return StatusCode(500, new { error = "Failed to generate document", message = ex.Message });
            }
        }

        /// <summary>
        /// Download document via GET request
        /// </summary>
        /// <param name="loanId">Loan ID</param>
        /// <param name="documentType">Type of document</param>
        /// <param name="format">Output format (PDF or DOCX)</param>
        /// <returns>Generated document file</returns>
        [HttpGet("download/{loanId}/{documentType}")]
        [ProducesResponseType(typeof(FileContentResult), 200)]
        [ProducesResponseType(500)]
        public async Task<IActionResult> DownloadDocument(
            long loanId,
            DocumentType documentType,
            [FromQuery] OutputFormat format = OutputFormat.PDF)
        {
            var request = new DocumentRequest
            {
                LoanId = loanId,
                DocumentType = documentType,
                OutputFormat = format
            };

            return await GenerateDocument(request);
        }

        /// <summary>
        /// Generate multiple documents and return as ZIP
        /// </summary>
        /// <param name="requests">List of document generation requests</param>
        /// <returns>ZIP file containing all documents</returns>
        [HttpPost("batch-generate")]
        [ProducesResponseType(typeof(FileContentResult), 200)]
        [ProducesResponseType(500)]
        public async Task<IActionResult> GenerateBatchDocuments([FromBody] List<DocumentRequest> requests)
        {
            try
            {
                var tasks = requests.Select(r => _documentService.GenerateDocumentAsync(r));
                var results = await Task.WhenAll(tasks);

                // Create ZIP file with all documents
                using (var memoryStream = new MemoryStream())
                {
                    using (var archive = new System.IO.Compression.ZipArchive(memoryStream, System.IO.Compression.ZipArchiveMode.Create, true))
                    {
                        for (int i = 0; i < requests.Count; i++)
                        {
                            var request = requests[i];
                            var extension = request.OutputFormat == OutputFormat.PDF ? "pdf" : "docx";
                            var fileName = $"{request.DocumentType}_{request.LoanId}.{extension}";

                            var entry = archive.CreateEntry(fileName);
                            using (var entryStream = entry.Open())
                            {
                                await entryStream.WriteAsync(results[i], 0, results[i].Length);
                            }
                        }
                    }

                    return File(memoryStream.ToArray(),
                        "application/zip",
                        $"Documents_{DateTime.Now:yyyyMMddHHmmss}.zip");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating batch documents");
                return StatusCode(500, new { error = "Failed to generate documents", message = ex.Message });
            }
        }

        /// <summary>
        /// Get all available document types
        /// </summary>
        /// <returns>List of document types</returns>
        [HttpGet("document-types")]
        [ProducesResponseType(typeof(IEnumerable<string>), 200)]
        public IActionResult GetDocumentTypes()
        {
            var types = Enum.GetNames(typeof(DocumentType));
            return Ok(new
            {
                documentTypes = types,
                count = types.Length
            });
        }

        /// <summary>
        /// Health check endpoint
        /// </summary>
        /// <returns>Service status</returns>
        [HttpGet("health")]
        [ProducesResponseType(200)]
        public IActionResult HealthCheck()
        {
            return Ok(new
            {
                status = "Healthy",
                timestamp = DateTime.Now,
                service = "Banking Document API"
            });
        }
    }
}
