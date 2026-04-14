using WebAPI.Models;
using WebAPI.Models.SalesProduction;
using WorksheetAPI.Models;

namespace WorksheetAPI.Interfaces;

/// <summary>
/// Handles printing workflows and cached print document access for sales production.
/// </summary>
public interface ISalesProductionPrintService
{
    Task<object> GeneratePrintProductionOrdersAsync(
        SapCompany company,
        CancelAllProductionsRequest request,
        string baseApiUrl);

    bool IsValidGeneratedPrintDocumentId(string? documentId);

    bool TryGetGeneratedPrint(string documentId, out GeneratedPrintDocument? document);

    string? BuildGeneratedPrintDownloadUrl(string baseApiUrl, string documentId);

    string? BuildGeneratedPrintOpenUrl(string baseApiUrl, string documentId);

    string? BuildGeneratedPrintStatusUrl(string baseApiUrl, string documentId);
}
