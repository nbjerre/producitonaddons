using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using Newtonsoft.Json.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using WebAPI.Configuration;
using WebAPI.Models;
using WebAPI.Models.SalesProduction;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;

namespace WorksheetAPI.Services;

/// <summary>
/// Encapsulates print generation and cached print document retrieval for sales production flows.
/// </summary>
public class SalesProductionPrintService : ISalesProductionPrintService
{
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<SalesProductionPrintService> _logger;
    private readonly PrintSettings _printSettings;
    private readonly IMemoryCache _memoryCache;
    private readonly IHttpClientFactory _httpClientFactory;
    private static readonly TimeSpan GeneratedPrintLifetime = TimeSpan.FromMinutes(15);

    public SalesProductionPrintService(
        IServiceLayerService serviceLayer,
        ILogger<SalesProductionPrintService> logger,
        IMemoryCache memoryCache,
        IOptions<PrintSettings> printSettingsOptions,
        IHttpClientFactory httpClientFactory)
    {
        _serviceLayer = serviceLayer;
        _logger = logger;
        _memoryCache = memoryCache;
        _printSettings = printSettingsOptions.Value ?? new PrintSettings();
        _httpClientFactory = httpClientFactory;
    }

    public async Task<object> GeneratePrintProductionOrdersAsync(
        SapCompany company,
        CancelAllProductionsRequest request,
        string baseApiUrl)
    {
        _logger.LogInformation(
            "PrintProductionOrders request received. Input SalesOrderDocEntry={SalesOrderDocEntry}, SalesOrderDocNum={SalesOrderDocNum}",
            request.SalesOrderDocEntry,
            request.SalesOrderDocNum);

        var salesOrderDocEntry = request.SalesOrderDocEntry;
        if (salesOrderDocEntry <= 0)
        {
            _logger.LogInformation(
                "PrintProductionOrders resolving sales order DocEntry from DocNum {SalesOrderDocNum}",
                request.SalesOrderDocNum);

            salesOrderDocEntry = await ResolveDocEntryByDocNum(company, request.SalesOrderDocNum ?? 0);
            if (salesOrderDocEntry <= 0)
                return new { success = false, error = "Could not resolve sales order from DocNum" };
        }

        _logger.LogInformation(
            "PrintProductionOrders loading sales order {SalesOrderDocEntry}",
            salesOrderDocEntry);

        var orderResult = await _serviceLayer.GetStringAsync(company, $"Orders({salesOrderDocEntry})");
        if (!orderResult.Success || string.IsNullOrEmpty(orderResult.Data))
            return new { success = false, error = orderResult.Error ?? "Could not read sales order" };

        var orderObj = JToken.Parse(orderResult.Data) as JObject;
        if (orderObj == null)
            return new { success = false, error = "Invalid sales order payload from SAP" };

        var salesOrderDocNum = GetInt(orderObj, "DocNum");
        var cardCode = GetString(orderObj, "CardCode") ?? string.Empty;

        _logger.LogInformation(
            "PrintProductionOrders loaded sales order {SalesOrderDocEntry}/{SalesOrderDocNum} for customer {CardCode}",
            salesOrderDocEntry,
            salesOrderDocNum,
            cardCode);

        var lines = orderObj["DocumentLines"] as JArray;
        var printableLineNums = new HashSet<int>();
        var salesOrderLineInfo = new Dictionary<int, JObject>();
        if (lines != null)
        {
            foreach (var line in lines.OfType<JObject>())
            {
                var lineNum = GetInt(line, "LineNum");
                salesOrderLineInfo[lineNum] = line;
                var onSto = (GetString(line, "U_RCS_ONSTO") ?? "Y").ToUpperInvariant();
                if (onSto != "N")
                {
                    printableLineNums.Add(lineNum);
                }
            }
        }

        if (printableLineNums.Count == 0)
        {
            return new
            {
                success = true,
                printableCount = 0,
                renderMode = "native",
                message = "Ingen produktionslinjer markeret til print (U_RCS_ONSTO = Y)."
            };
        }

        var filter =
            $"ProductionOrderOriginEntry eq {salesOrderDocEntry} " +
            $"and (ProductionOrderStatus eq 'boposPlanned' or ProductionOrderStatus eq 'boposReleased')";

        var listResult = await _serviceLayer.GetStringAsync(
            company,
            "ProductionOrders",
            select: "AbsoluteEntry,DocumentNumber,ItemNo,U_RCS_OL,ProductionOrderStatus,ProductionOrderOriginEntry,ProductionOrderOriginNumber",
            filter: filter,
            top: 200);

        if (!listResult.Success || string.IsNullOrEmpty(listResult.Data))
            return new { success = false, error = listResult.Error ?? "Could not fetch production orders" };

        var root = JToken.Parse(listResult.Data) as JObject;
        var orders = root?["value"] as JArray;

        if (orders == null || orders.Count == 0)
        {
            var fallbackResult = await _serviceLayer.GetStringAsync(
                company,
                "ProductionOrders",
                filter: "ProductionOrderStatus eq 'boposPlanned' or ProductionOrderStatus eq 'boposReleased'",
                top: 500);

            if (fallbackResult.Success && !string.IsNullOrWhiteSpace(fallbackResult.Data))
            {
                var fallbackRoot = JToken.Parse(fallbackResult.Data) as JObject;
                var fallbackOrders = fallbackRoot?["value"] as JArray;

                if (fallbackOrders != null && fallbackOrders.Count > 0)
                {
                    orders = new JArray(
                        fallbackOrders
                            .OfType<JObject>()
                            .Where(order => ProductionOrderMatchesSalesOrder(order, salesOrderDocEntry, salesOrderDocNum, printableLineNums, salesOrderLineInfo)));
                }
            }
        }

        if (orders == null || orders.Count == 0)
        {
            _logger.LogWarning(
                "PrintProductionOrders found no open/released orders for sales order {SalesOrderDocEntry}/{SalesOrderDocNum}. Printable lines: {PrintableLines}",
                salesOrderDocEntry,
                salesOrderDocNum,
                string.Join(",", printableLineNums.OrderBy(x => x)));

            return new
            {
                success = true,
                printableCount = 0,
                renderMode = "native",
                message = "Ingen åbne/frigivne produktionsordrer fundet."
            };
        }

        var printable = orders
            .OfType<JObject>()
            .Select(order => new
            {
                ProductionDocEntry = GetInt(order, "AbsoluteEntry"),
                DocumentNumber = GetInt(order, "DocumentNumber"),
                ItemCode = GetString(order, "ItemNo") ?? string.Empty,
                OrderLine = GetInt(order, "U_RCS_OL"),
                Status = GetString(order, "ProductionOrderStatus") ?? string.Empty,
                OriginEntry = GetInt(order, "ProductionOrderOriginEntry"),
                OriginNumber = GetInt(order, "ProductionOrderOriginNumber")
            })
            .Where(x => x.ProductionDocEntry > 0 && IsPrintableProductionOrder(x.OriginEntry, x.OriginNumber, x.OrderLine, salesOrderDocEntry, salesOrderDocNum, printableLineNums))
            .OrderBy(x => x.DocumentNumber)
            .ToList();

        if (printable.Count == 0)
        {
            return new
            {
                success = true,
                salesOrderDocEntry,
                printableCount = 0,
                renderMode = "native",
                message = "Ingen produktionsordrer opfylder print-kriterierne."
            };
        }

        var fileStem = $"printprod_so_{salesOrderDocEntry}_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}";
        var generatedFiles = new List<GeneratedPrintFile>();
        var generatedOrders = new List<object>();
        var failedOrders = new List<object>();
        var attachmentWarnings = new List<object>();
        var attachmentCache = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var anyCrystalRender = false;
        var productionReportCode = string.IsNullOrWhiteSpace(_printSettings.ProductionReportCode)
            ? "WOR10003"
            : _printSettings.ProductionReportCode.Trim();

        if (string.IsNullOrWhiteSpace(_printSettings.CrystalBaseUrl))
        {
            return new
            {
                success = false,
                salesOrderDocEntry,
                printableCount = printable.Count,
                generatedCount = 0,
                failedCount = printable.Count,
                failedOrders = printable.Select(order => new
                {
                    productionDocEntry = order.ProductionDocEntry,
                    documentNumber = order.DocumentNumber,
                    itemCode = order.ItemCode,
                    error = "CrystalBaseUrl er ikke konfigureret. Native fallback er deaktiveret."
                }).ToList(),
                renderMode = "crystal",
                message = $"Crystal-print er påkrævet, men ikke konfigureret. Rapportkode: {productionReportCode}."
            };
        }

        foreach (var order in printable)
        {
            try
            {
                var orderFiles = new List<GeneratedPrintFile>();

                var crystalRender = await TryRenderProductionOrderViaCrystalAsync(company, order.ProductionDocEntry);
                if (!crystalRender.Success || crystalRender.PdfBytes == null)
                {
                    throw new InvalidOperationException(
                    $"Crystal-render fejlede for produktionsordre {order.ProductionDocEntry}. Rapportkode: {productionReportCode}. {crystalRender.ErrorMessage}".Trim());
                }

                var renderMode = "crystal";
                anyCrystalRender = true;

                orderFiles.Add(new GeneratedPrintFile
                {
                    FileName = CreateOrderPdfFileName(order.DocumentNumber, order.ItemCode),
                    Content = crystalRender.PdfBytes,
                    ContentType = "application/pdf"
                });

                var includedAttachmentCount = 0;

                if (!attachmentCache.TryGetValue(order.ItemCode, out var attachmentPaths))
                {
                    attachmentPaths = await GetPdfAttachmentsForItemAsync(company, order.ItemCode);
                    attachmentCache[order.ItemCode] = attachmentPaths;
                }

                foreach (var attachmentPath in attachmentPaths)
                {
                    if (System.IO.File.Exists(attachmentPath))
                    {
                        if (IsPdfFile(attachmentPath))
                        {
                            var attachmentBytes = await System.IO.File.ReadAllBytesAsync(attachmentPath);
                            using var attachmentStream = new MemoryStream(attachmentBytes, writable: false);
                            if (attachmentBytes.Length > 0 && IsPdfStream(attachmentStream))
                            {
                                orderFiles.Add(new GeneratedPrintFile
                                {
                                    FileName = CreateAttachmentFileName(order.DocumentNumber, order.ItemCode, attachmentPath, includedAttachmentCount + 1),
                                    Content = attachmentBytes,
                                    ContentType = "application/pdf"
                                });
                                includedAttachmentCount++;
                            }
                            else
                            {
                                attachmentWarnings.Add(new
                                {
                                    productionDocEntry = order.ProductionDocEntry,
                                    itemCode = order.ItemCode,
                                    path = attachmentPath,
                                    reason = "INVALID_PDF"
                                });
                            }
                        }
                        else
                        {
                            attachmentWarnings.Add(new
                            {
                                productionDocEntry = order.ProductionDocEntry,
                                itemCode = order.ItemCode,
                                path = attachmentPath,
                                reason = "INVALID_PDF"
                            });
                        }
                    }
                    else
                    {
                        attachmentWarnings.Add(new
                        {
                            productionDocEntry = order.ProductionDocEntry,
                            itemCode = order.ItemCode,
                            path = attachmentPath,
                            reason = "FILE_NOT_FOUND"
                        });
                    }
                }

                generatedFiles.AddRange(orderFiles);
                generatedOrders.Add(new
                {
                    productionDocEntry = order.ProductionDocEntry,
                    documentNumber = order.DocumentNumber,
                    itemCode = order.ItemCode,
                    orderLine = order.OrderLine,
                    status = order.Status,
                    renderMode,
                    attachmentCount = includedAttachmentCount,
                    measureIncluded = false
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex,
                    "Print generation failed for productionDocEntry {ProductionDocEntry}",
                    order.ProductionDocEntry);

                failedOrders.Add(new
                {
                    productionDocEntry = order.ProductionDocEntry,
                    documentNumber = order.DocumentNumber,
                    itemCode = order.ItemCode,
                    error = ex.Message
                });
            }
        }

        if (generatedFiles.Count == 0)
        {
            return new
            {
                success = false,
                salesOrderDocEntry,
                printableCount = printable.Count,
                generatedCount = 0,
                failedCount = failedOrders.Count,
                failedOrders,
                renderMode = "crystal",
                message = $"Ingen PDF-filer kunne genereres via Crystal. Rapportkode: {productionReportCode}."
            };
        }

        var finalDocument = BuildGeneratedPrintDocument(fileStem, generatedFiles);
        var documentId = StoreGeneratedPrint(finalDocument, request.DocumentId);
        var downloadUrl = BuildGeneratedPrintDownloadUrl(baseApiUrl, documentId);
        var openUrl = BuildGeneratedPrintOpenUrl(baseApiUrl, documentId);

        return new
        {
            success = true,
            salesOrderDocEntry,
            printableCount = printable.Count,
            generatedCount = generatedOrders.Count,
            failedCount = failedOrders.Count,
            attachmentWarningsCount = attachmentWarnings.Count,
            includedMeasureReport = false,
            renderMode = anyCrystalRender ? "crystal" : "native",
            documentId,
            openUrl,
            downloadUrl,
            contentType = finalDocument.ContentType,
            openInline = finalDocument.OpenInline,
            fileName = finalDocument.FileName,
            orders = generatedOrders,
            failedOrders,
            message = $"Klar til print: {generatedOrders.Count} produktionsordre(r) samlet i én PDF."
        };
    }

    public string? BuildGeneratedPrintDownloadUrl(string baseApiUrl, string documentId)
    {
        if (string.IsNullOrWhiteSpace(documentId))
            return null;

        return $"{baseApiUrl.TrimEnd('/')}/api/sales-production/print-download/{Uri.EscapeDataString(documentId)}";
    }

    public string? BuildGeneratedPrintOpenUrl(string baseApiUrl, string documentId)
    {
        if (string.IsNullOrWhiteSpace(documentId))
            return null;

        return $"{baseApiUrl.TrimEnd('/')}/api/sales-production/print-open/{Uri.EscapeDataString(documentId)}";
    }

    public string? BuildGeneratedPrintStatusUrl(string baseApiUrl, string documentId)
    {
        if (string.IsNullOrWhiteSpace(documentId))
            return null;

        return $"{baseApiUrl.TrimEnd('/')}/api/sales-production/print-status/{Uri.EscapeDataString(documentId)}";
    }

    public bool TryGetGeneratedPrint(string documentId, out GeneratedPrintDocument? document)
    {
        document = null;
        if (!IsValidGeneratedPrintDocumentId(documentId))
            return false;

        return _memoryCache.TryGetValue(BuildGeneratedPrintCacheKey(documentId), out document);
    }

    public bool IsValidGeneratedPrintDocumentId(string? documentId)
    {
        return !string.IsNullOrWhiteSpace(documentId)
            && Regex.IsMatch(documentId, "^[A-Za-z0-9_-]+$");
    }

    private async Task<int> ResolveDocEntryByDocNum(SapCompany company, int docNum)
    {
        if (docNum <= 0)
            return 0;

        var result = await _serviceLayer.GetStringAsync(company, "Orders", select: "DocEntry", filter: $"DocNum eq {docNum}", top: 1);
        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return 0;

        var root = JToken.Parse(result.Data) as JObject;
        var arr = root?["value"] as JArray;
        var row = arr?.OfType<JObject>().FirstOrDefault();
        return GetInt(row, "DocEntry");
    }

    private async Task<List<string>> GetPdfAttachmentsForItemAsync(SapCompany company, string itemCode)
    {
        var attachmentPaths = new List<string>();
        if (string.IsNullOrWhiteSpace(itemCode))
            return attachmentPaths;

        var itemResult = await _serviceLayer.GetStringAsync(
            company,
            $"Items('{EscapeODataString(itemCode)}')",
            select: "ItemCode,AttachmentEntry");

        if (!itemResult.Success || string.IsNullOrWhiteSpace(itemResult.Data))
        {
            _logger.LogWarning("Could not resolve attachment entry for item {ItemCode}: {Error}", itemCode, itemResult.Error);
            return attachmentPaths;
        }

        var itemObject = JToken.Parse(itemResult.Data) as JObject;
        var attachmentEntry = GetInt(itemObject, "AttachmentEntry");
        if (attachmentEntry <= 0)
            return attachmentPaths;

        var attachmentResult = await _serviceLayer.GetStringAsync(company, $"Attachments2({attachmentEntry})");
        if (!attachmentResult.Success || string.IsNullOrWhiteSpace(attachmentResult.Data))
        {
            _logger.LogWarning("Could not load Attachments2 {AttachmentEntry}: {Error}", attachmentEntry, attachmentResult.Error);
            return attachmentPaths;
        }

        var attachmentObject = JToken.Parse(attachmentResult.Data) as JObject;
        var lines = attachmentObject?["Attachments2_Lines"] as JArray
            ?? attachmentObject?["Attachments2Lines"] as JArray;

        if (lines == null)
            return attachmentPaths;

        foreach (var line in lines.OfType<JObject>())
        {
            var extension = (GetString(line, "FileExtension") ?? GetString(line, "FileExt") ?? string.Empty).Trim().TrimStart('.');
            if (!extension.Equals("PDF", StringComparison.OrdinalIgnoreCase))
                continue;

            var sourcePath = GetString(line, "SourcePath") ?? GetString(line, "trgtPath") ?? string.Empty;
            var fileName = GetString(line, "FileName") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sourcePath) || string.IsNullOrWhiteSpace(fileName))
                continue;

            var fullPath = Path.Combine(sourcePath, $"{fileName}.{extension}");
            attachmentPaths.Add(fullPath);
        }

        return attachmentPaths;
    }

    private async Task<(bool Success, byte[]? PdfBytes, string? ErrorMessage)> TryRenderProductionOrderViaCrystalAsync(SapCompany company, int productionDocEntry)
    {
        if (string.IsNullOrWhiteSpace(_printSettings.CrystalBaseUrl))
            return (false, null, "CrystalBaseUrl er ikke konfigureret.");

        var apiRenderResult = await TryRenderProductionOrderViaCrystalApiAsync(company, productionDocEntry);
        if (!apiRenderResult.Success)
        {
            _logger.LogWarning(
                "Crystal API render failed for productionDocEntry {ProductionDocEntry}. Error {Error}",
                productionDocEntry,
                apiRenderResult.ErrorMessage);
        }

        return apiRenderResult;
    }

    private async Task<(bool Success, byte[]? PdfBytes, string? ErrorMessage)> TryRenderProductionOrderViaCrystalApiAsync(
        SapCompany company,
        int productionDocEntry)
    {
        var renderApiUrl = BuildCrystalRenderApiUrl();
        if (string.IsNullOrWhiteSpace(renderApiUrl))
            return (false, null, "Crystal API-url kunne ikke bygges.");

        try
        {
            var client = _httpClientFactory.CreateClient("crystal");
            client.Timeout = TimeSpan.FromSeconds(10);

            var reportCode = string.IsNullOrWhiteSpace(_printSettings.ProductionReportCode)
                ? "WOR10003"
                : _printSettings.ProductionReportCode.Trim();

            var payload = new JObject
            {
                ["database"] = company.CompanyDb,
                ["reportCode"] = reportCode,
                ["docKey"] = productionDocEntry
            };

            if (!reportCode.StartsWith("WOR", StringComparison.OrdinalIgnoreCase))
                payload["objectId"] = _printSettings.ProductionObjectId;

            _logger.LogInformation(
                "Crystal API render POST started for productionDocEntry {ProductionDocEntry}. Url {Url}. ReportCode {ReportCode}",
                productionDocEntry,
                renderApiUrl,
                reportCode);

            using var response = await client.PostAsync(
                renderApiUrl,
                new StringContent(payload.ToString(), Encoding.UTF8, "application/json"));

            if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                return (false, null, "Crystal API endpoint ikke fundet.");

            if (!response.IsSuccessStatusCode)
            {
                var errorBody = await response.Content.ReadAsStringAsync();
                var detailedMessage = ExtractCrystalHtmlMessage(errorBody) ?? errorBody;
                if (!string.IsNullOrWhiteSpace(detailedMessage) && detailedMessage.Length > 500)
                    detailedMessage = detailedMessage[..500].Trim();

                return (false, null, string.IsNullOrWhiteSpace(detailedMessage)
                    ? $"Crystal API svarede med HTTP {(int)response.StatusCode}."
                    : detailedMessage);
            }

            var pdfBytes = await TryReadPdfResponseAsync(response);
            if (pdfBytes != null)
            {
                _logger.LogInformation(
                    "Crystal API render POST succeeded for productionDocEntry {ProductionDocEntry}. Url {Url}",
                    productionDocEntry,
                    renderApiUrl);
                return (true, pdfBytes, null);
            }

            var responseBody = await response.Content.ReadAsStringAsync();
            var mediaType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;
            var message = ExtractCrystalHtmlMessage(responseBody);

            return (false, null, string.IsNullOrWhiteSpace(message)
                ? $"Crystal API returnerede {mediaType} i stedet for PDF."
                : message);
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogWarning(ex,
                "Crystal API render timeout for productionDocEntry {ProductionDocEntry}. Url {Url}",
                productionDocEntry,
                renderApiUrl);

            return (false, null, "Crystal API timeout efter 10 sekunder.");
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex,
                "Crystal API render failed for productionDocEntry {ProductionDocEntry}. Url {Url}",
                productionDocEntry,
                renderApiUrl);

            return (false, null, ex.Message);
        }
    }

    private static async Task<byte[]?> TryReadPdfResponseAsync(HttpResponseMessage response)
    {
        var mediaType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;
        if (!mediaType.Contains("pdf", StringComparison.OrdinalIgnoreCase))
            return null;

        var content = await response.Content.ReadAsByteArrayAsync();
        if (content.Length == 0)
            return null;

        using var stream = new MemoryStream(content, writable: false);
        return IsPdfStream(stream) ? content : null;
    }

    private static string? ExtractCrystalHtmlMessage(string? html)
    {
        if (string.IsNullOrWhiteSpace(html))
            return null;

        var match = Regex.Match(html,
            @"<td[^>]*colspan=""3""[^>]*>(.*?)</td>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        var candidate = match.Success ? match.Groups[1].Value : html;
        candidate = Regex.Replace(candidate, "<[^>]+>", " ");
        candidate = System.Net.WebUtility.HtmlDecode(candidate);
        candidate = Regex.Replace(candidate, @"\s+", " ").Trim();

        if (candidate.Length > 500)
            candidate = candidate[..500].Trim();

        return string.IsNullOrWhiteSpace(candidate) ? null : candidate;
    }

    private string? BuildCrystalRenderApiUrl()
    {
        if (string.IsNullOrWhiteSpace(_printSettings.CrystalBaseUrl))
            return null;

        return $"{_printSettings.CrystalBaseUrl.TrimEnd('/')}/api/report/render";
    }

    private string StoreGeneratedPrint(GeneratedPrintDocument document, string? requestedDocumentId = null)
    {
        var documentId = IsValidGeneratedPrintDocumentId(requestedDocumentId)
            ? requestedDocumentId!
            : Guid.NewGuid().ToString("N");
        _memoryCache.Set(
            BuildGeneratedPrintCacheKey(documentId),
            document,
            new MemoryCacheEntryOptions
            {
                AbsoluteExpirationRelativeToNow = GeneratedPrintLifetime,
                SlidingExpiration = TimeSpan.FromMinutes(5)
            });

        return documentId;
    }

    private static string BuildGeneratedPrintCacheKey(string documentId)
    {
        return $"generated-print:{documentId}";
    }

    private static GeneratedPrintDocument BuildGeneratedPrintDocument(string fileStem, IReadOnlyCollection<GeneratedPrintFile> files)
    {
        if (files.Count == 0)
            throw new InvalidOperationException("Ingen filer fundet til generering.");

        return new GeneratedPrintDocument
        {
            FileName = $"{fileStem}.pdf",
            Content = MergePdfFiles(files),
            ContentType = "application/pdf",
            OpenInline = true
        };
    }

    private static byte[] MergePdfFiles(IReadOnlyCollection<GeneratedPrintFile> files)
    {
        using var outputDocument = new PdfDocument();
        var validFileCount = 0;

        foreach (var file in files)
        {
            if (file.Content == null || file.Content.Length == 0)
                continue;

            using var inputStream = new MemoryStream(file.Content, writable: false);
            if (!IsPdfStream(inputStream))
                continue;

            using var importStream = new MemoryStream(file.Content, writable: false);
            using var inputDocument = PdfReader.Open(importStream, PdfDocumentOpenMode.Import);
            validFileCount++;

            for (var i = 0; i < inputDocument.PageCount; i++)
            {
                outputDocument.AddPage(inputDocument.Pages[i]);
            }
        }

        if (validFileCount == 0 || outputDocument.PageCount == 0)
            throw new InvalidOperationException("Ingen PDF-sider fundet til merge.");

        using var outputStream = new MemoryStream();
        outputDocument.Save(outputStream, false);
        return outputStream.ToArray();
    }

    private static string CreateOrderPdfFileName(int documentNumber, string itemCode)
    {
        return CreateSafeFileName($"produktionsordre_{documentNumber}_{itemCode}.pdf");
    }

    private static string CreateAttachmentFileName(int documentNumber, string itemCode, string attachmentPath, int attachmentIndex)
    {
        var attachmentName = Path.GetFileName(attachmentPath);
        if (string.IsNullOrWhiteSpace(attachmentName))
            attachmentName = $"attachment_{attachmentIndex}.pdf";

        return CreateSafeFileName($"produktionsordre_{documentNumber}_{itemCode}_{attachmentName}");
    }

    private static string CreateSafeFileName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "document";

        var invalidChars = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(value.Length);

        foreach (var ch in value)
        {
            builder.Append(invalidChars.Contains(ch) ? '_' : ch);
        }

        return builder.ToString();
    }

    private static bool IsPdfFile(string filePath)
    {
        try
        {
            using var stream = System.IO.File.OpenRead(filePath);
            return IsPdfStream(stream);
        }
        catch
        {
            return false;
        }
    }

    private static bool IsPdfStream(Stream stream)
    {
        if (stream == null || !stream.CanRead)
            return false;

        var header = new byte[5];
        var bytesRead = stream.Read(header, 0, header.Length);
        if (stream.CanSeek)
            stream.Position = 0;

        return bytesRead >= 5
            && header[0] == (byte)'%'
            && header[1] == (byte)'P'
            && header[2] == (byte)'D'
            && header[3] == (byte)'F'
            && header[4] == (byte)'-';
    }

    private static string EscapeODataString(string value)
    {
        return (value ?? string.Empty).Replace("'", "''");
    }

    private static bool ProductionOrderMatchesSalesOrder(
        JObject order,
        int salesOrderDocEntry,
        int salesOrderDocNum,
        ISet<int> printableLineNums,
        IReadOnlyDictionary<int, JObject> salesOrderLineInfo)
    {
        var originEntry = GetInt(order, "ProductionOrderOriginEntry");
        var originNumber = GetInt(order, "ProductionOrderOriginNumber");
        var orderLine = GetInt(order, "U_RCS_OL");

        if (IsPrintableProductionOrder(originEntry, originNumber, orderLine, salesOrderDocEntry, salesOrderDocNum, printableLineNums))
            return true;

        if (orderLine < 0 || !salesOrderLineInfo.TryGetValue(orderLine, out var salesLine))
            return false;

        var productionItemCode = GetString(order, "ItemNo") ?? string.Empty;
        var salesItemCode = GetString(salesLine, "ItemCode") ?? string.Empty;
        return !string.IsNullOrWhiteSpace(productionItemCode)
               && productionItemCode.Equals(salesItemCode, StringComparison.OrdinalIgnoreCase)
               && printableLineNums.Contains(orderLine);
    }

    private static bool IsPrintableProductionOrder(
        int originEntry,
        int originNumber,
        int orderLine,
        int salesOrderDocEntry,
        int salesOrderDocNum,
        ISet<int> printableLineNums)
    {
        var originMatches = originEntry > 0
            ? originEntry == salesOrderDocEntry
            : originNumber > 0 && salesOrderDocNum > 0 && originNumber == salesOrderDocNum;

        if (!originMatches)
            return false;

        return orderLine < 0 || printableLineNums.Contains(orderLine);
    }

    private static string? GetString(JToken? token, string name)
    {
        var value = token?[name];
        if (value == null || value.Type == JTokenType.Null)
            return null;

        var str = value.ToString();
        return string.IsNullOrWhiteSpace(str) ? null : str.Trim();
    }

    private static int GetInt(JToken? token, string name)
    {
        var value = token?[name];
        if (value == null || value.Type == JTokenType.Null)
            return 0;

        if (value.Type == JTokenType.Integer)
            return value.Value<int>();

        if (int.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
            return parsed;

        return 0;
    }
}
