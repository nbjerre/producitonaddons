using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using Newtonsoft.Json.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Net.Http;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Text;
using System.Text.RegularExpressions;
using WorksheetAPI.Configuration;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;
using WebAPI.Configuration;

namespace WebAPI.Controllers.salescontrollers;

/// <summary>
/// Endpoints til oprettelse, statusændring og print af produktionsordrer baseret på salgsordrer.
/// </summary>
[ApiController]
[Route("api/sales-production")]
public class SalesProductionController : ControllerBase
{
    private readonly ISapConnectionService _sapConnection;
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<SalesProductionController> _logger;
    private readonly PrintSettings _printSettings;
    private readonly IMemoryCache _memoryCache;
    private readonly IHttpClientFactory _httpClientFactory;
    private static readonly TimeSpan GeneratedPrintLifetime = TimeSpan.FromMinutes(15);

    public SalesProductionController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        ILogger<SalesProductionController> logger,
        IMemoryCache memoryCache,
        IOptions<PrintSettings> printSettingsOptions,
        IHttpClientFactory httpClientFactory)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _logger = logger;
        _memoryCache = memoryCache;
        _printSettings = printSettingsOptions.Value ?? new PrintSettings();
        _httpClientFactory = httpClientFactory;
    }

    private sealed class GeneratedPrintDocument
    {
        public required string FileName { get; init; }
        public required byte[] Content { get; init; }
        public required string ContentType { get; init; }
        public bool OpenInline { get; init; }
    }

    private sealed class GeneratedPrintFile
    {
        public required string FileName { get; init; }
        public required byte[] Content { get; init; }
        public string ContentType { get; init; } = "application/pdf";
    }

    /// <summary>
    /// Opretter en produktionsordre for en bestemt salgsordrelinje.
    /// </summary>
    [HttpPost("create-for-line")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> CreateProductionForLine([FromBody] CreateProductionForLineRequest request)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (request.SalesOrderDocEntry <= 0 && (!request.SalesOrderDocNum.HasValue || request.SalesOrderDocNum <= 0))
            return BadRequest(new { success = false, error = "SalesOrderDocEntry or SalesOrderDocNum is required" });

        var company = _sapConnection.GetMainCompany();

        var salesOrderDocEntry = request.SalesOrderDocEntry;
        if (salesOrderDocEntry <= 0)
        {
            salesOrderDocEntry = await ResolveDocEntryByDocNum(company, request.SalesOrderDocNum ?? 0);
            if (salesOrderDocEntry <= 0)
                return BadRequest(new { success = false, error = "Could not resolve sales order from DocNum" });
        }

        var orderResult = await _serviceLayer.GetStringAsync(company, $"Orders({salesOrderDocEntry})");
        if (!orderResult.Success || string.IsNullOrEmpty(orderResult.Data))
        {
            return BadRequest(new { success = false, error = orderResult.Error ?? "Could not read sales order" });
        }

        var orderObj = JToken.Parse(orderResult.Data) as JObject;
        if (orderObj == null)
            return BadRequest(new { success = false, error = "Invalid sales order payload from SAP" });

        var line = ResolveOrderLine(orderObj, request.LineNum, request.ItemCode);
        if (line == null)
        {
            return BadRequest(new { success = false, error = "No matching sales order line found" });
        }

        var itemCode = GetString(line, "ItemCode") ?? request.ItemCode ?? string.Empty;
        if (string.IsNullOrWhiteSpace(itemCode))
            return BadRequest(new { success = false, error = "ItemCode is required" });

        var hasBom = await HasProductionBom(company, itemCode);
        if (!hasBom)
        {
            return Ok(new
            {
                success = true,
                created = false,
                reason = "NO_BOM",
                message = "Den valgte linje har ingen produktionsstykliste (TreeType = P)."
            });
        }

        var lineNum = GetInt(line, "LineNum");
        var visOrder = GetInt(line, "VisOrder");
        var quantity = ResolveQuantity(line);
        var dueDate = ResolveDueDate(orderObj, line);
        var lineProject = GetString(line, "ProjectCode")
                          ?? GetString(line, "Project")
                          ?? GetString(orderObj, "Project")
                          ?? string.Empty;
        var customerCode = GetString(orderObj, "CardCode") ?? string.Empty;
        var removeDelivery = (GetString(orderObj, "U_RCS_RD") ?? "N").Equals("Y", StringComparison.OrdinalIgnoreCase);
        var uRcsOnSto = GetString(line, "U_RCS_ONSTO") ?? "Y";
        var delDays = ResolveDelDays(orderObj, line);

        if (quantity <= 0)
            return BadRequest(new { success = false, error = "Quantity is 0. Der kan ikke oprettes produktionsordre." });

        if (await ExistsOpenProductionOrder(company, itemCode, salesOrderDocEntry, visOrder, lineNum))
        {
            return Ok(new
            {
                success = true,
                created = false,
                reason = "ALREADY_EXISTS",
                message = "Der findes allerede en åben produktionsordre for denne linje."
            });
        }

        var subBomCodes = await GetSubBomCodes(company, itemCode);
        if (subBomCodes.Count > 0 && !request.ConfirmSubBoms)
        {
            var defaults = subBomCodes.Select(code => new { itemCode = code, u_RCS_PQT = 0m, u_RCS_ONSTO = "Y" }).ToList();
            return Ok(new
            {
                success = true,
                created = false,
                requiresSubBomConfirmation = true,
                message = "Understyklister fundet. Bekræft oprettelse.",
                subBoms = defaults
            });
        }

        var originalValues = new List<TempBomSnapshot>();
        try
        {
            if (subBomCodes.Count > 0)
            {
                await ApplyTemporarySubBomUpdates(company, subBomCodes, request.SubBomAdjustments, originalValues);
            }

            var payload = new JObject
            {
                ["ItemNo"] = itemCode,
                ["CustomerCode"] = customerCode,
                ["DueDate"] = dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                ["ProductionOrderOrigin"] = "bopooSalesOrder",
                ["ProductionOrderOriginEntry"] = salesOrderDocEntry,
                ["ProductionOrderType"] = "bopotStandard",
                ["PlannedQuantity"] = quantity,
                ["Project"] = lineProject,
                ["U_RCS_BVO"] = visOrder,
                ["U_RCS_OL"] = lineNum,
                ["U_RCS_DelDays"] = delDays
            };

            var createResult = await _serviceLayer.PostAsync<JObject>(company, "ProductionOrders", payload);
            if (!createResult.Success)
            {
                return BadRequest(new { success = false, error = createResult.Error ?? "Failed to create production order" });
            }

            var createdDocEntry = GetInt(createResult.Data, "AbsoluteEntry");

            if (createdDocEntry > 0)
            {
                if (removeDelivery)
                {
                    await RemoveDeliveryFromProductionOrder(company, createdDocEntry);
                }

                if (!uRcsOnSto.Equals("N", StringComparison.OrdinalIgnoreCase))
                {
                    await CreateSubProductionOrdersForSubBom(
                        company,
                        parentItemCode: itemCode,
                        parentQuantity: quantity,
                        project: lineProject,
                        shipDate: dueDate,
                        cardCode: customerCode,
                        orderDocEntry: salesOrderDocEntry,
                        orderLine: lineNum,
                        visOrder: visOrder,
                        productionBaseEntry: createdDocEntry,
                        productionBaseLine: 0,
                        removeDelivery: removeDelivery,
                        rcsDelDays: delDays,
                        depth: 0);
                }
            }

            return Ok(new
            {
                success = true,
                created = true,
                productionDocEntry = createdDocEntry,
                removeDelivery,
                u_RCS_ONSTO = uRcsOnSto,
                message = "Produktionsordre er oprettet"
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CreateProductionForLine failed for docEntry {DocEntry}, line {LineNum}, item {ItemCode}",
                salesOrderDocEntry, lineNum, itemCode);
            return StatusCode(500, new { success = false, error = ex.Message });
        }
        finally
        {
            if (originalValues.Count > 0)
            {
                await RestoreTemporarySubBomUpdates(company, originalValues);
            }
        }
    }
    /// <summary>
    /// Opretter produktionsordrer for alle gyldige linjer på en salgsordre.
    /// </summary>
    [HttpPost("create-all")]
    [Consumes("application/json", "text/plain", "application/x-www-form-urlencoded")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> CreateAllProductions([FromBody] CreateAllProductionsRequest request)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (request.SalesOrderDocEntry <= 0 && (!request.SalesOrderDocNum.HasValue || request.SalesOrderDocNum <= 0))
            return BadRequest(new { success = false, error = "SalesOrderDocEntry or SalesOrderDocNum is required" });

        var company = _sapConnection.GetMainCompany();

        var salesOrderDocEntry = request.SalesOrderDocEntry;
        if (salesOrderDocEntry <= 0)
        {
            salesOrderDocEntry = await ResolveDocEntryByDocNum(company, request.SalesOrderDocNum ?? 0);
            if (salesOrderDocEntry <= 0)
                return BadRequest(new { success = false, error = "Could not resolve sales order from DocNum" });
        }

        var orderResult = await _serviceLayer.GetStringAsync(company, $"Orders({salesOrderDocEntry})");
        if (!orderResult.Success || string.IsNullOrEmpty(orderResult.Data))
            return BadRequest(new { success = false, error = orderResult.Error ?? "Could not read sales order" });

        var orderObj = JToken.Parse(orderResult.Data) as JObject;
        if (orderObj == null)
            return BadRequest(new { success = false, error = "Invalid sales order payload from SAP" });

        var lines = orderObj["DocumentLines"] as JArray;
        if (lines == null || lines.Count == 0)
        {
            return Ok(new
            {
                success = true,
                created = false,
                message = "Salgsordren har ingen linjer."
            });
        }

        var removeDelivery = (GetString(orderObj, "U_RCS_RD") ?? "N").Equals("Y", StringComparison.OrdinalIgnoreCase);
        var customerCode = GetString(orderObj, "CardCode") ?? string.Empty;
        var allSubBomCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var creatableLines = new List<CreateAllLineContext>();
        var skipped = new List<object>();

        foreach (var line in lines.OfType<JObject>())
        {
            var itemCode = GetString(line, "ItemCode") ?? string.Empty;
            var lineNum = GetInt(line, "LineNum");
            var visOrder = GetInt(line, "VisOrder");

            if (string.IsNullOrWhiteSpace(itemCode))
            {
                skipped.Add(new { lineNum, itemCode, reason = "EMPTY_ITEM" });
                continue;
            }

            var hasBom = await HasProductionBom(company, itemCode);
            if (!hasBom)
            {
                skipped.Add(new { lineNum, itemCode, reason = "NO_BOM" });
                continue;
            }

            var quantity = ResolveQuantity(line);
            if (quantity <= 0)
            {
                skipped.Add(new { lineNum, itemCode, reason = "QTY_0" });
                continue;
            }

            var uRcsOnSto = (GetString(line, "U_RCS_ONSTO") ?? "Y").ToUpperInvariant();
            if (uRcsOnSto == "N")
            {
                skipped.Add(new { lineNum, itemCode, reason = "ONSTO_N" });
                continue;
            }

            if (await ExistsOpenProductionOrder(company, itemCode, salesOrderDocEntry, visOrder, lineNum))
            {
                skipped.Add(new { lineNum, itemCode, reason = "ALREADY_EXISTS" });
                continue;
            }

            var subBomCodes = await GetSubBomCodes(company, itemCode);
            foreach (var code in subBomCodes)
                allSubBomCodes.Add(code);

            creatableLines.Add(new CreateAllLineContext
            {
                ItemCode = itemCode,
                LineNum = lineNum,
                VisOrder = visOrder,
                Quantity = quantity,
                DueDate = ResolveDueDate(orderObj, line),
                Project = GetString(line, "ProjectCode")
                          ?? GetString(line, "Project")
                          ?? GetString(orderObj, "Project")
                          ?? string.Empty,
                DelDays = ResolveDelDays(orderObj, line)
            });
        }

        if (creatableLines.Count == 0)
        {
            return Ok(new
            {
                success = true,
                created = false,
                message = "Ingen linjer opfylder betingelserne for oprettelse.",
                skipped
            });
        }

        if (allSubBomCodes.Count > 0 && !request.ConfirmSubBoms)
        {
            var defaults = allSubBomCodes
                .OrderBy(v => v)
                .Select(code => new { itemCode = code, u_RCS_PQT = 0m, u_RCS_ONSTO = "Y" })
                .ToList();

            return Ok(new
            {
                success = true,
                created = false,
                requiresSubBomConfirmation = true,
                message = "Understyklister fundet. Bekræft oprettelse.",
                subBoms = defaults,
                candidateLines = creatableLines.Count,
                skipped
            });
        }

        var snapshots = new List<TempBomSnapshot>();
        var created = new List<object>();
        var failed = new List<object>();

        try
        {
            if (allSubBomCodes.Count > 0)
            {
                await ApplyTemporarySubBomUpdates(company, allSubBomCodes.ToList(), request.SubBomAdjustments, snapshots);
            }

            foreach (var line in creatableLines)
            {
                var payload = new JObject
                {
                    ["ItemNo"] = line.ItemCode,
                    ["CustomerCode"] = customerCode,
                    ["DueDate"] = line.DueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                    ["ProductionOrderOrigin"] = "bopooSalesOrder",
                    ["ProductionOrderOriginEntry"] = salesOrderDocEntry,
                    ["ProductionOrderType"] = "bopotStandard",
                    ["PlannedQuantity"] = line.Quantity,
                    ["Project"] = line.Project,
                    ["U_RCS_BVO"] = line.VisOrder,
                    ["U_RCS_OL"] = line.LineNum,
                    ["U_RCS_DelDays"] = line.DelDays
                };

                var createResult = await _serviceLayer.PostAsync<JObject>(
     company,
     "ProductionOrders",
     payload);

                if (!createResult.Success)
                {
                    failed.Add(new
                    {
                        lineNum = line.LineNum,
                        itemCode = line.ItemCode,
                        error = createResult.Error ?? "Failed to create production order"
                    });
                    continue;
                }

                var createdDocEntry = GetInt(createResult.Data, "AbsoluteEntry");
                if (createdDocEntry > 0)
                {
                    if (removeDelivery)
                    {
                        await RemoveDeliveryFromProductionOrder(company, createdDocEntry);
                    }

                    await CreateSubProductionOrdersForSubBom(
                        company,
                        parentItemCode: line.ItemCode,
                        parentQuantity: line.Quantity,
                        project: line.Project,
                        shipDate: line.DueDate,
                        cardCode: customerCode,
                        orderDocEntry: salesOrderDocEntry,
                        orderLine: line.LineNum,
                        visOrder: line.VisOrder,
                        productionBaseEntry: createdDocEntry,
                        productionBaseLine: 0,
                        removeDelivery: removeDelivery,
                        rcsDelDays: line.DelDays,
                        depth: 0);
                }

                created.Add(new
                {
                    lineNum = line.LineNum,
                    itemCode = line.ItemCode,
                    productionDocEntry = createdDocEntry
                });
            }

            return Ok(new
            {
                success = true,
                created = created.Count > 0,
                createdCount = created.Count,
                failedCount = failed.Count,
                skippedCount = skipped.Count,
                createdLines = created,
                failedLines = failed,
                skipped,
                message = $"Oprettet {created.Count} produktionsordrer."
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CreateAllProductions failed for docEntry {DocEntry}", salesOrderDocEntry);
            return StatusCode(500, new { success = false, error = ex.Message });
        }
        finally
        {
            if (snapshots.Count > 0)
            {
                await RestoreTemporarySubBomUpdates(company, snapshots);
            }
        }
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

    private async Task<int> ResolveCreatedProductionDocEntry(
        SapCompany company,
        string? createPayload,
        string itemCode,
        int salesOrderDocEntry,
        int visOrder,
        int lineNum)
    {
        if (!string.IsNullOrWhiteSpace(createPayload))
        {
            try
            {
                var obj = JToken.Parse(createPayload) as JObject;
                var docEntry = GetInt(obj, "AbsoluteEntry");
                if (docEntry <= 0)
                    docEntry = GetInt(obj, "DocEntry");
                if (docEntry > 0)
                    return docEntry;
            }
            catch
            {
            }
        }

        var filter =
            $"ItemNo eq '{Escape(itemCode)}' and ProductionOrderOriginEntry eq {salesOrderDocEntry} and U_RCS_BVO eq {visOrder} and U_RCS_OL eq {lineNum}";
        var result = await _serviceLayer.GetStringAsync(company, "ProductionOrders", select: "AbsoluteEntry,DocEntry", filter: filter, top: 20);
        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return 0;

        var root = JToken.Parse(result.Data) as JObject;
        var values = root?["value"] as JArray;
        if (values == null || values.Count == 0)
            return 0;

        var best = values.OfType<JObject>()
            .Select(v =>
            {
                var docEntry = GetInt(v, "AbsoluteEntry");
                if (docEntry <= 0)
                    docEntry = GetInt(v, "DocEntry");
                return docEntry;
            })
            .OrderByDescending(v => v)
            .FirstOrDefault();

        return best;
    }

    private async Task RemoveDeliveryFromProductionOrder(SapCompany company, int docEntry)
    {
        try
        {
            var getResult = await _serviceLayer.GetStringAsync(company, $"ProductionOrders({docEntry})");
            if (!getResult.Success || string.IsNullOrEmpty(getResult.Data))
                return;

            var order = JToken.Parse(getResult.Data) as JObject;
            if (order == null)
                return;

            var lines = order["ProductionOrderLines"] as JArray;
            if (lines == null || lines.Count == 0)
                return;

            var removeStageIds = new HashSet<int>();
            var filteredLines = new JArray();
            foreach (var lineToken in lines.OfType<JObject>())
            {
                var item = GetString(lineToken, "ItemNo") ?? string.Empty;
                if (item == "999")
                {
                    var stage = GetInt(lineToken, "StageID");
                    if (stage > 0)
                        removeStageIds.Add(stage);
                    continue;
                }
                filteredLines.Add(lineToken);
            }

            if (filteredLines.Count == lines.Count)
                return;

            var patch = new JObject
            {
                ["ProductionOrderLines"] = filteredLines
            };

            var stages = order["ProductionOrderStages"] as JArray;
            if (stages != null && removeStageIds.Count > 0)
            {
                var filteredStages = new JArray();
                foreach (var stageToken in stages.OfType<JObject>())
                {
                    var stageId = GetInt(stageToken, "StageID");
                    if (!removeStageIds.Contains(stageId))
                        filteredStages.Add(stageToken);
                }
                patch["ProductionOrderStages"] = filteredStages;
            }

            var patchResult = await _serviceLayer.PatchAsync(company, $"ProductionOrders({docEntry})", patch);
            if (!patchResult.Success)
            {
                _logger.LogWarning("RemoveDeliveryFromProductionOrder failed for docEntry {DocEntry}: {Error}", docEntry, patchResult.Error);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "RemoveDeliveryFromProductionOrder exception for docEntry {DocEntry}", docEntry);
        }
    }

    private async Task CreateSubProductionOrdersForSubBom(
        SapCompany company,
        string parentItemCode,
        decimal parentQuantity,
        string project,
        DateTime shipDate,
        string cardCode,
        int orderDocEntry,
        int orderLine,
        int visOrder,
        int productionBaseEntry,
        int productionBaseLine,
        bool removeDelivery,
        int rcsDelDays,
        int depth)
    {
        if (depth > 10)
            return;

        var treeResult = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(parentItemCode)}')");
        if (!treeResult.Success || string.IsNullOrEmpty(treeResult.Data))
            return;

        var tree = JToken.Parse(treeResult.Data) as JObject;
        var lines = ResolveProductTreeLines(tree);
        if (lines == null || lines.Count == 0)
            return;

        foreach (var line in lines.OfType<JObject>())
        {
            var childItem = GetString(line, "ItemCode") ?? GetString(line, "Code") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(childItem))
                continue;

            if (!await HasProductionBom(company, childItem))
                continue;

            var childTreeResult = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(childItem)}')");
            if (!childTreeResult.Success || string.IsNullOrEmpty(childTreeResult.Data))
                continue;

            var childTree = JToken.Parse(childTreeResult.Data) as JObject;
            var uRcsOnSto = (GetString(childTree, "U_RCS_ONSTO") ?? "Y").ToUpperInvariant();
            if (uRcsOnSto == "N")
                continue;

            var lineQty = GetDecimal(line, "Quantity");
            if (lineQty <= 0)
                lineQty = 1m;

            var forcedQty = GetDecimal(childTree, "U_RCS_PQT");
            var createQty = forcedQty > 0 ? forcedQty : Math.Max(parentQuantity * lineQty, 1m);

            if (await ExistsOpenSubProductionOrder(company, childItem, orderDocEntry, visOrder, orderLine, productionBaseEntry))
                continue;

            var payload = new JObject
            {
                ["ItemNo"] = childItem,
                ["CustomerCode"] = cardCode,
                ["DueDate"] = shipDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                ["ProductionOrderOrigin"] = "bopooSalesOrder",
                ["ProductionOrderOriginEntry"] = orderDocEntry,
                ["ProductionOrderType"] = "bopotStandard",
                ["PlannedQuantity"] = createQty,
                ["Project"] = project,
                ["U_RCS_BVO"] = visOrder,
                ["U_RCS_OL"] = orderLine,
                ["U_RCS_PB"] = productionBaseEntry,
                ["U_RCS_PBL"] = productionBaseLine,
                ["U_RCS_DelDays"] = rcsDelDays
            };

            var create = await _serviceLayer.PostAsync<string>(company, "ProductionOrders", payload);
            if (!create.Success)
            {
                _logger.LogWarning("Create sub production failed for {ItemCode}: {Error}", childItem, create.Error);
                continue;
            }

            var childProdEntry = await ResolveCreatedSubProductionDocEntry(
                company,
                create.Data,
                childItem,
                orderDocEntry,
                visOrder,
                orderLine,
                productionBaseEntry);

            if (childProdEntry <= 0)
                continue;

            if (removeDelivery)
            {
                await RemoveDeliveryFromProductionOrder(company, childProdEntry);
            }

            await CreateSubProductionOrdersForSubBom(
                company,
                parentItemCode: childItem,
                parentQuantity: createQty,
                project: project,
                shipDate: shipDate,
                cardCode: cardCode,
                orderDocEntry: orderDocEntry,
                orderLine: orderLine,
                visOrder: visOrder,
                productionBaseEntry: childProdEntry,
                productionBaseLine: GetInt(line, "LineNum"),
                removeDelivery: removeDelivery,
                rcsDelDays: rcsDelDays,
                depth: depth + 1);
        }
    }
    /// <summary>
    /// Annullerer alle planlagte eller frigivne produktionsordrer for en salgsordre.
    /// </summary>
    [HttpPost("cancel-all")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> CancelAllProductions([FromBody] CancelAllProductionsRequest request)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (request.SalesOrderDocEntry <= 0 && (!request.SalesOrderDocNum.HasValue || request.SalesOrderDocNum <= 0))
            return BadRequest(new { success = false, error = "SalesOrderDocEntry or SalesOrderDocNum is required" });

        var company = _sapConnection.GetMainCompany();

        var salesOrderDocEntry = request.SalesOrderDocEntry;
        if (salesOrderDocEntry <= 0)
        {
            salesOrderDocEntry = await ResolveDocEntryByDocNum(company, request.SalesOrderDocNum ?? 0);
            if (salesOrderDocEntry <= 0)
                return BadRequest(new { success = false, error = "Could not resolve sales order from DocNum" });
        }

        var filter =
            $"ProductionOrderOriginEntry eq {salesOrderDocEntry} " +
            $"and (ProductionOrderStatus eq 'boposPlanned' or ProductionOrderStatus eq 'boposReleased')";

        var listResult = await _serviceLayer.GetStringAsync(company, "ProductionOrders",
    select: "AbsoluteEntry,DocumentNumber", filter: filter, top: 100);

        if (!listResult.Success || string.IsNullOrEmpty(listResult.Data))
            return BadRequest(new { success = false, error = listResult.Error ?? "Could not fetch production orders" });

        var root = JToken.Parse(listResult.Data) as JObject;
        var orders = root?["value"] as JArray;

        if (orders == null || orders.Count == 0)
            return Ok(new { success = true, cancelled = false, message = "Ingen åbne produktionsordrer fundet." });

        var cancelled = new List<object>();
        var failed = new List<object>();

        foreach (var order in orders.OfType<JObject>())
        {
            var docEntry = GetInt(order, "AbsoluteEntry");
            var docNum = GetInt(order, "DocNumber");

            var patch = new JObject
            {
                ["ProductionOrderStatus"] = "boposCancelled"
            };

            var patchResult = await _serviceLayer.PatchAsync(company, $"ProductionOrders({docEntry})", patch);
            if (!patchResult.Success)
            {
                failed.Add(new { docEntry, docNum, error = patchResult.Error });
            }
            else
            {
                cancelled.Add(new { docEntry, docNum });
            }
        }

        return Ok(new
        {
            success = true,
            cancelled = cancelled.Count > 0,
            cancelledCount = cancelled.Count,
            failedCount = failed.Count,
            cancelledOrders = cancelled,
            failedOrders = failed,
            message = $"Annulleret {cancelled.Count} produktionsordrer."
        });
    }
    /// <summary>
    /// Frigiver alle planlagte produktionsordrer for en salgsordre.
    /// </summary>
    [HttpPost("release-all")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> ReleaseAllProductions([FromBody] CancelAllProductionsRequest request)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (request.SalesOrderDocEntry <= 0 && (!request.SalesOrderDocNum.HasValue || request.SalesOrderDocNum <= 0))
            return BadRequest(new { success = false, error = "SalesOrderDocEntry or SalesOrderDocNum is required" });

        var company = _sapConnection.GetMainCompany();

        var salesOrderDocEntry = request.SalesOrderDocEntry;
        if (salesOrderDocEntry <= 0)
        {
            salesOrderDocEntry = await ResolveDocEntryByDocNum(company, request.SalesOrderDocNum ?? 0);
            if (salesOrderDocEntry <= 0)
                return BadRequest(new { success = false, error = "Could not resolve sales order from DocNum" });
        }

        var filter =
            $"ProductionOrderOriginEntry eq {salesOrderDocEntry} " +
            $"and ProductionOrderStatus eq 'boposPlanned'";

        var listResult = await _serviceLayer.GetStringAsync(company, "ProductionOrders",
            select: "AbsoluteEntry,DocumentNumber", filter: filter, top: 100);

        if (!listResult.Success || string.IsNullOrEmpty(listResult.Data))
            return BadRequest(new { success = false, error = listResult.Error ?? "Could not fetch production orders" });

        var root = JToken.Parse(listResult.Data) as JObject;
        var orders = root?["value"] as JArray;

        if (orders == null || orders.Count == 0)
            return Ok(new { success = true, released = false, message = "Ingen planlagte produktionsordrer fundet." });

        var released = new List<object>();
        var failed = new List<object>();

        foreach (var order in orders.OfType<JObject>())
        {
            var docEntry = GetInt(order, "AbsoluteEntry");
            var docNum = GetInt(order, "DocumentNumber");

            var patch = new JObject { ["ProductionOrderStatus"] = "boposReleased" };
            var patchResult = await _serviceLayer.PatchAsync(company, $"ProductionOrders({docEntry})", patch);

            if (!patchResult.Success)
                failed.Add(new { docEntry, docNum, error = patchResult.Error });
            else
                released.Add(new { docEntry, docNum });
        }

        return Ok(new
        {
            success = true,
            released = released.Count > 0,
            releasedCount = released.Count,
            failedCount = failed.Count,
            releasedOrders = released,
            failedOrders = failed,
            message = $"Frigivet {released.Count} produktionsordrer."
        });
    }

    /// <summary>
    /// Færdigmelder alle frigivne produktionsordrer for en salgsordre.
    /// </summary>
    [HttpPost("finish-all")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> FinishAllProductions([FromBody] CancelAllProductionsRequest request)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (request.SalesOrderDocEntry <= 0 && (!request.SalesOrderDocNum.HasValue || request.SalesOrderDocNum <= 0))
            return BadRequest(new { success = false, error = "SalesOrderDocEntry or SalesOrderDocNum is required" });

        var company = _sapConnection.GetMainCompany();

        var salesOrderDocEntry = request.SalesOrderDocEntry;
        if (salesOrderDocEntry <= 0)
        {
            salesOrderDocEntry = await ResolveDocEntryByDocNum(company, request.SalesOrderDocNum ?? 0);
            if (salesOrderDocEntry <= 0)
                return BadRequest(new { success = false, error = "Could not resolve sales order from DocNum" });
        }

        var filter =
            $"ProductionOrderOriginEntry eq {salesOrderDocEntry} " +
            $"and ProductionOrderStatus eq 'boposReleased'";

        var listResult = await _serviceLayer.GetStringAsync(company, "ProductionOrders",
            select: "AbsoluteEntry,DocumentNumber", filter: filter, top: 100);

        if (!listResult.Success || string.IsNullOrEmpty(listResult.Data))
            return BadRequest(new { success = false, error = listResult.Error ?? "Could not fetch production orders" });

        var root = JToken.Parse(listResult.Data) as JObject;
        var orders = root?["value"] as JArray;

        if (orders == null || orders.Count == 0)
            return Ok(new { success = true, finished = false, message = "Ingen frigivne produktionsordrer fundet." });

        var finished = new List<object>();
        var failed = new List<object>();

        foreach (var order in orders.OfType<JObject>())
        {
            var docEntry = GetInt(order, "AbsoluteEntry");
            var docNum = GetInt(order, "DocumentNumber");

            try
            {
                // Hent fuld produktionsordre
                var prodResult = await _serviceLayer.GetStringAsync(company, $"ProductionOrders({docEntry})");
                if (!prodResult.Success || string.IsNullOrEmpty(prodResult.Data))
                {
                    failed.Add(new { docEntry, docNum, error = "Could not fetch production order" });
                    continue;
                }

                var prodObj = JToken.Parse(prodResult.Data) as JObject;
                if (prodObj == null) continue;

                var plannedQty = GetDecimal(prodObj, "PlannedQuantity");
                var completedQty = GetDecimal(prodObj, "CompletedQuantity");
                var lines = prodObj["ProductionOrderLines"] as JArray;

                // Skift backflush linjer til manuel
                if (lines != null)
                {
                    var hasBackflush = lines.OfType<JObject>()
                        .Any(l => GetString(l, "IssueMethod") == "im_Backflush");

                    if (hasBackflush)
                    {
                        var updatedLines = new JArray();
                        foreach (var line in lines.OfType<JObject>())
                        {
                            var updatedLine = (JObject)line.DeepClone();
                            if (GetString(line, "IssueMethod") == "im_Backflush")
                                updatedLine["IssueMethod"] = "im_Manual";
                            updatedLines.Add(updatedLine);
                        }

                        var backflushPatch = new JObject { ["ProductionOrderLines"] = updatedLines };
                        await _serviceLayer.PatchAsync(company, $"ProductionOrders({docEntry})", backflushPatch);
                    }

                    // Opret lagerudgang for linjer hvor U_RCS_Issued > IssuedQuantity
                    foreach (var line in lines.OfType<JObject>())
                    {
                        var issued = GetDecimal(line, "U_RCS_Issued");
                        var alreadyIssued = GetDecimal(line, "IssuedQuantity");
                        var lineNum = GetInt(line, "LineNumber");

                        if (issued > alreadyIssued)
                        {
                            var exitPayload = new JObject
                            {
                                ["Lines"] = new JArray
                            {
                                new JObject
                                {
                                    ["BaseEntry"] = docEntry,
                                    ["BaseLine"] = lineNum,
                                    ["BaseType"] = 202,
                                    ["Quantity"] = issued - alreadyIssued
                                }
                            }
                            };

                            var exitResult = await _serviceLayer.PostAsync<JObject>(company, "InventoryGenExits", exitPayload);
                            if (!exitResult.Success)
                                _logger.LogWarning("InventoryGenExit failed for docEntry {DocEntry}, line {Line}: {Error}",
                                    docEntry, lineNum, exitResult.Error);
                        }
                    }
                }

                if (plannedQty > completedQty)
                {
                    var itemCode = GetString(prodObj, "ItemNo");
                    var warehouse = GetString(prodObj, "Warehouse") ?? "01";
                    var receiveQty = plannedQty - completedQty;

                    _logger.LogDebug("Finish produktionsordre - docEntry: {DocEntry}, itemCode: {ItemCode}, warehouse: {Warehouse}, qty: {Qty}",
                        docEntry, itemCode, warehouse, receiveQty);

                    var receiptPayload = new JObject
                    {
                        ["DocDate"] = DateTime.Today.ToString("yyyy-MM-dd"),
                        ["DocumentLines"] = new JArray
    {
        new JObject
        {
            ["BaseEntry"] = docEntry,
            ["BaseType"] = 202
        }
    }
                    };

                    _logger.LogDebug("Finish kvittering payload: {Payload}", receiptPayload.ToString());

                    var receiptResult = await _serviceLayer.PostAsync<JObject>(company, "InventoryGenEntries", receiptPayload);
                    if (!receiptResult.Success)
                    {
                        failed.Add(new { docEntry, docNum, error = receiptResult.Error });
                        continue;
                    }
                }

                finished.Add(new { docEntry, docNum });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "FinishAllProductions failed for docEntry {DocEntry}", docEntry);
                failed.Add(new { docEntry, docNum, error = ex.Message });
            }
        }

        return Ok(new
        {
            success = true,
            finished = finished.Count > 0,
            finishedCount = finished.Count,
            failedCount = failed.Count,
            finishedOrders = finished,
            failedOrders = failed,
            message = $"Færdigmeldt {finished.Count} produktionsordrer."
        });
    }

    /// <summary>
    /// Genererer et samlet printdokument for produktionsordrer relateret til en salgsordre.
    /// </summary>
    [HttpPost("print-production-orders")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> PrintProductionOrders([FromBody] CancelAllProductionsRequest request)
    {
        _logger.LogInformation(
            "PrintProductionOrders request received. Input SalesOrderDocEntry={SalesOrderDocEntry}, SalesOrderDocNum={SalesOrderDocNum}",
            request.SalesOrderDocEntry,
            request.SalesOrderDocNum);

        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (request.SalesOrderDocEntry <= 0 && (!request.SalesOrderDocNum.HasValue || request.SalesOrderDocNum <= 0))
            return BadRequest(new { success = false, error = "SalesOrderDocEntry or SalesOrderDocNum is required" });

        var company = _sapConnection.GetMainCompany();

        var salesOrderDocEntry = request.SalesOrderDocEntry;
        if (salesOrderDocEntry <= 0)
        {
            _logger.LogInformation(
                "PrintProductionOrders resolving sales order DocEntry from DocNum {SalesOrderDocNum}",
                request.SalesOrderDocNum);

            salesOrderDocEntry = await ResolveDocEntryByDocNum(company, request.SalesOrderDocNum ?? 0);
            if (salesOrderDocEntry <= 0)
                return BadRequest(new { success = false, error = "Could not resolve sales order from DocNum" });
        }

        _logger.LogInformation(
            "PrintProductionOrders loading sales order {SalesOrderDocEntry}",
            salesOrderDocEntry);

        var orderResult = await _serviceLayer.GetStringAsync(company, $"Orders({salesOrderDocEntry})");
        if (!orderResult.Success || string.IsNullOrEmpty(orderResult.Data))
            return BadRequest(new { success = false, error = orderResult.Error ?? "Could not read sales order" });

        var orderObj = JToken.Parse(orderResult.Data) as JObject;
        if (orderObj == null)
            return BadRequest(new { success = false, error = "Invalid sales order payload from SAP" });

        var salesOrderDocNum = GetInt(orderObj, "DocNum");
        var cardCode = GetString(orderObj, "CardCode") ?? string.Empty;
        var cardName = GetString(orderObj, "CardName") ?? string.Empty;

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
            return Ok(new
            {
                success = true,
                printableCount = 0,
                renderMode = "native",
                message = "Ingen produktionslinjer markeret til print (U_RCS_ONSTO = Y)."
            });
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
            return BadRequest(new { success = false, error = listResult.Error ?? "Could not fetch production orders" });

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

            return Ok(new
            {
                success = true,
                printableCount = 0,
                renderMode = "native",
                message = "Ingen åbne/frigivne produktionsordrer fundet."
            });
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
            return Ok(new
            {
                success = true,
                salesOrderDocEntry,
                printableCount = 0,
                renderMode = "native",
                message = "Ingen produktionsordrer opfylder print-kriterierne."
            });
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
            return Ok(new
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
            });
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
            return Ok(new
            {
                success = false,
                salesOrderDocEntry,
                printableCount = printable.Count,
                generatedCount = 0,
                failedCount = failedOrders.Count,
                failedOrders,
                renderMode = "crystal",
                message = $"Ingen PDF-filer kunne genereres via Crystal. Rapportkode: {productionReportCode}."
            });
        }

        var finalDocument = BuildGeneratedPrintDocument(fileStem, generatedFiles);
        var documentId = StoreGeneratedPrint(finalDocument, request.DocumentId);
        var downloadUrl = BuildGeneratedPrintDownloadUrl(documentId);
        var openUrl = BuildGeneratedPrintOpenUrl(documentId);

        return Ok(new
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
        });
    }

    /// <summary>
    /// GET-variant af printflowet for produktionsordrer.
    /// </summary>
    [HttpGet("print-production-orders")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public Task<IActionResult> PrintProductionOrdersGet([FromQuery] int? salesOrderDocEntry, [FromQuery] int? salesOrderDocNum, [FromQuery] string? documentId)
    {
        return PrintProductionOrders(new CancelAllProductionsRequest
        {
            SalesOrderDocEntry = salesOrderDocEntry ?? 0,
            SalesOrderDocNum = salesOrderDocNum,
            DocumentId = documentId
        });
    }

    /// <summary>
    /// Returnerer status for et tidligere genereret printdokument.
    /// </summary>
    [HttpGet("print-status/{documentId}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public IActionResult GetGeneratedPrintStatus(string documentId)
    {
        if (!IsValidGeneratedPrintDocumentId(documentId))
            return BadRequest(new { ready = false, error = "Ugyldigt documentId." });

        if (!TryGetGeneratedPrint(documentId, out var document) || document == null)
        {
            return Ok(new
            {
                ready = false
            });
        }

        return Ok(new
        {
            ready = true,
            fileName = document.FileName,
            openUrl = BuildGeneratedPrintOpenUrl(documentId),
            downloadUrl = BuildGeneratedPrintDownloadUrl(documentId),
            contentType = document.ContentType,
            openInline = document.OpenInline
        });
    }

    /// <summary>
    /// Returnerer en lille HTML-side der poller indtil printdokumentet er klar.
    /// </summary>
    [HttpGet("print-wait/{documentId}")]
    [Produces("text/html")]
    [ProducesResponseType(typeof(string), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(string), StatusCodes.Status400BadRequest)]
    public ContentResult WaitForGeneratedPrint(string documentId)
    {
        if (!IsValidGeneratedPrintDocumentId(documentId))
        {
            return Content("Ugyldigt documentId.", "text/plain; charset=utf-8", Encoding.UTF8);
        }

        var statusUrl = BuildGeneratedPrintStatusUrl(documentId) ?? string.Empty;
        var html = $$"""
<!DOCTYPE html>
<html lang="da">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Genererer PDF...</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 32px; color: #222; }
        .muted { color: #666; }
    </style>
</head>
<body>
    <h2>Genererer PDF...</h2>
    <p class="muted">Fanen opdateres automatisk når dokumentet er klar.</p>
    <script>
        const statusUrl = {{JsonSerializer.Serialize(statusUrl)}};
        async function poll() {
            try {
                const response = await fetch(statusUrl, { cache: 'no-store' });
                if (response.ok) {
                    const data = await response.json();
                    if (data && data.ready === true && data.openUrl) {
                        window.location.replace(data.openUrl);
                        return;
                    }
                }
            } catch (e) {
            }

            window.setTimeout(poll, 1000);
        }

        poll();
    </script>
</body>
</html>
""";

        return Content(html, "text/html; charset=utf-8", Encoding.UTF8);
    }

    /// <summary>
    /// Downloader et genereret printdokument.
    /// </summary>
    [HttpGet("print-download/{documentId}")]
    [Produces("application/pdf")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public IActionResult DownloadGeneratedPrint(string documentId)
    {
        if (!TryGetGeneratedPrint(documentId, out var document) || document == null)
            return NotFound();

        Response.Headers.CacheControl = "no-store, no-cache";
        return File(document.Content, document.ContentType, document.FileName);
    }

    /// <summary>
    /// Åbner et genereret printdokument inline hvis muligt.
    /// </summary>
    [HttpGet("print-open/{documentId}")]
    [Produces("application/pdf")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public IActionResult OpenGeneratedPrint(string documentId)
    {
        if (!TryGetGeneratedPrint(documentId, out var document) || document == null)
            return NotFound();

        Response.Headers.CacheControl = "no-store, no-cache";

        if (!document.OpenInline)
            return File(document.Content, document.ContentType, document.FileName);

        Response.Headers.ContentDisposition = $"inline; filename*=UTF-8''{Uri.EscapeDataString(document.FileName)}";
        return File(document.Content, document.ContentType);
    }

    private string? BuildGeneratedPrintDownloadUrl(string documentId)
    {
        if (string.IsNullOrWhiteSpace(documentId))
            return null;

        var baseApiUrl = $"{Request.Scheme}://{Request.Host}{Request.PathBase}".TrimEnd('/');
        return $"{baseApiUrl}/api/sales-production/print-download/{Uri.EscapeDataString(documentId)}";
    }

    private string? BuildGeneratedPrintOpenUrl(string documentId)
    {
        if (string.IsNullOrWhiteSpace(documentId))
            return null;

        var baseApiUrl = $"{Request.Scheme}://{Request.Host}{Request.PathBase}".TrimEnd('/');
        return $"{baseApiUrl}/api/sales-production/print-open/{Uri.EscapeDataString(documentId)}";
    }

    private string? BuildGeneratedPrintStatusUrl(string documentId)
    {
        if (string.IsNullOrWhiteSpace(documentId))
            return null;

        var baseApiUrl = $"{Request.Scheme}://{Request.Host}{Request.PathBase}".TrimEnd('/');
        return $"{baseApiUrl}/api/sales-production/print-status/{Uri.EscapeDataString(documentId)}";
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
        if (apiRenderResult.Success)
            return apiRenderResult;

        var crystalUrl = BuildCrystalProductionOrderUrl(company, productionDocEntry);
        if (string.IsNullOrWhiteSpace(crystalUrl))
            return apiRenderResult.Success
                ? apiRenderResult
                : (false, null, apiRenderResult.ErrorMessage ?? "Crystal-url kunne ikke bygges.");

        try
        {
            var client = _httpClientFactory.CreateClient("crystal");
            client.Timeout = TimeSpan.FromSeconds(20);

            _logger.LogInformation(
                "Crystal render GET started for productionDocEntry {ProductionDocEntry}. Url {Url}",
                productionDocEntry,
                crystalUrl);

            using var response = await client.GetAsync(crystalUrl, HttpCompletionOption.ResponseHeadersRead);
            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Crystal render failed for productionDocEntry {ProductionDocEntry}. Status {StatusCode}. Url {Url}",
                    productionDocEntry, (int)response.StatusCode, crystalUrl);
                return (false, null, $"Crystal-site svarede med HTTP {(int)response.StatusCode}.");
            }

            var pdfBytes = await TryReadPdfResponseAsync(response);
            if (pdfBytes != null)
            {
                _logger.LogInformation(
                    "Crystal render GET succeeded for productionDocEntry {ProductionDocEntry}. Url {Url}",
                    productionDocEntry,
                    crystalUrl);
                return (true, pdfBytes, null);
            }

            var responseBody = await response.Content.ReadAsStringAsync();
            var mediaType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;
            if (mediaType.Contains("html", StringComparison.OrdinalIgnoreCase) || LooksLikeHtml(responseBody))
            {
                var postbackResult = await TryRenderCrystalViaPostbackAsync(client, crystalUrl, responseBody);
                if (postbackResult.Success)
                    return postbackResult;

                _logger.LogWarning(
                    "Crystal postback render failed for productionDocEntry {ProductionDocEntry}. Url {Url}. Error {Error}",
                    productionDocEntry,
                    crystalUrl,
                    postbackResult.ErrorMessage);

                return postbackResult;
            }

            _logger.LogWarning("Crystal render returned non-PDF for productionDocEntry {ProductionDocEntry}. ContentType {ContentType}. Url {Url}",
                productionDocEntry, mediaType, crystalUrl);
            var detailedMessage = ExtractCrystalHtmlMessage(responseBody);
            return (false, null, string.IsNullOrWhiteSpace(detailedMessage)
                ? $"Crystal-site returnerede {mediaType} i stedet for PDF."
                : detailedMessage);
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogWarning(ex,
                "Crystal render timeout for productionDocEntry {ProductionDocEntry}. Url {Url}",
                productionDocEntry,
                crystalUrl);
            return (false, null, "Crystal-site timeout efter 20 sekunder.");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Crystal render exception for productionDocEntry {ProductionDocEntry}. Url {Url}",
                productionDocEntry, crystalUrl);
            return (false, null, ex.Message);
        }
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

    private async Task<(bool Success, byte[]? PdfBytes, string? ErrorMessage)> TryRenderCrystalViaPostbackAsync(
        HttpClient client,
        string crystalUrl,
        string html)
    {
        var viewState = ExtractHiddenInputValue(html, "__VIEWSTATE");
        if (string.IsNullOrWhiteSpace(viewState))
        {
            var message = ExtractCrystalHtmlMessage(html);
            return (false, null, string.IsNullOrWhiteSpace(message)
                ? "Crystal-site returnerede HTML uden ViewState, så export-knappen kunne ikke trigges."
                : message);
        }

        var postUrl = BuildCrystalPostbackUrl(crystalUrl, html);
        var form = new Dictionary<string, string>
        {
            ["__VIEWSTATE"] = viewState,
            ["ctl00$MainContent$RadioButtonList1"] = "AttachPDF",
            ["ctl00$MainContent$Button_Print1"] = "Vis rapport"
        };

        AddIfHasValue(form, "__VIEWSTATEGENERATOR", ExtractHiddenInputValue(html, "__VIEWSTATEGENERATOR"));
        AddIfHasValue(form, "__EVENTVALIDATION", ExtractHiddenInputValue(html, "__EVENTVALIDATION"));
        AddIfHasValue(form, "__EVENTTARGET", string.Empty);
        AddIfHasValue(form, "__EVENTARGUMENT", string.Empty);

        using var request = new HttpRequestMessage(HttpMethod.Post, postUrl)
        {
            Content = new FormUrlEncodedContent(form)
        };

        _logger.LogInformation("Crystal postback started. Url {Url}", postUrl);

        using var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
        if (!response.IsSuccessStatusCode)
            return (false, null, $"Crystal postback svarede med HTTP {(int)response.StatusCode}.");

        var pdfBytes = await TryReadPdfResponseAsync(response);
        if (pdfBytes != null)
        {
            _logger.LogInformation("Crystal postback succeeded. Url {Url}", postUrl);
            return (true, pdfBytes, null);
        }

        var responseBody = await response.Content.ReadAsStringAsync();
        var message2 = ExtractCrystalHtmlMessage(responseBody);
        var mediaType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;

        return (false, null, string.IsNullOrWhiteSpace(message2)
            ? $"Crystal postback returnerede {mediaType} i stedet for PDF."
            : message2);
    }

    private static void AddIfHasValue(IDictionary<string, string> form, string key, string? value)
    {
        if (value != null)
            form[key] = value;
    }

    private static string BuildCrystalPostbackUrl(string crystalUrl, string html)
    {
        var actionMatch = Regex.Match(
            html,
            @"<form[^>]*action=""(?<action>[^""]+)""",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        var action = actionMatch.Success ? System.Net.WebUtility.HtmlDecode(actionMatch.Groups["action"].Value) : string.Empty;
        if (string.IsNullOrWhiteSpace(action))
            return crystalUrl;

        if (Uri.TryCreate(action, UriKind.Absolute, out var absoluteUri))
            return absoluteUri.ToString();

        return new Uri(new Uri(crystalUrl), action).ToString();
    }

    private static string? ExtractHiddenInputValue(string html, string inputId)
    {
        if (string.IsNullOrWhiteSpace(html) || string.IsNullOrWhiteSpace(inputId))
            return null;

        var pattern = $@"id=""{Regex.Escape(inputId)}""\s+value=""(?<value>[^""]*)""|name=""{Regex.Escape(inputId)}""\s+id=""{Regex.Escape(inputId)}""\s+value=""(?<value2>[^""]*)""";
        var match = Regex.Match(html, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
        if (!match.Success)
            return null;

        var value = match.Groups["value"].Success
            ? match.Groups["value"].Value
            : match.Groups["value2"].Value;

        return System.Net.WebUtility.HtmlDecode(value);
    }

    private static bool LooksLikeHtml(string? body)
    {
        if (string.IsNullOrWhiteSpace(body))
            return false;

        return body.Contains("<html", StringComparison.OrdinalIgnoreCase)
            || body.Contains("<!DOCTYPE html", StringComparison.OrdinalIgnoreCase)
            || body.Contains("<form", StringComparison.OrdinalIgnoreCase);
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

    private string? BuildCrystalProductionOrderUrl(SapCompany company, int productionDocEntry)
    {
        if (string.IsNullOrWhiteSpace(_printSettings.CrystalBaseUrl))
            return null;

        var baseUrl = _printSettings.CrystalBaseUrl.TrimEnd('/');
        var reportCode = string.IsNullOrWhiteSpace(_printSettings.ProductionReportCode)
            ? "WOR10003"
            : _printSettings.ProductionReportCode.Trim();

        var query = new List<string>
        {
            $"database={Uri.EscapeDataString(company.CompanyDb)}",
            $"code={Uri.EscapeDataString(reportCode)}",
            "type=AttachPDF",
            $"cr_input_value_DocKey={Uri.EscapeDataString(productionDocEntry.ToString(CultureInfo.InvariantCulture))}",

              // Legacy compatibility for endpoints that may still read the raw parameter name.
              $"DocKey={Uri.EscapeDataString(productionDocEntry.ToString(CultureInfo.InvariantCulture))}"
        };

        return $"{baseUrl}/SelectParams.aspx?{string.Join("&", query)}";
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

    private bool TryGetGeneratedPrint(string documentId, out GeneratedPrintDocument? document)
    {
        document = null;
        if (!IsValidGeneratedPrintDocumentId(documentId))
            return false;

        return _memoryCache.TryGetValue(BuildGeneratedPrintCacheKey(documentId), out document);
    }

    private static bool IsValidGeneratedPrintDocumentId(string? documentId)
    {
        return !string.IsNullOrWhiteSpace(documentId)
            && Regex.IsMatch(documentId, "^[A-Za-z0-9_-]+$");
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

    private async Task<bool> ExistsOpenSubProductionOrder(
        SapCompany company,
        string itemCode,
        int orderDocEntry,
        int visOrder,
        int orderLine,
        int productionBaseEntry)
    {
        var filter =
    $"ItemNo eq '{Escape(itemCode)}' " +
    $"and ProductionOrderOriginEntry eq {orderDocEntry} " +
    $"and U_RCS_BVO eq {visOrder} " +
    $"and U_RCS_OL eq {orderLine} " +
    $"and U_RCS_PB eq {productionBaseEntry} " +
    $"and (ProductionOrderStatus eq 'boposPlanned' " +
    $"or ProductionOrderStatus eq 'boposReleased')";
        var result = await _serviceLayer.GetStringAsync(company, "ProductionOrders", select: "AbsoluteEntry", filter: filter, top: 1);
        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return false;

        var root = JToken.Parse(result.Data) as JObject;
        var arr = root?["value"] as JArray;
        return arr != null && arr.Count > 0;
    }

    private async Task<int> ResolveCreatedSubProductionDocEntry(
        SapCompany company,
        string? createPayload,
        string itemCode,
        int orderDocEntry,
        int visOrder,
        int orderLine,
        int productionBaseEntry)
    {
        if (!string.IsNullOrWhiteSpace(createPayload))
        {
            try
            {
                var obj = JToken.Parse(createPayload) as JObject;
                var docEntry = GetInt(obj, "AbsoluteEntry");
                if (docEntry <= 0)
                    docEntry = GetInt(obj, "DocEntry");
                if (docEntry > 0)
                    return docEntry;
            }
            catch
            {
            }
        }

        var filter =
            $"ItemNo eq '{Escape(itemCode)}' and ProductionOrderOriginEntry eq {orderDocEntry} and U_RCS_BVO eq {visOrder} and U_RCS_OL eq {orderLine} and U_RCS_PB eq {productionBaseEntry}";
        var result = await _serviceLayer.GetStringAsync(company, "ProductionOrders", select: "AbsoluteEntry,DocEntry", filter: filter, top: 20);
        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return 0;

        var root = JToken.Parse(result.Data) as JObject;
        var values = root?["value"] as JArray;
        if (values == null || values.Count == 0)
            return 0;

        return values.OfType<JObject>()
            .Select(v =>
            {
                var docEntry = GetInt(v, "AbsoluteEntry");
                if (docEntry <= 0)
                    docEntry = GetInt(v, "DocEntry");
                return docEntry;
            })
            .OrderByDescending(v => v)
            .FirstOrDefault();
    }

    private static JObject? ResolveOrderLine(JObject orderObj, int? lineNum, string? itemCode)
    {
        var lines = orderObj["DocumentLines"] as JArray;
        if (lines == null || lines.Count == 0)
            return null;

        if (lineNum.HasValue)
        {
            var exactLine = lines
                .OfType<JObject>()
                .FirstOrDefault(l => GetInt(l, "LineNum") == lineNum.Value
                                  && (string.IsNullOrWhiteSpace(itemCode)
                                      || string.Equals(GetString(l, "ItemCode"), itemCode, StringComparison.OrdinalIgnoreCase)));
            if (exactLine != null)
                return exactLine;
        }

        if (!string.IsNullOrWhiteSpace(itemCode))
        {
            var byItem = lines
                .OfType<JObject>()
                .FirstOrDefault(l => string.Equals(GetString(l, "ItemCode"), itemCode, StringComparison.OrdinalIgnoreCase));
            if (byItem != null)
                return byItem;
        }

        return lines.OfType<JObject>().FirstOrDefault();
    }

    private async Task<bool> HasProductionBom(SapCompany company, string itemCode)
    {
        var result = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(itemCode)}')");
        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return false;

        var root = JToken.Parse(result.Data) as JObject;
        var treeType = GetString(root, "TreeType") ?? string.Empty;
        return treeType.Equals("iProductionTree", StringComparison.OrdinalIgnoreCase)
               || treeType.Equals("P", StringComparison.OrdinalIgnoreCase);
    }

    private async Task<bool> ExistsOpenProductionOrder(SapCompany company, string itemCode, int orderDocEntry, int visOrder, int lineNum)
    {
        var filter =
    $"ItemNo eq '{Escape(itemCode)}' " +
    $"and ProductionOrderOriginEntry eq {orderDocEntry} " +
    $"and U_RCS_BVO eq {visOrder} " +
    $"and U_RCS_OL eq {lineNum} " +
    $"and (ProductionOrderStatus eq 'boposPlanned' " +
    $"or ProductionOrderStatus eq 'boposReleased')";

        var result = await _serviceLayer.GetStringAsync(company, "ProductionOrders", filter: filter, top: 1);
        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return false;



        var root = JToken.Parse(result.Data) as JObject;
        var value = root?["value"] as JArray;
        return value != null && value.Count > 0;
    }

    private async Task<List<string>> GetSubBomCodes(SapCompany company, string rootItemCode)
    {
        var result = new List<string>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { rootItemCode };
        var queue = new Queue<string>();
        queue.Enqueue(rootItemCode);

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            var treeResult = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(current)}')");
            if (!treeResult.Success || string.IsNullOrEmpty(treeResult.Data))
                continue;

            var tree = JToken.Parse(treeResult.Data) as JObject;
            var lines = ResolveProductTreeLines(tree);
            if (lines == null)
                continue;

            foreach (var line in lines.OfType<JObject>())
            {
                var child = GetString(line, "ItemCode") ?? GetString(line, "Code") ?? string.Empty;
                if (string.IsNullOrWhiteSpace(child) || visited.Contains(child))
                    continue;

                visited.Add(child);
                var childResult = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(child)}')");
                if (!childResult.Success || string.IsNullOrEmpty(childResult.Data))
                    continue;

                var childTree = JToken.Parse(childResult.Data) as JObject;
                var childTreeType = GetString(childTree, "TreeType") ?? string.Empty;
                var isSubBom = childTreeType.Equals("iProductionTree", StringComparison.OrdinalIgnoreCase)
                               || childTreeType.Equals("P", StringComparison.OrdinalIgnoreCase);
                if (!isSubBom)
                    continue;

                result.Add(child);
                queue.Enqueue(child);
            }
        }

        return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static JArray? ResolveProductTreeLines(JObject? tree)
    {
        if (tree == null)
            return null;

        return tree["ProductTreeLines"] as JArray
               ?? tree["ProductTrees_Lines"] as JArray
               ?? tree["ProductTreesLines"] as JArray;
    }

    private async Task ApplyTemporarySubBomUpdates(
        SapCompany company,
        List<string> subBomCodes,
        List<SubBomAdjustmentDto>? adjustments,
        List<TempBomSnapshot> snapshots)
    {
        foreach (var code in subBomCodes)
        {
            var getResult = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(code)}')");
            if (!getResult.Success || string.IsNullOrEmpty(getResult.Data))
                continue;

            var tree = JToken.Parse(getResult.Data) as JObject;
            if (tree == null)
                continue;

            var originalPqt = GetDecimal(tree, "U_RCS_PQT");
            var originalOnSto = (GetString(tree, "U_RCS_ONSTO") ?? "Y").ToUpperInvariant();

            var requested = adjustments?.FirstOrDefault(a =>
                string.Equals(a.ItemCode, code, StringComparison.OrdinalIgnoreCase));

            var newPqt = requested?.U_RCS_PQT ?? 0m;
            var newOnSto = (requested?.U_RCS_ONSTO ?? "Y").ToUpperInvariant();
            if (newOnSto != "Y" && newOnSto != "N")
                newOnSto = "Y";

            var patch = new JObject
            {
                ["U_RCS_PQT"] = newPqt,
                ["U_RCS_ONSTO"] = newOnSto
            };

            var patchResult = await _serviceLayer.PatchAsync(company, $"ProductTrees('{Escape(code)}')", patch);
            if (!patchResult.Success)
            {
                _logger.LogWarning("Could not patch sub BOM {ItemCode}: {Error}", code, patchResult.Error);
                continue;
            }

            snapshots.Add(new TempBomSnapshot
            {
                ItemCode = code,
                U_RCS_PQT = originalPqt,
                U_RCS_ONSTO = originalOnSto
            });
        }
    }

    private async Task RestoreTemporarySubBomUpdates(SapCompany company, List<TempBomSnapshot> snapshots)
    {
        foreach (var snapshot in snapshots)
        {
            var patch = new JObject
            {
                ["U_RCS_PQT"] = snapshot.U_RCS_PQT,
                ["U_RCS_ONSTO"] = snapshot.U_RCS_ONSTO
            };

            var restore = await _serviceLayer.PatchAsync(company, $"ProductTrees('{Escape(snapshot.ItemCode)}')", patch);
            if (!restore.Success)
            {
                _logger.LogWarning("Failed restoring sub BOM {ItemCode}: {Error}", snapshot.ItemCode, restore.Error);
            }
        }
    }

    private static decimal ResolveQuantity(JObject line)
    {
        var pqt = GetDecimal(line, "U_RCS_PQT");
        if (pqt > 0)
            return pqt;

        return GetDecimal(line, "Quantity");
    }

    private static int ResolveDelDays(JObject order, JObject line)
    {
        var lineValue = GetInt(line, "U_RCS_DelDays");
        if (lineValue != 0)
            return lineValue;
        return GetInt(order, "U_RCS_DelDays");
    }

    private static DateTime ResolveDueDate(JObject order, JObject line)
    {
        var afd = GetDate(line, "U_RCS_AFS");
        if (afd.HasValue && !IsUnsetDate(afd.Value))
            return EnsureNotPast(afd.Value);

        var ship = GetDate(line, "ShipDate");
        if (ship.HasValue)
            return EnsureNotPast(ship.Value);

        var docDue = GetDate(order, "DocDueDate");
        return docDue.HasValue ? EnsureNotPast(docDue.Value) : DateTime.Today;
    }

    private static DateTime EnsureNotPast(DateTime date)
    {
        return date.Date < DateTime.Today ? DateTime.Today : date.Date;
    }


    private static bool IsUnsetDate(DateTime date)
    {
        return date.Year <= 1900;
    }

    private static string Escape(string value)
    {
        return value.Replace("'", "''");
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

    private static decimal GetDecimal(JToken? token, string name)
    {
        var value = token?[name];
        if (value == null || value.Type == JTokenType.Null)
            return 0m;

        if (value.Type == JTokenType.Float || value.Type == JTokenType.Integer)
            return value.Value<decimal>();

        if (decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
            return parsed;

        return 0m;
    }

    private static DateTime? GetDate(JToken? token, string name)
    {
        var value = token?[name];
        if (value == null || value.Type == JTokenType.Null)
            return null;

        if (value.Type == JTokenType.Date)
            return value.Value<DateTime>();

        if (DateTime.TryParse(value.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed))
            return parsed;

        return null;
    }

    private sealed class TempBomSnapshot
    {
        public string ItemCode { get; set; } = string.Empty;
        public decimal U_RCS_PQT { get; set; }
        public string U_RCS_ONSTO { get; set; } = "Y";
    }

    private sealed class CreateAllLineContext
    {
        public string ItemCode { get; set; } = string.Empty;
        public int LineNum { get; set; }
        public int VisOrder { get; set; }
        public decimal Quantity { get; set; }
        public DateTime DueDate { get; set; }
        public string Project { get; set; } = string.Empty;
        public int DelDays { get; set; }
    }
    public class CancelAllProductionsRequest
    {
        public int SalesOrderDocEntry { get; set; }
        public int? SalesOrderDocNum { get; set; }
        public string? DocumentId { get; set; }
    }
}
