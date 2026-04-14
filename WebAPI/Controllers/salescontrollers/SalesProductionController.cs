using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Text.Json;
using System.Text;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;
using WebAPI.Models;
using WebAPI.Models.SalesProduction;

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
    private readonly IBomService _bomService;
    private readonly IProductionOrderHierarchyService _productionOrderHierarchyService;
    private readonly ISalesProductionPrintService _salesProductionPrintService;
    private readonly ILogger<SalesProductionController> _logger;

    public SalesProductionController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        IBomService bomService,
        IProductionOrderHierarchyService productionOrderHierarchyService,
        ISalesProductionPrintService salesProductionPrintService,
        ILogger<SalesProductionController> logger)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _bomService = bomService;
        _productionOrderHierarchyService = productionOrderHierarchyService;
        _salesProductionPrintService = salesProductionPrintService;
        _logger = logger;
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

        var hasBom = await _bomService.HasProductionBomAsync(company, itemCode);
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

        var subBomCodes = await _bomService.GetSubBomCodesAsync(company, itemCode);
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

        var originalValues = new List<SubBomSnapshot>();
        try
        {
            if (subBomCodes.Count > 0)
            {
                originalValues = await _bomService.ApplyTemporarySubBomUpdatesAsync(company, subBomCodes, request.SubBomAdjustments);
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
                    await _productionOrderHierarchyService.RemoveDeliveryFromProductionOrderAsync(company, createdDocEntry);
                }

                if (!uRcsOnSto.Equals("N", StringComparison.OrdinalIgnoreCase))
                {
                    await _productionOrderHierarchyService.CreateSubProductionOrdersForSubBomAsync(company, new CreateSubProductionOrdersRequest
                    {
                        ParentItemCode = itemCode,
                        ParentQuantity = quantity,
                        Project = lineProject,
                        ShipDate = dueDate,
                        CardCode = customerCode,
                        OrderDocEntry = salesOrderDocEntry,
                        OrderLine = lineNum,
                        VisOrder = visOrder,
                        ProductionBaseEntry = createdDocEntry,
                        ProductionBaseLine = 0,
                        RemoveDelivery = removeDelivery,
                        RcsDelDays = delDays,
                        Depth = 0
                    });
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
                await _bomService.RestoreTemporarySubBomUpdatesAsync(company, originalValues);
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

            var hasBom = await _bomService.HasProductionBomAsync(company, itemCode);
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

            var subBomCodes = await _bomService.GetSubBomCodesAsync(company, itemCode);
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

        var snapshots = new List<SubBomSnapshot>();
        var created = new List<object>();
        var failed = new List<object>();

        try
        {
            if (allSubBomCodes.Count > 0)
            {
                snapshots = await _bomService.ApplyTemporarySubBomUpdatesAsync(company, allSubBomCodes.ToList(), request.SubBomAdjustments);
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
                        await _productionOrderHierarchyService.RemoveDeliveryFromProductionOrderAsync(company, createdDocEntry);
                    }

                    await _productionOrderHierarchyService.CreateSubProductionOrdersForSubBomAsync(company, new CreateSubProductionOrdersRequest
                    {
                        ParentItemCode = line.ItemCode,
                        ParentQuantity = line.Quantity,
                        Project = line.Project,
                        ShipDate = line.DueDate,
                        CardCode = customerCode,
                        OrderDocEntry = salesOrderDocEntry,
                        OrderLine = line.LineNum,
                        VisOrder = line.VisOrder,
                        ProductionBaseEntry = createdDocEntry,
                        ProductionBaseLine = 0,
                        RemoveDelivery = removeDelivery,
                        RcsDelDays = line.DelDays,
                        Depth = 0
                    });
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
                await _bomService.RestoreTemporarySubBomUpdatesAsync(company, snapshots);
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
            var docNum = GetInt(order, "DocumentNumber");

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
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (request.SalesOrderDocEntry <= 0 && (!request.SalesOrderDocNum.HasValue || request.SalesOrderDocNum <= 0))
            return BadRequest(new { success = false, error = "SalesOrderDocEntry or SalesOrderDocNum is required" });

        var company = _sapConnection.GetMainCompany();
        var baseApiUrl = $"{Request.Scheme}://{Request.Host}{Request.PathBase}".TrimEnd('/');
        var result = await _salesProductionPrintService.GeneratePrintProductionOrdersAsync(company, request, baseApiUrl);
        return Ok(result);
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
        if (!_salesProductionPrintService.IsValidGeneratedPrintDocumentId(documentId))
            return BadRequest(new { ready = false, error = "Ugyldigt documentId." });

        if (!_salesProductionPrintService.TryGetGeneratedPrint(documentId, out var document) || document == null)
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
            openUrl = _salesProductionPrintService.BuildGeneratedPrintOpenUrl($"{Request.Scheme}://{Request.Host}{Request.PathBase}", documentId),
            downloadUrl = _salesProductionPrintService.BuildGeneratedPrintDownloadUrl($"{Request.Scheme}://{Request.Host}{Request.PathBase}", documentId),
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
        if (!_salesProductionPrintService.IsValidGeneratedPrintDocumentId(documentId))
        {
            return Content("Ugyldigt documentId.", "text/plain; charset=utf-8", Encoding.UTF8);
        }

        var statusUrl = _salesProductionPrintService.BuildGeneratedPrintStatusUrl($"{Request.Scheme}://{Request.Host}{Request.PathBase}", documentId) ?? string.Empty;
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
        if (!_salesProductionPrintService.TryGetGeneratedPrint(documentId, out var document) || document == null)
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
        if (!_salesProductionPrintService.TryGetGeneratedPrint(documentId, out var document) || document == null)
            return NotFound();

        Response.Headers.CacheControl = "no-store, no-cache";

        if (!document.OpenInline)
            return File(document.Content, document.ContentType, document.FileName);

        Response.Headers.ContentDisposition = $"inline; filename*=UTF-8''{Uri.EscapeDataString(document.FileName)}";
        return File(document.Content, document.ContentType);
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

}
