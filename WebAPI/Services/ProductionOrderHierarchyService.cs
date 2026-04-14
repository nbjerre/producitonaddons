using System.Globalization;
using Newtonsoft.Json.Linq;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;
using WebAPI.Models.SalesProduction;

namespace WorksheetAPI.Services;

/// <summary>
/// Service for recursive sub-BOM production-order orchestration.
/// </summary>
public class ProductionOrderHierarchyService : IProductionOrderHierarchyService
{
    private readonly IServiceLayerService _serviceLayer;
    private readonly IBomService _bomService;
    private readonly ILogger<ProductionOrderHierarchyService> _logger;

    public ProductionOrderHierarchyService(
        IServiceLayerService serviceLayer,
        IBomService bomService,
        ILogger<ProductionOrderHierarchyService> logger)
    {
        _serviceLayer = serviceLayer;
        _bomService = bomService;
        _logger = logger;
    }

    public async Task RemoveDeliveryFromProductionOrderAsync(SapCompany company, int docEntry)
    {
        if (docEntry <= 0)
            return;

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

    public async Task CreateSubProductionOrdersForSubBomAsync(SapCompany company, CreateSubProductionOrdersRequest request)
    {
        if (request == null)
            return;

        if (string.IsNullOrWhiteSpace(request.ParentItemCode) || request.OrderDocEntry <= 0)
            return;

        if (request.Depth > 10)
            return;

        var treeResult = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(request.ParentItemCode)}')");
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

            if (!await _bomService.HasProductionBomAsync(company, childItem))
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
            var createQty = forcedQty > 0 ? forcedQty : Math.Max(request.ParentQuantity * lineQty, 1m);

            if (await ExistsOpenSubProductionOrder(company, childItem, request.OrderDocEntry, request.VisOrder, request.OrderLine, request.ProductionBaseEntry))
                continue;

            var payload = new JObject
            {
                ["ItemNo"] = childItem,
                ["CustomerCode"] = request.CardCode,
                ["DueDate"] = request.ShipDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                ["ProductionOrderOrigin"] = "bopooSalesOrder",
                ["ProductionOrderOriginEntry"] = request.OrderDocEntry,
                ["ProductionOrderType"] = "bopotStandard",
                ["PlannedQuantity"] = createQty,
                ["Project"] = request.Project,
                ["U_RCS_BVO"] = request.VisOrder,
                ["U_RCS_OL"] = request.OrderLine,
                ["U_RCS_PB"] = request.ProductionBaseEntry,
                ["U_RCS_PBL"] = request.ProductionBaseLine,
                ["U_RCS_DelDays"] = request.RcsDelDays
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
                request.OrderDocEntry,
                request.VisOrder,
                request.OrderLine,
                request.ProductionBaseEntry);

            if (childProdEntry <= 0)
                continue;

            if (request.RemoveDelivery)
            {
                await RemoveDeliveryFromProductionOrderAsync(company, childProdEntry);
            }

            await CreateSubProductionOrdersForSubBomAsync(company, new CreateSubProductionOrdersRequest
            {
                ParentItemCode = childItem,
                ParentQuantity = createQty,
                Project = request.Project,
                ShipDate = request.ShipDate,
                CardCode = request.CardCode,
                OrderDocEntry = request.OrderDocEntry,
                OrderLine = request.OrderLine,
                VisOrder = request.VisOrder,
                ProductionBaseEntry = childProdEntry,
                ProductionBaseLine = GetInt(line, "LineNum"),
                RemoveDelivery = request.RemoveDelivery,
                RcsDelDays = request.RcsDelDays,
                Depth = request.Depth + 1
            });
        }
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

    private static JArray? ResolveProductTreeLines(JObject? tree)
    {
        if (tree == null)
            return null;

        return tree["ProductTreeLines"] as JArray
               ?? tree["ProductTrees_Lines"] as JArray
               ?? tree["ProductTreesLines"] as JArray;
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
}
