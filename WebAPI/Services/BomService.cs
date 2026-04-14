using System.Globalization;
using Newtonsoft.Json.Linq;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;

namespace WorksheetAPI.Services;

/// <summary>
/// Centralized BOM operations used by production order flows.
/// </summary>
public class BomService : IBomService
{
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<BomService> _logger;

    public BomService(IServiceLayerService serviceLayer, ILogger<BomService> logger)
    {
        _serviceLayer = serviceLayer;
        _logger = logger;
    }

    public async Task<bool> HasProductionBomAsync(SapCompany company, string itemCode)
    {
        if (string.IsNullOrWhiteSpace(itemCode))
            return false;

        var result = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(itemCode)}')");
        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return false;

        var root = JToken.Parse(result.Data) as JObject;
        var treeType = GetString(root, "TreeType") ?? string.Empty;
        return treeType.Equals("iProductionTree", StringComparison.OrdinalIgnoreCase)
               || treeType.Equals("P", StringComparison.OrdinalIgnoreCase);
    }

    public async Task<List<string>> GetSubBomCodesAsync(SapCompany company, string rootItemCode)
    {
        if (string.IsNullOrWhiteSpace(rootItemCode))
            return new List<string>();

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

    public async Task<List<SubBomSnapshot>> ApplyTemporarySubBomUpdatesAsync(
        SapCompany company,
        List<string> subBomCodes,
        List<SubBomAdjustmentDto>? adjustments)
    {
        var snapshots = new List<SubBomSnapshot>();

        if (subBomCodes == null || subBomCodes.Count == 0)
            return snapshots;

        foreach (var code in subBomCodes
            .Where(c => !string.IsNullOrWhiteSpace(c))
            .Distinct(StringComparer.OrdinalIgnoreCase))
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

            snapshots.Add(new SubBomSnapshot
            {
                ItemCode = code,
                U_RCS_PQT = originalPqt,
                U_RCS_ONSTO = originalOnSto
            });
        }

        return snapshots;
    }

    public async Task RestoreTemporarySubBomUpdatesAsync(SapCompany company, List<SubBomSnapshot> snapshots)
    {
        if (snapshots == null || snapshots.Count == 0)
            return;

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
