using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using WorksheetAPI.Interfaces;

namespace WebAPI.Controllers;

/// <summary>
/// Endpoints til opslag og kopiering af produktionsstyklister.
/// </summary>
[ApiController]
[Route("api/bom")]
public class BomController : ControllerBase
{
    private readonly ISapConnectionService _sapConnection;
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<BomController> _logger;

    public BomController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        ILogger<BomController> logger)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _logger = logger;
    }

    /// <summary>
    /// Henter varer der kan bruges som kilde ved kopiering af rute eller stykliste.
    /// </summary>
    [HttpGet("copy-candidates")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetCopyCandidates()
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();

        // Hent varer med QryGroup1='Y' - samme logik som original SQL
        var result = await _serviceLayer.GetStringAsync(
            company,
            "Items",
            select: "ItemCode,ItemName",
            filter: "QryGroup1 eq 'tYES'",
            top: 500
        );

        if (!result.Success || string.IsNullOrEmpty(result.Data))
            return BadRequest(new { success = false, error = result.Error ?? "Kunne ikke hente varer" });

        var root = JToken.Parse(result.Data) as JObject;
        var items = root?["value"] as JArray ?? new JArray();

        var candidates = items.OfType<JObject>().Select(i => new
        {
            itemCode = i["ItemCode"]?.ToString(),
            itemName = i["ItemName"]?.ToString()
        }).ToList();

        return Ok(new { success = true, items = candidates });
    }

    /// <summary>
    /// Tjekker om der allerede findes en produktionsstykliste for varen.
    /// </summary>
    [HttpGet("check-exists/{itemCode}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> CheckBomExists(string itemCode)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();
        var result = await _serviceLayer.GetStringAsync(company, $"ProductTrees('{Escape(itemCode)}')");

        var exists = result.Success && !string.IsNullOrEmpty(result.Data);
        return Ok(new { success = true, exists, itemCode });
    }

    /// <summary>
    /// Kopierer en produktionsstykliste fra en kildevare til en targetvare.
    /// </summary>
    [HttpPost("copy-route")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> CopyRoute([FromBody] CopyRouteRequest request)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (string.IsNullOrWhiteSpace(request.SourceItemCode) || string.IsNullOrWhiteSpace(request.TargetItemCode))
            return BadRequest(new { success = false, error = "SourceItemCode og TargetItemCode er påkrævet" });

        var company = _sapConnection.GetMainCompany();

        // Hent kilde-styklisten
        var sourceResult = await _serviceLayer.GetStringAsync(
            company, $"ProductTrees('{Escape(request.SourceItemCode)}')");

        if (!sourceResult.Success || string.IsNullOrEmpty(sourceResult.Data))
            return BadRequest(new { success = false, error = $"Kunne ikke hente stykliste for {request.SourceItemCode}" });

        var sourceTree = JToken.Parse(sourceResult.Data) as JObject;
        if (sourceTree == null)
            return BadRequest(new { success = false, error = "Ugyldig stykliste-data fra SAP" });

        // Tjek at target ikke allerede eksisterer
        var targetCheck = await _serviceLayer.GetStringAsync(
            company, $"ProductTrees('{Escape(request.TargetItemCode)}')");

        if (targetCheck.Success && !string.IsNullOrEmpty(targetCheck.Data))
            return BadRequest(new
            {
                success = false,
                error = $"{request.TargetItemCode} er allerede oprettet som stykliste"
            });

        // Byg ny stykliste baseret på kilden
        var newTree = (JObject)sourceTree.DeepClone();
        newTree["TreeCode"] = request.TargetItemCode;

        // Fjern top-level read-only felter
        foreach (var field in new[] { "@odata.context", "odata.metadata", "AttachmentEntry" })
            newTree.Remove(field);

        // Rens linjer for read-only felter SAP ikke accepterer ved oprettelse
        var lines = newTree["ProductTreeLines"] as JArray;
        if (lines != null)
        {
            foreach (var line in lines.OfType<JObject>())
            {
                foreach (var field in new[] { "InventoryUOM", "WipAccount", "LineText", "ItemName" })
                    line.Remove(field);
            }
        }

        var createResult = await _serviceLayer.PostAsync<JObject>(company, "ProductTrees", newTree);

        if (!createResult.Success)
        {
            _logger.LogError("CopyRoute failed: source={Source}, target={Target}, error={Error}",
                request.SourceItemCode, request.TargetItemCode, createResult.Error);
            return BadRequest(new { success = false, error = createResult.Error ?? "Oprettelse fejlede" });
        }

        return Ok(new
        {
            success = true,
            message = $"Stykliste fra {request.SourceItemCode} kopieret til {request.TargetItemCode}"
        });
    }

    private static string Escape(string value) => value.Replace("'", "''");

    /// <summary>
    /// Requestmodel til kopiering af rute eller stykliste.
    /// </summary>
    public class CopyRouteRequest
    {
        /// <summary>
        /// Varekoden der kopieres fra.
        /// </summary>
        public string SourceItemCode { get; set; } = string.Empty;

        /// <summary>
        /// Varekoden der kopieres til.
        /// </summary>
        public string TargetItemCode { get; set; } = string.Empty;
    }
}