using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Services;

namespace WebAPI.Controllers.worksheetscontrollers;

/// <summary>
/// Endpoints til læsning og opdatering af worksheets fra UDT-tabellerne AIG_KLADDE.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class WorksheetsController : ControllerBase
{
    private readonly ISapConnectionService _sapConnection;
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<WorksheetsController> _logger;

    public WorksheetsController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        ILogger<WorksheetsController> logger)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _logger = logger;
    }

    /// <summary>
    /// Henter worksheet-headere fra U_AIG_KLADDE_01.
    /// </summary>
    ///     
    [HttpGet]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetWorksheets([FromQuery] string? status = null, [FromQuery] int? top = 1000)
    {
        var sw = Stopwatch.StartNew();
        if (!_sapConnection.IsSapEnabled)
        {
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        }

        var company = _sapConnection.GetMainCompany();
        
        string? filter = null;
        if (!string.IsNullOrEmpty(status) && status != "ALL")
        {
            filter = $"U_status eq '{status}'";
        }

        var result = await _serviceLayer.GetStringAsync(
            company,
            "U_AIG_KLADDE_01",
            filter: filter,
            top: top
        );

        sw.Stop();
        Response.Headers["Server-Timing"] = $"worksheets;dur={sw.Elapsed.TotalMilliseconds:0.00}";

        if (!result.Success)
        {
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get worksheets" });
        }

        var logSw = Stopwatch.StartNew();
        _logger.LogInformation("[DEBUG] Starting BuildSuccessResponse for worksheets");
        var jsonResponse = BuildSuccessResponse(result.Data);
        logSw.Stop();
        _logger.LogInformation("[DEBUG] BuildSuccessResponse for worksheets took {ElapsedMs} ms", logSw.Elapsed.TotalMilliseconds);
        // Log response size (records and bytes)
        try
        {
            var token = JToken.Parse(jsonResponse);
            int count = 0;
            if (token["data"] is JArray arr) count = arr.Count;
            var byteSize = System.Text.Encoding.UTF8.GetByteCount(jsonResponse);
            _logger.LogInformation("[DEBUG] Response contains {Count} records, {ByteSize} bytes", count, byteSize);
        }
        catch { }
        return Content(jsonResponse, "application/json");
    }

    /// <summary>
    /// Henter et enkelt worksheet via worksheet-nummer.
    /// </summary>
    [HttpGet("{worksheetNo}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetWorksheet(string worksheetNo)
    {
        var sw = Stopwatch.StartNew();
        if (!_sapConnection.IsSapEnabled)
        {
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        }

        var company = _sapConnection.GetMainCompany();
        var slSw = Stopwatch.StartNew();
        _logger.LogInformation("[DEBUG] Starting ServiceLayer GetStringAsync for worksheet header");
        var result = await _serviceLayer.GetStringAsync(
            company,
            "U_AIG_KLADDE_01",
            filter: $"U_nr eq {worksheetNo}"
        );
        slSw.Stop();
        _logger.LogInformation("[DEBUG] ServiceLayer GetStringAsync for worksheet header took {ElapsedMs} ms", slSw.Elapsed.TotalMilliseconds);

        sw.Stop();
        Response.Headers["Server-Timing"] = $"worksheet;dur={sw.Elapsed.TotalMilliseconds:0.00}";

        if (!result.Success)
        {
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get worksheet" });
        }

        var logSw = Stopwatch.StartNew();
        _logger.LogInformation("[DEBUG] Starting BuildSuccessResponseFirst for worksheet");
        var jsonResponse = BuildSuccessResponseFirst(result.Data);
        logSw.Stop();
        _logger.LogInformation("[DEBUG] BuildSuccessResponseFirst for worksheet took {ElapsedMs} ms", logSw.Elapsed.TotalMilliseconds);
        return Content(jsonResponse, "application/json");
    }

    /// <summary>
    /// Henter worksheet-linjer fra U_AIG_KLADDE_02.
    /// </summary>
    [HttpGet("{worksheetNo}/lines")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetWorksheetLines(string worksheetNo, [FromQuery] int? top = 1000)
    {
        var sw = Stopwatch.StartNew();
        if (!_sapConnection.IsSapEnabled)
        {
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        }

        var company = _sapConnection.GetMainCompany();
        var result = await _serviceLayer.GetStringAsync(
            company,
            "U_AIG_KLADDE_02",
            filter: $"U_buntnr eq {worksheetNo}"
        );

        sw.Stop();
        Response.Headers["Server-Timing"] = $"worksheetLines;dur={sw.Elapsed.TotalMilliseconds:0.00}";

        if (!result.Success)
        {
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get worksheet lines" });
        }

        var logSw = Stopwatch.StartNew();
        _logger.LogInformation("[DEBUG] Starting BuildSuccessResponse for worksheet lines");
        var jsonResponse = BuildSuccessResponse(result.Data);
        logSw.Stop();
        _logger.LogInformation("[DEBUG] BuildSuccessResponse for worksheet lines took {ElapsedMs} ms", logSw.Elapsed.TotalMilliseconds);
        // Log response size (records and bytes)
        try
        {
            var token = JToken.Parse(jsonResponse);
            int count = 0;
            if (token["data"] is JArray arr) count = arr.Count;
            var byteSize = System.Text.Encoding.UTF8.GetByteCount(jsonResponse);
            _logger.LogInformation("[DEBUG] Response contains {Count} records, {ByteSize} bytes", count, byteSize);
        }
        catch { }
        return Content(jsonResponse, "application/json");
    }

    /// <summary>
    /// Opdaterer worksheet-linjer i U_AIG_KLADDE_02.
    /// </summary>
    [HttpPatch("{worksheetNo}/lines")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> UpdateWorksheetLines(string worksheetNo, [FromBody] WorksheetAPI.Models.WorksheetLinesUpdateDto linesData)
    {
        if (!_sapConnection.IsSapEnabled)
        {
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        }

        var company = _sapConnection.GetMainCompany();
        try
        {
            var lines = linesData.Lines;
            if (lines == null)
            {
                return BadRequest(new { success = false, error = "Missing or invalid 'lines' array in request body" });
            }

            foreach (var line in lines)
            {
                // Centralized mapping and business logic
                var mapped = MapLineToSapFields(line);
                var code = mapped.ContainsKey("Code") ? mapped["Code"]?.ToString() : null;
                if (string.IsNullOrEmpty(code)) continue;
                mapped.Remove("Code");

                // Always fetch current SAP line and merge
                var getResult = await _serviceLayer.GetStringAsync(
                    company,
                    "U_AIG_KLADDE_02",
                    filter: $"Code eq '{code}'"
                );
                Dictionary<string, object> merged = new Dictionary<string, object>();
                if (getResult.Success && !string.IsNullOrEmpty(getResult.Data))
                {
                    var token = Newtonsoft.Json.Linq.JToken.Parse(getResult.Data);
                    var valueArr = token["value"] as Newtonsoft.Json.Linq.JArray;
                    if (valueArr != null && valueArr.Count > 0)
                    {
                        var currentLine = valueArr[0];
                        // Copy all current SAP fields
                        foreach (var prop in currentLine.Children<Newtonsoft.Json.Linq.JProperty>())
                        {
                            if (prop.Value != null && !string.IsNullOrEmpty(prop.Value.ToString()))
                                merged[prop.Name] = prop.Value.ToString();
                        }
                    }
                }
                // Overwrite with incoming mapped fields (including null/empty)
                foreach (var kv in mapped)
                {
                    if (kv.Value != null && !string.IsNullOrEmpty(kv.Value.ToString()))
                        merged[kv.Key] = kv.Value;
                    else if (merged.ContainsKey(kv.Key))
                        merged.Remove(kv.Key); // Remove if incoming is null/empty
                }
                // Only allow valid SAP fields for @AIG_KLADDE_02 (kun felter der typisk må opdateres)
                var validFields = new HashSet<string> {
                    "U_typ", "U_beskr", "U_Dkonto", "U_Kkonto",
                    "U_valkode", "U_mvakode", "U_difference", "U_ref1", "U_ref2",
                    "U_betalldato", "U_OK", "Code"
                };
                var keysToRemove = merged.Keys.Where(k => !validFields.Contains(k) || merged[k] == null || string.IsNullOrEmpty(merged[k]?.ToString())).ToList();
                foreach (var k in keysToRemove)
                    merged.Remove(k);

                var updateObj = Newtonsoft.Json.Linq.JObject.FromObject(merged);
                var resource = $"U_AIG_KLADDE_02('{code}')";
                var result = await _serviceLayer.PatchAsync(
                    company,
                    resource,
                    updateObj
                );
                if (!result.Success)
                {
                    return BadRequest(new { success = false, error = result.Error ?? $"Failed to update line {code}" });
                }
            }

            return Ok(new { success = true });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error updating worksheet lines");
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    /// <summary>
    /// Centralized mapping and business logic for worksheet lines
    /// </summary>
    private static Dictionary<string, object> MapLineToSapFields(Dictionary<string, object> line)
    {
        var mapped = new Dictionary<string, object>();
        void AddIfPresent(string sapKey, string uiKey)
        {
            if (line.ContainsKey(uiKey) && line[uiKey] != null)
            {
                mapped[sapKey] = line[uiKey];
            }
        }

        AddIfPresent("Code", "code");
        AddIfPresent("U_bilagsnr", "docNum");
        AddIfPresent("U_typ", "docType");
        AddIfPresent("U_beskr", "description");
        AddIfPresent("U_Dkonto", "debitAccount");
        AddIfPresent("U_Kkonto", "creditAccount");
        AddIfPresent("U_debetLV", "debitLC");
        AddIfPresent("U_kreditLV", "creditLC");
        AddIfPresent("U_samletLV", "totalLC");
        AddIfPresent("U_debetUV", "debitFC");
        AddIfPresent("U_kreditUV", "creditFC");
        AddIfPresent("U_samletUV", "totalFC");
        AddIfPresent("U_valkode", "currency");
        AddIfPresent("U_mvakode", "vatCode");
        AddIfPresent("U_ref1", "ref1");
        AddIfPresent("U_ref2", "ref2");
        AddIfPresent("U_betalldato", "cancelDate");
        // Business logic: ok (UI) -> U_OK (SAP)
        if (line.ContainsKey("ok"))
        {
            var okVal = line["ok"];
            var okStr = okVal?.ToString()?.ToLower();
            bool isTrue = okVal is bool b ? b : okStr == "true";
            mapped["U_OK"] = isTrue? "Y" : "N";
        }
        // Add more mappings and business logic as needed
        return mapped;
    }
    

    /// <summary>
    /// Henter worksheet-typer fra U_AIG_KLADDE_04.
    /// </summary>
    [HttpGet("types")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetWorksheetTypes()
    {
        var sw = Stopwatch.StartNew();
        if (!_sapConnection.IsSapEnabled)
        {
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        }

        var company = _sapConnection.GetMainCompany();
        var result = await _serviceLayer.GetStringAsync(
            company,
            "U_AIG_KLADDE_04"
        );

        sw.Stop();
        Response.Headers["Server-Timing"] = $"worksheetTypes;dur={sw.Elapsed.TotalMilliseconds:0.00}";

        if (!result.Success)
        {
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get worksheet types" });
        }

        var jsonResponse = BuildSuccessResponse(result.Data);
        return Content(jsonResponse, "application/json");
    }    
    /// <summary>
    /// Build success response by extracting value array from SAP response
    /// </summary>
    private static string BuildSuccessResponse(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse))
            return "{\"success\":true,\"data\":[]}";

        try
        {
            var token = JToken.Parse(sapResponse);
            
            // If it's an object with a "value" property, extract the value array
            if (token is JObject obj && obj.TryGetValue("value", out var value))
            {
                return $"{{\"success\":true,\"data\":{value}}}";
            }
            
            // Otherwise use the token as-is
            return $"{{\"success\":true,\"data\":{token}}}";
        }
        catch
        {
            return "{\"success\":true,\"data\":[]}";
        }
    }

    /// <summary>
    /// Build success response with first item from value array
    /// </summary>
    private static string BuildSuccessResponseFirst(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse))
            return "{\"success\":true,\"data\":null}";

        try
        {
            var token = JToken.Parse(sapResponse);
            
            // If it's an object with a "value" property, extract first item
            if (token is JObject obj && obj.TryGetValue("value", out var value) && value is JArray arr && arr.Count > 0)
            {
                return $"{{\"success\":true,\"data\":{arr[0]}}}";
            }
            
            return "{\"success\":true,\"data\":null}";
        }
        catch
        {
            return "{\"success\":true,\"data\":null}";
        }
    }
}
