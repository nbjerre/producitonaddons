using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using WorksheetAPI.Interfaces;

namespace WebAPI.Controllers.userscontroller;

/// <summary>
/// Endpoints til konti, omkostningscentre og medarbejderes account-links.
/// </summary>
[ApiController]
[Route("api")]
public class AccountsController : ControllerBase
{
    // ─── UDT TABLE NAME ───────────────────────────────────────────────────
    // SAP B1 UDT: @EMP_ACC_CC_LINK  →  Service Layer entity: U_EMP_ACC_CC_LINK
    // Fields:
    //   Code          (PK, string, max 20 chars)
    //   Name          (string)
    //   U_EmployeeID  (int)
    //   U_AcctCode    (nvarchar)
    //   U_PrcCode     (nvarchar)
    // ─────────────────────────────────────────────────────────────────────
    private const string UDT_LINKS = "EMP_ACC_CC_LINK";

    private readonly ISapConnectionService _sapConnection;
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<AccountsController> _logger;

    public AccountsController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        ILogger<AccountsController> logger)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _logger = logger;
    }

    // ─────────────────────────────────────────────────────────────────────
    // GET: api/accounts
    // Returns all GL accounts from SAP ChartOfAccounts (OACT)
    // ─────────────────────────────────────────────────────────────────────
    // ─────────────────────────────────────────────────────────────────────
    // GET: api/accounts?search=%søgeterm%
    //
    // Bruges af Tab+* finanskonto-opslag i worksheet-matrixen.
    // Svarer til AIG_FindAC.FindAC i gammel SAP B1 addon:
    //
    //   SELECT AcctCode, AcctName FROM OACT
    //   WHERE ((AcctCode LIKE 'søgeterm') OR (AcctName LIKE 'søgeterm'))
    //   AND Postable = 'Y'
    //   ORDER BY AcctCode
    //
    // Frontend erstatter bruger-wildcard (*) med SQL-wildcard (%) inden kald.
    // Returnerer max 200 resultater — som den gamle grid-form i SAP B1.
    // ─────────────────────────────────────────────────────────────────────
    /// <summary>
    /// Henter bogførbare finanskonti med valgfrit søgefilter.
    /// </summary>
    [HttpGet("accounts")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetAccounts(
        [FromQuery] string? search = null,
        [FromQuery] int? top = null)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();

        // Hent alle konti med parent-info så vi kan beregne Postable
        var result = await _serviceLayer.GetStringAsync(company, "ChartOfAccounts",
            select: "Code,Name,FatherAccountKey",
            top: top ?? 5000);   // hent rigeligt til client-side filter

        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get accounts" });

        try
        {
            var token = JToken.Parse(result.Data!);
            var arr = token["value"] as JArray ?? new JArray();

            // Beregn hvilke konti der er title-konti (ikke bogførbare)
            var parentCodes = arr
                .Where(a => !string.IsNullOrEmpty(a["FatherAccountKey"]?.ToString()))
                .Select(a => a["FatherAccountKey"]!.ToString())
                .Distinct()
                .ToHashSet();

            // Map til AcctCode / AcctName og filtrer på Postable
            var accounts = arr
                .Select(a => new {
                    AcctCode = a["Code"]?.ToString() ?? "",
                    AcctName = a["Name"]?.ToString() ?? "",
                    Postable = !parentCodes.Contains(a["Code"]?.ToString() ?? "")
                })
                .Where(a => a.Postable)          // kun bogførbare — svarer til Postable='Y'
                .OrderBy(a => a.AcctCode)
                .ToList();

            // Søgefilter — svarer til VB: AcctCode LIKE søgeterm OR AcctName LIKE søgeterm
            // Frontend sender "%" som wildcard (f.eks. "%115%" eller "115%")
            if (!string.IsNullOrWhiteSpace(search))
            {
                // Konvertér SQL LIKE-mønster til C# string-sammenligning
                // Understøttede mønstre: %...%, ...%, %...
                var s = search.Trim();

                accounts = accounts
                    .Where(a => LikeMatch(a.AcctCode, s) || LikeMatch(a.AcctName, s))
                    .ToList();
            }

            // Begræns resultater som den gamle SAP B1 grid (max 200)
            var limited = accounts.Take(200);

            return Ok(new { success = true, data = limited });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error parsing accounts response");
            return Ok(new { success = true, data = Array.Empty<object>() });
        }
    }

    /// <summary>
    /// Simpel SQL LIKE-emulering til client-side søgning.
    /// Understøtter % som wildcard (begyndelse, slutning eller begge).
    /// Case-insensitiv.
    /// </summary>
    private static bool LikeMatch(string value, string pattern)
    {
        if (string.IsNullOrEmpty(pattern) || pattern == "%")
            return true;

        var v = value.ToLowerInvariant();
        var p = pattern.ToLowerInvariant();

        bool startsWild = p.StartsWith("%");
        bool endsWild = p.EndsWith("%");
        var core = p.Trim('%');

        if (string.IsNullOrEmpty(core)) return true;       // kun wildcards

        if (startsWild && endsWild) return v.Contains(core);
        if (startsWild) return v.EndsWith(core);
        if (endsWild) return v.StartsWith(core);
        return v == core;                                   // eksakt match
    }


    // ─────────────────────────────────────────────────────────────────────
    // GET: api/costcenters
    // Returns all cost centers from SAP ProfitCenters (OPRC)
    // ─────────────────────────────────────────────────────────────────────
    /// <summary>
    /// Henter omkostningscentre fra SAP ProfitCenters.
    /// </summary>
    [HttpGet("costcenters")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetCostCenters([FromQuery] int? top = null)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();

        var result = await _serviceLayer.GetStringAsync(company, "ProfitCenters",
            select: "CenterCode,CenterName",            
            top: top);

        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get cost centers" });

        try
        {
            var token = JToken.Parse(result.Data!);
            var arr = token["value"] as JArray ?? new JArray();
            var mapped = arr
    .Where(a =>        
        !(a["CenterCode"]?.ToString().StartsWith("Stelle_") ?? false)
    )
    .Select(a => new {
        PrcCode = a["CenterCode"]?.ToString(),
        PrcName = a["CenterName"]?.ToString()
    });
            return Ok(new { success = true, data = mapped });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error parsing cost centers response");
            return Ok(new { success = true, data = Array.Empty<object>() });
        }
    }

    // ─────────────────────────────────────────────────────────────────────
    // GET: api/employees/{employeeId}/account-links
    // Returns the employee's current account + cost-center assignments
    // grouped by account:
    //   [ { acctCode: "1000", prcCodes: ["CC1", "CC2"] }, ... ]
    // ─────────────────────────────────────────────────────────────────────
    /// <summary>
    /// Henter account-links for en medarbejder grupperet pr. konto.
    /// </summary>
    [HttpGet("employees/{employeeId}/account-links")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetAccountLinks(int employeeId)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();

        var result = await _serviceLayer.GetStringAsync(company, UDT_LINKS,
            filter: $"U_EmployeeID eq {employeeId}");

        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get account links" });

        try
        {
            var token = JToken.Parse(result.Data!);
            var arr = token["value"] as JArray ?? new JArray();

            // Group rows by AcctCode, collect PrcCodes for each
            var grouped = arr
                .GroupBy(r => r["U_AcctCode"]?.ToString() ?? "")
                .Where(g => !string.IsNullOrEmpty(g.Key))
                .Select(g => new {
                    acctCode = g.Key,
                    prcCodes = g
                        .Select(r => r["U_PrcCode"]?.ToString())
                        .Where(c => c != null)
                        .ToList()
                });

            return Ok(new { success = true, data = grouped });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error parsing account links response");
            return Ok(new { success = true, data = Array.Empty<object>() });
        }
    }

    // ─────────────────────────────────────────────────────────────────────
    // POST: api/employees/{employeeId}/account-links
    // Replaces ALL account + cost-center links for the employee.
    //
    // Body:
    // [
    //   { "acctCode": "1000", "prcCodes": ["CC1", "CC2"] },
    //   { "acctCode": "1100", "prcCodes": ["CC1"] }
    // ]
    //
    // Strategy: delete all existing rows for employee, then insert incoming rows.
    // Code (PK) format: "{employeeId}_{sequenceIndex}"  (max 20 chars, safe up to billions)
    // ─────────────────────────────────────────────────────────────────────
    /// <summary>
    /// Erstatter alle account-links for en medarbejder med de links der sendes i body.
    /// </summary>
    [HttpPost("employees/{employeeId}/account-links")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> SetAccountLinks(int employeeId, [FromBody] List<AccountLinkDto> links)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();
        links ??= new List<AccountLinkDto>();

        // ── 1. Fetch all existing rows for this employee ──────────────────
        var getResult = await _serviceLayer.GetStringAsync(company, UDT_LINKS,
            filter: $"U_EmployeeID eq {employeeId}");

        var existingCodes = new List<string>();
        if (getResult.Success && !string.IsNullOrEmpty(getResult.Data))
        {
            var token = JToken.Parse(getResult.Data);
            var arr = token["value"] as JArray ?? new JArray();
            existingCodes = arr
                .Select(r => r["Code"]?.ToString())
                .Where(c => !string.IsNullOrEmpty(c))
                .ToList()!;
        }

        // ── 2. Delete all existing rows ───────────────────────────────────
        foreach (var code in existingCodes)
        {
            var delResult = await _serviceLayer.DeleteAsync(company, $"{UDT_LINKS}('{code}')");
            if (!delResult.Success)
                _logger.LogWarning("[AccountsController] Failed to delete link {Code}: {Error}", code, delResult.Error);
        }

        // ── 3. Insert new rows ────────────────────────────────────────────
        // Expand each {acctCode, prcCodes[]} into one row per (acctCode, prcCode) pair.
        // If prcCodes is empty but account is selected, insert one row with empty prcCode
        // so the account selection is preserved even without a cost center choice.
        var seq = 0;
        var insertErrors = 0;

        foreach (var link in links)
        {
            if (string.IsNullOrEmpty(link.AcctCode)) continue;

            var prcList = (link.PrcCodes != null && link.PrcCodes.Count > 0)
                ? link.PrcCodes
                : new List<string> { "" };   // account selected with no cost center

            foreach (var prcCode in prcList)
            {
                seq++;
                var pkCode = $"{employeeId}_{seq}";

                // Safety: SAP Code field is max 20 chars
                if (pkCode.Length > 20)
                    pkCode = pkCode.Substring(0, 20);

                var row = new JObject
                {
                    ["Code"]         = pkCode,
                    ["Name"]         = pkCode,
                    ["U_EmployeeID"] = employeeId.ToString(),
                    ["U_AcctCode"]   = link.AcctCode,
                    ["U_PrcCode"]    = prcCode
                };

                var addResult = await _serviceLayer.PostAsync<JObject>(company, UDT_LINKS, row);
                if (!addResult.Success)
                {
                    insertErrors++;
                    _logger.LogWarning("[AccountsController] Failed to insert link (emp={EmpId}, acct={Acct}, prc={Prc}): {Error}",
                        employeeId, link.AcctCode, prcCode, addResult.Error);
                }
            }
        }

        _logger.LogInformation(
            "[AccountsController] Saved account-links for employee {EmpId}: {Deleted} deleted, {Inserted} inserted, {Errors} errors",
            employeeId, existingCodes.Count, seq, insertErrors);

        if (insertErrors > 0)
            return Ok(new { success = true, warning = $"{insertErrors} row(s) failed to save — check server logs" });

        return Ok(new { success = true });
    }
}

[ApiController]
[Route("api/[controller]")]
public class PingController : ControllerBase
{
    /// <summary>
    /// Simpelt ping-endpoint til healthcheck og forbindelsestest.
    /// </summary>
    [HttpGet]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    public IActionResult Get()
    {
        return Ok(new { message = "pong", time = DateTime.UtcNow });
    }
}
// ─── DTOs ────────────────────────────────────────────────────────────────────

/// <summary>
/// One account with its list of allowed cost-center codes.
/// </summary>
public class AccountLinkDto
{
    /// <summary>
    /// Finanskontonummer.
    /// </summary>
    public string? AcctCode { get; set; }

    /// <summary>
    /// Liste af tilladte omkostningscentre for kontoen.
    /// </summary>
    public List<string>? PrcCodes { get; set; }
}

/// <summary>
/// Legacy DTO kept for any existing callers.
/// </summary>
public class AccessUpdateDto
{
    public List<string>? Codes { get; set; }
}