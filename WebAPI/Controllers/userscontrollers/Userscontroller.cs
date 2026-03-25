using Microsoft.AspNetCore.Mvc;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;
using Newtonsoft.Json.Linq;

namespace WebAPI.Controllers.userscontroller;

/// <summary>
/// Endpoints til læsning og vedligeholdelse af SAP-brugere.
/// </summary>
[ApiController]
[Route("api/users")]
public class UsersController : ControllerBase
{
    private readonly ISapConnectionService _sapConnection;
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<UsersController> _logger;

    public UsersController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        ILogger<UsersController> logger)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _logger = logger;
    }

    /// <summary>
    /// Henter alle SAP-brugere.
    /// </summary>
    [HttpGet]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetAllUsers()
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();
        var result = await _serviceLayer.GetStringAsync(company, "Users");
        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get users" });

        return Content(BuildSuccessResponse(result.Data), "application/json");
    }

    /// <summary>
    /// Finder en SAP-bruger via InternalKey.
    /// </summary>
    [HttpGet("by-internal/{internalKey}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetUserByInternalKey(int internalKey)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();

        // SAP Service Layer OData field name for OUSR.INTERNAL_K is "InternalKey"
        // If this returns no results, check your metadata:
        // GET https://[your-sap]:50000/b1s/v1/$metadata and search for "Users" entity
        // The field may be named "InternalKey", "UserID", or similar
        var result = await _serviceLayer.GetStringAsync(company, "Users", filter: $"InternalKey eq {internalKey}");

        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get user" });

        return Content(BuildSuccessResponseFirst(result.Data), "application/json");
    }

    /// <summary>
    /// Henter en SAP-bruger via UserCode.
    /// </summary>
    [HttpGet("{userCode}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetUser(string userCode)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();
        var result = await _serviceLayer.GetStringAsync(company, $"Users('{userCode}')");
        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get user" });

        return Content(BuildSuccessResponseFirst(result.Data), "application/json");
    }

    /// <summary>
    /// Opretter en ny SAP-bruger.
    /// </summary>
    [HttpPost]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> CreateUser([FromBody] SapUserCreateDto user)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (string.IsNullOrWhiteSpace(user.USER_CODE))
            return BadRequest(new { success = false, error = "USER_CODE is required" });

        if (string.IsNullOrWhiteSpace(user.PASSWORD))
            return BadRequest(new { success = false, error = "PASSWORD is required" });

        var company = _sapConnection.GetMainCompany();
        var jObject = BuildUserJObject(user);

        _logger.LogInformation("[UsersController] Creating SAP user: {UserCode}", user.USER_CODE);

        var result = await _serviceLayer.PostAsync<string>(company, "Users", jObject);
        if (!result.Success)
        {
            _logger.LogWarning("[UsersController] Failed to create SAP user {UserCode}: {Error}", user.USER_CODE, result.Error);
            return BadRequest(new { success = false, error = result.Error ?? "Failed to create user" });
        }

        _logger.LogInformation("[UsersController] SAP user created: {UserCode}", user.USER_CODE);
        return Ok(new { success = true });
    }

    /// <summary>
    /// Opdaterer en eksisterende SAP-bruger.
    /// </summary>
    [HttpPatch("{userCode}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> UpdateUser(string userCode, [FromBody] SapUserCreateDto user)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();
        var jObject = BuildUserJObject(user);
        // UserCode cannot be updated in SAP, so always remove it from PATCH payload
        jObject.Remove("UserCode");

        // Password is optional on update — only include if provided
        if (string.IsNullOrWhiteSpace(user.PASSWORD))
            jObject.Remove("Password");

        _logger.LogInformation("[UsersController] Updating SAP user: {UserCode}", userCode);

        var result = await _serviceLayer.PatchAsync(company, $"Users('{userCode}')", jObject);
        if (!result.Success)
        {
            _logger.LogWarning("[UsersController] Failed to update SAP user {UserCode}: {Error}", userCode, result.Error);
            return BadRequest(new { success = false, error = result.Error ?? "Failed to update user" });
        }

        _logger.LogInformation("[UsersController] SAP user updated: {UserCode}", userCode);
        return Ok(new { success = true });
    }

    // ─── Shared JObject builder ──────────────────────────────────────────

    private static JObject BuildUserJObject(SapUserCreateDto user)
    {
        var jObject = new JObject();
        if (!string.IsNullOrEmpty(user.USER_CODE)) jObject["UserCode"] = user.USER_CODE.Trim();
        if (!string.IsNullOrEmpty(user.PASSWORD)) jObject["Password"] = user.PASSWORD;
        if (!string.IsNullOrEmpty(user.U_NAME)) jObject["UserName"] = user.U_NAME;
        if (!string.IsNullOrEmpty(user.E_Mail)) jObject["EMail"] = user.E_Mail;
        if (!string.IsNullOrEmpty(user.Department)) jObject["Department"] = user.Department;
        if (!string.IsNullOrEmpty(user.Branch)) jObject["Branch"] = user.Branch;
        if (user.SUPERUSER != null) jObject["Superuser"] = user.SUPERUSER;
        if (user.Locked != null) jObject["Locked"] = user.Locked;
        return jObject;
    }

    // ─── Response helpers ────────────────────────────────────────────────

    private static string BuildSuccessResponse(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse))
            return "{\"success\":true,\"data\":[]}";
        try
        {
            var token = JToken.Parse(sapResponse);
            if (token is JObject obj && obj.TryGetValue("value", out var value))
                return $"{{\"success\":true,\"data\":{value}}}";
            return $"{{\"success\":true,\"data\":{token}}}";
        }
        catch { return "{\"success\":true,\"data\":[]}"; }
    }

    private static string BuildSuccessResponseFirst(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse))
            return "{\"success\":true,\"data\":null}";
        try
        {
            var token = JToken.Parse(sapResponse);
            if (token is JObject obj && obj.TryGetValue("value", out var value) && value is JArray arr && arr.Count > 0)
                return $"{{\"success\":true,\"data\":{arr[0]}}}";
            return $"{{\"success\":true,\"data\":{token}}}";
        }
        catch { return "{\"success\":true,\"data\":null}"; }
    }
}