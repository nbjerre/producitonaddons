namespace WebAPI.Models.SalesProduction;

/// <summary>
/// Cached generated print payload served by follow-up download/view endpoints.
/// </summary>
public sealed class GeneratedPrintDocument
{
    public required string FileName { get; init; }
    public required byte[] Content { get; init; }
    public required string ContentType { get; init; }
    public bool OpenInline { get; init; }
}
