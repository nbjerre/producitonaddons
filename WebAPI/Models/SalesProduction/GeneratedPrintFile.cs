namespace WebAPI.Models.SalesProduction;

/// <summary>
/// Represents one generated file before final merge/packaging.
/// </summary>
public sealed class GeneratedPrintFile
{
    public required string FileName { get; init; }
    public required byte[] Content { get; init; }
    public string ContentType { get; init; } = "application/pdf";
}
