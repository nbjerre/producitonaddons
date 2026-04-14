using WorksheetAPI.Models;

namespace WorksheetAPI.Interfaces;

/// <summary>
/// Encapsulates BOM-specific business operations used by production flows.
/// </summary>
public interface IBomService
{
    Task<bool> HasProductionBomAsync(SapCompany company, string itemCode);

    Task<List<string>> GetSubBomCodesAsync(SapCompany company, string rootItemCode);

    Task<List<SubBomSnapshot>> ApplyTemporarySubBomUpdatesAsync(
        SapCompany company,
        List<string> subBomCodes,
        List<SubBomAdjustmentDto>? adjustments);

    Task RestoreTemporarySubBomUpdatesAsync(SapCompany company, List<SubBomSnapshot> snapshots);
}
