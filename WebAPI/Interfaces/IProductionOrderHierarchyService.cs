using WorksheetAPI.Models;
using WebAPI.Models.SalesProduction;

namespace WorksheetAPI.Interfaces;

/// <summary>
/// Handles hierarchical production-order creation logic for sub-BOM structures.
/// </summary>
public interface IProductionOrderHierarchyService
{
    Task RemoveDeliveryFromProductionOrderAsync(SapCompany company, int docEntry);

    Task CreateSubProductionOrdersForSubBomAsync(SapCompany company, CreateSubProductionOrdersRequest request);
}
