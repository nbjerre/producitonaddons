using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using WebAPI.Tests.Support;
using WorksheetAPI.Interfaces;

namespace WebAPI.Tests.Integration;

public class CrystalPrintIntegrationTests
{
  [Fact]
  public async Task PrintProductionOrders_WhenNoPrintableLines_ReturnsZero()
  {
    var sap = new FakeSapConnectionService(isEnabled: true);
    var runner = new FakePlanUnlimitedRunnerService();

    var serviceLayer = new FakeServiceLayerService
    {
      GetStringHandler = (_, resource, _, _, _) =>
      {
        if (resource.StartsWith("Orders(", StringComparison.Ordinal))
        {
          const string orderJson = """
          {
            "DocNum": 777,
            "CardCode": "C0001",
            "CardName": "Test Customer",
            "DocumentLines": [
            { "LineNum": 0, "U_RCS_ONSTO": "N" }
            ]
          }
          """;

          return Task.FromResult(ServiceLayerResult<string>.Ok(orderJson));
        }

        return Task.FromResult(ServiceLayerResult<string>.Fail($"Unexpected call: resource={resource}"));
      }
    };

    await using var factory = new IntegrationTestApplicationFactory(
      sap,
      serviceLayer,
      runner,
      print =>
      {
        print.CrystalBaseUrl = "https://crystal.test.local";
        print.ProductionReportCode = "WOR10003";
      });

    using var client = factory.CreateClient();

    var response = await client.PostAsJsonAsync("/api/sales-production/print-production-orders", new
    {
      salesOrderDocEntry = 123
    });

    Assert.Equal(HttpStatusCode.OK, response.StatusCode);

    var doc = await response.Content.ReadFromJsonAsync<JsonDocument>();
    Assert.NotNull(doc);

    var root = doc!.RootElement;
    Assert.True(root.GetProperty("success").GetBoolean());
    Assert.Equal(0, root.GetProperty("printableCount").GetInt32());
  }

    [Fact]
    public async Task PrintProductionOrders_WhenCrystalBaseUrlMissing_ReturnsConfiguredCrystalError()
    {
        var sap = new FakeSapConnectionService(isEnabled: true);
        var runner = new FakePlanUnlimitedRunnerService();

        var serviceLayer = new FakeServiceLayerService
        {
            GetStringHandler = (_, resource, _, filter, _) =>
            {
                if (resource.StartsWith("Orders(", StringComparison.Ordinal))
                {
                    const string orderJson = """
                    {
                      "DocNum": 777,
                      "CardCode": "C0001",
                      "CardName": "Test Customer",
                      "DocumentLines": [
                        { "LineNum": 0, "U_RCS_ONSTO": "Y" }
                      ]
                    }
                    """;

                    return Task.FromResult(ServiceLayerResult<string>.Ok(orderJson));
                }

                if (resource == "ProductionOrders" && !string.IsNullOrWhiteSpace(filter) && filter.Contains("ProductionOrderOriginEntry eq 123", StringComparison.Ordinal))
                {
                    const string ordersJson = """
                    {
                      "value": [
                        {
                          "AbsoluteEntry": 1001,
                          "DocumentNumber": 5001,
                          "ItemNo": "ITEM-01",
                          "U_RCS_OL": 0,
                          "ProductionOrderStatus": "boposReleased",
                          "ProductionOrderOriginEntry": 123,
                          "ProductionOrderOriginNumber": 777
                        }
                      ]
                    }
                    """;

                    return Task.FromResult(ServiceLayerResult<string>.Ok(ordersJson));
                }

                return Task.FromResult(ServiceLayerResult<string>.Fail($"Unexpected call: resource={resource}, filter={filter}"));
            }
        };

        await using var factory = new IntegrationTestApplicationFactory(
            sap,
            serviceLayer,
            runner,
            print =>
            {
                print.CrystalBaseUrl = string.Empty;
                print.ProductionReportCode = "WOR10003";
            });

        using var client = factory.CreateClient();

        var response = await client.PostAsJsonAsync("/api/sales-production/print-production-orders", new
        {
            salesOrderDocEntry = 123
        });

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);

        var doc = await response.Content.ReadFromJsonAsync<JsonDocument>();
        Assert.NotNull(doc);

        var root = doc!.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("crystal", root.GetProperty("renderMode").GetString());

        var message = root.GetProperty("message").GetString();
        Assert.Contains("Crystal-print er påkrævet", message, StringComparison.OrdinalIgnoreCase);
    }
}
