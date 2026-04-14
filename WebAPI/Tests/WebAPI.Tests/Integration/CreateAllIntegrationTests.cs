using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Newtonsoft.Json.Linq;
using WebAPI.Tests.Support;
using WorksheetAPI.Interfaces;

namespace WebAPI.Tests.Integration;

public class CreateAllIntegrationTests
{
    [Fact]
    public async Task CreateAll_WhenSingleValidLine_CreatesProductionOrder()
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
                      "DocEntry": 123,
                      "CardCode": "C0001",
                      "Project": "P-01",
                      "U_RCS_RD": "N",
                      "DocumentLines": [
                        {
                          "LineNum": 0,
                          "VisOrder": 0,
                          "ItemCode": "FG-100",
                          "Quantity": 2,
                          "U_RCS_ONSTO": "Y",
                          "ShipDate": "2026-04-20",
                          "U_RCS_DelDays": 3
                        }
                      ]
                    }
                    """;

                    return Task.FromResult(ServiceLayerResult<string>.Ok(orderJson));
                }

                if (resource == "ProductTrees('FG-100')")
                {
                    const string bomJson = """
                    {
                      "TreeType": "iProductionTree",
                      "ProductTreeLines": []
                    }
                    """;

                    return Task.FromResult(ServiceLayerResult<string>.Ok(bomJson));
                }

                if (resource == "ProductionOrders" && !string.IsNullOrWhiteSpace(filter) && filter.Contains("ItemNo eq 'FG-100'", StringComparison.Ordinal))
                {
                    const string emptyOrders = """
                    {
                      "value": []
                    }
                    """;

                    return Task.FromResult(ServiceLayerResult<string>.Ok(emptyOrders));
                }

                return Task.FromResult(ServiceLayerResult<string>.Fail($"Unexpected call: resource={resource}, filter={filter}"));
            },
            PostJObjectHandler = (_, resource, _) =>
            {
                if (resource == "ProductionOrders")
                {
                    var created = new JObject
                    {
                        ["AbsoluteEntry"] = 1001,
                        ["DocumentNumber"] = 5001
                    };

                    return Task.FromResult(ServiceLayerResult<JObject>.Ok(created));
                }

                return Task.FromResult(ServiceLayerResult<JObject>.Fail($"Unexpected POST resource: {resource}"));
            }
        };

        await using var factory = new IntegrationTestApplicationFactory(sap, serviceLayer, runner);
        using var client = factory.CreateClient();

        var response = await client.PostAsJsonAsync("/api/sales-production/create-all", new
        {
            salesOrderDocEntry = 123,
            confirmSubBoms = true
        });

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);

        var doc = await response.Content.ReadFromJsonAsync<JsonDocument>();
        Assert.NotNull(doc);

        var root = doc!.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("created").GetBoolean());
        Assert.Equal(1, root.GetProperty("createdCount").GetInt32());
        Assert.Equal(0, root.GetProperty("failedCount").GetInt32());
    }
}
