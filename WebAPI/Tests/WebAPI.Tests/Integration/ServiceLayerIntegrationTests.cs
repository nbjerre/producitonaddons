using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using WebAPI.Tests.Support;
using WorksheetAPI.Interfaces;

namespace WebAPI.Tests.Integration;

public class ServiceLayerIntegrationTests
{
    [Fact]
    public async Task GetCopyCandidates_ReturnsItemsFromServiceLayer()
    {
        var sap = new FakeSapConnectionService(isEnabled: true);
        var serviceLayer = new FakeServiceLayerService
        {
            GetStringHandler = (_, resource, _, _, _) =>
            {
                if (resource == "Items")
                {
                    const string payload = """
                    {
                      "value": [
                        { "ItemCode": "FG-100", "ItemName": "Finished good" }
                      ]
                    }
                    """;

                    return Task.FromResult(ServiceLayerResult<string>.Ok(payload));
                }

                return Task.FromResult(ServiceLayerResult<string>.Fail($"Unexpected resource: {resource}"));
            }
        };

        var runner = new FakePlanUnlimitedRunnerService();

        await using var factory = new IntegrationTestApplicationFactory(sap, serviceLayer, runner);
        using var client = factory.CreateClient();

        var response = await client.GetAsync("/api/bom/copy-candidates");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);

        var doc = await response.Content.ReadFromJsonAsync<JsonDocument>();
        Assert.NotNull(doc);

        var root = doc!.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal("FG-100", root.GetProperty("items")[0].GetProperty("itemCode").GetString());
    }
}
