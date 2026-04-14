using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using WebAPI.Tests.Support;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;

namespace WebAPI.Tests.Integration;

public class PlanUnlimitedIntegrationTests
{
    [Fact]
    public async Task RunBtnPu_ReturnsOkAndUsesBtnPuCommand()
    {
        var sap = new FakeSapConnectionService(isEnabled: true);
        var serviceLayer = new FakeServiceLayerService();

        var runner = new FakePlanUnlimitedRunnerService
        {
            RunHandler = (request, _) =>
            {
                return Task.FromResult(new PlanUnlimitedRunResult
                {
                    Success = true,
                    ExitCode = 0,
                    Message = $"Executed {request.Btn}",
                    Output = "done",
                    Error = string.Empty
                });
            }
        };

        await using var factory = new IntegrationTestApplicationFactory(sap, serviceLayer, runner);
        using var client = factory.CreateClient();

        var response = await client.PostAsJsonAsync("/api/plan/order/btnpu", new
        {
            docEntry = 123,
            user = "tester"
        });

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);

        var doc = await response.Content.ReadFromJsonAsync<JsonDocument>();
        Assert.NotNull(doc);

        var root = doc!.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal("Executed BtnPU", root.GetProperty("message").GetString());

        Assert.NotNull(runner.LastRunRequest);
        Assert.Equal("BtnPU", runner.LastRunRequest!.Btn);
        Assert.Equal(123, runner.LastRunRequest.DocEntry);
    }
}
