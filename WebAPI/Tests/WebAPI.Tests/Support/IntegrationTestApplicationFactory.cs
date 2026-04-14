using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using WebAPI.Configuration;
using WorksheetAPI.Interfaces;

namespace WebAPI.Tests.Support;

public sealed class IntegrationTestApplicationFactory : WebApplicationFactory<Program>
{
    private readonly ISapConnectionService _sapConnectionService;
    private readonly IServiceLayerService _serviceLayerService;
    private readonly IPlanUnlimitedRunnerService _planUnlimitedRunnerService;
    private readonly Action<PrintSettings>? _configurePrintSettings;

    public IntegrationTestApplicationFactory(
        ISapConnectionService sapConnectionService,
        IServiceLayerService serviceLayerService,
        IPlanUnlimitedRunnerService planUnlimitedRunnerService,
        Action<PrintSettings>? configurePrintSettings = null)
    {
        _sapConnectionService = sapConnectionService;
        _serviceLayerService = serviceLayerService;
        _planUnlimitedRunnerService = planUnlimitedRunnerService;
        _configurePrintSettings = configurePrintSettings;
    }

    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.UseEnvironment("Development");

        builder.ConfigureServices(services =>
        {
            services.RemoveAll<ISapConnectionService>();
            services.RemoveAll<IServiceLayerService>();
            services.RemoveAll<IPlanUnlimitedRunnerService>();

            services.AddSingleton(_sapConnectionService);
            services.AddScoped(_ => _serviceLayerService);
            services.AddScoped(_ => _planUnlimitedRunnerService);

            if (_configurePrintSettings != null)
            {
                services.PostConfigure(_configurePrintSettings);
            }
        });
    }
}
