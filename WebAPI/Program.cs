using System.Reflection;
using Microsoft.OpenApi;
using WorksheetAPI.Configuration;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Services;

using Serilog;

using WebAPI.Configuration;


var builder = WebApplication.CreateBuilder(args);

// Serilog konfiguration: log til fil
builder.Host.UseSerilog((ctx, lc) => lc
    .WriteTo.File("logs/log.txt", rollingInterval: RollingInterval.Day)
    .ReadFrom.Configuration(ctx.Configuration)
);

// Add configuration
builder.Services.Configure<SapSettings>(builder.Configuration.GetSection(SapSettings.SectionName));
builder.Services.Configure<PlanUnlimitedSettings>(builder.Configuration.GetSection(PlanUnlimitedSettings.SectionName));
builder.Services.Configure<PrintSettings>(builder.Configuration.GetSection(PrintSettings.SectionName));

// Add services
builder.Services.AddSingleton<ISapConnectionService, SapConnectionService>();
builder.Services.AddScoped<IServiceLayerService, ServiceLayerService>();
builder.Services.AddScoped<IBomService, BomService>();
builder.Services.AddScoped<IProductionOrderHierarchyService, ProductionOrderHierarchyService>();
builder.Services.AddScoped<ISalesProductionPrintService, SalesProductionPrintService>();
builder.Services.AddScoped<IPlanUnlimitedRunnerService, PlanUnlimitedRunnerService>();
builder.Services.AddMemoryCache();

// HttpClient til Crystal Reports - SSL-validering er slået fra fordi Crystal-serveren bruger self-signed cert
builder.Services.AddHttpClient("crystal")
    .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler
    {
        ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
    });

// Add controllers with text/plain input formatter
builder.Services.AddControllers(options =>
{
    options.InputFormatters.Insert(0, new TextPlainInputFormatter());
});

builder.Services.AddEndpointsApiExplorer();

builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "WebAPI",
        Version = "v1",
        Description = "API til SAP Business One integration, produktionsflows, planlægning og worksheets."
    });

    var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    if (File.Exists(xmlPath))
    {
        options.IncludeXmlComments(xmlPath, includeControllerXmlComments: true);
    }

    options.CustomSchemaIds(type => type.FullName?.Replace("+", ".") ?? type.Name);
});

// Add CORS (configure as needed for your frontend)
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
      .AllowAnyMethod()
      .AllowAnyHeader();
    });
});

var app = builder.Build();

// Log alle indgående HTTP-kald centralt (metode, path, statuskode og varighed)
app.UseSerilogRequestLogging(options =>
{
    options.MessageTemplate = "HTTP {RequestMethod} {RequestPath} responded {StatusCode} in {Elapsed:0.0000} ms";
});

// Configure the HTTP request pipeline
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(options =>
    {
        options.SwaggerEndpoint("/swagger/v1/swagger.json", "WebAPI v1");
        options.RoutePrefix = "swagger";
        options.DocumentTitle = "WebAPI Swagger";
    });
}

// Use CORS before other middleware - allow cross-origin requests from UI5 frontend
app.UseCors("AllowAll");

// Only redirect to HTTPS in production (not in development to avoid CORS issues)
if (!app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
}

app.UseAuthorization();

app.MapControllers();

// Log startup info
Log.Information("WebAPI started at {StartedAtUtc}", DateTime.UtcNow);

app.Run();

public partial class Program;
