using System.Reflection;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.OpenApi.Models;
using sharepointIntegration.Options;
using sharepointIntegration.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "SharePoint Integration API",
        Version = "v1",
        Description = "API endpoints for Microsoft Graph / SharePoint integration."
    });

    var xmlFileName = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFileName);
    options.IncludeXmlComments(xmlPath);
});

builder.Services.Configure<GraphOptions>(builder.Configuration.GetSection("Graph"));
builder.Services.AddSingleton<GraphServiceClient>(serviceProvider =>
{
    var graphOptions = serviceProvider.GetRequiredService<IOptions<GraphOptions>>().Value;

    if (string.IsNullOrWhiteSpace(graphOptions.TenantId) ||
        string.IsNullOrWhiteSpace(graphOptions.ClientId) ||
        string.IsNullOrWhiteSpace(graphOptions.ClientSecret))
    {
        throw new InvalidOperationException("Graph configuration is missing. Ensure Graph:TenantId, Graph:ClientId, and Graph:ClientSecret are set.");
    }

    var confidentialClient = ConfidentialClientApplicationBuilder
        .Create(graphOptions.ClientId)
        .WithTenantId(graphOptions.TenantId)
        .WithClientSecret(graphOptions.ClientSecret)
        .Build();

    var scopes = graphOptions.Scopes?.Length > 0
        ? graphOptions.Scopes
        : new[] { "https://graph.microsoft.com/.default" };

    var authProvider = new GraphAuthProvider(confidentialClient, scopes);
    return new GraphServiceClient(authProvider);
});

var app = builder.Build();

// Configure the HTTP request pipeline.
app.UseSwagger();
app.UseSwaggerUI(options =>
{
    options.SwaggerEndpoint("/swagger/v1/swagger.json", "SharePoint Integration API v1");
    options.RoutePrefix = "swagger";
});

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
