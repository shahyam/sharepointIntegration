using System.Net;
using System.Net.Http;
using System.Reflection;
using Microsoft.Extensions.Options;
using MSGraph = Microsoft.Graph;
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
builder.Services.AddSingleton<MSGraph.GraphServiceClient>(serviceProvider =>
{
    var graphOptions = serviceProvider.GetRequiredService<IOptions<GraphOptions>>().Value;

    if (string.IsNullOrWhiteSpace(graphOptions.TenantId) ||
        string.IsNullOrWhiteSpace(graphOptions.ClientId) ||
        string.IsNullOrWhiteSpace(graphOptions.ClientSecret))
    {
        throw new InvalidOperationException("Graph configuration is missing. Ensure Graph:TenantId, Graph:ClientId, and Graph:ClientSecret are set.");
    }

    var handler = new HttpClientHandler
    {
        UseProxy = false
    };

    if (!string.IsNullOrWhiteSpace(graphOptions.ProxyUrl) || graphOptions.ProxyUseDefaultCredentials)
    {
        IWebProxy? proxy = null;
        if (!string.IsNullOrWhiteSpace(graphOptions.ProxyUrl))
        {
            proxy = new WebProxy(graphOptions.ProxyUrl);

            if (!string.IsNullOrWhiteSpace(graphOptions.ProxyUsername))
            {
                proxy.Credentials = new NetworkCredential(
                    graphOptions.ProxyUsername,
                    graphOptions.ProxyPassword,
                    graphOptions.ProxyDomain);
            }
            else if (graphOptions.ProxyUseDefaultCredentials)
            {
                proxy.Credentials = CredentialCache.DefaultNetworkCredentials;
            }
        }
        else
        {
            proxy = WebRequest.DefaultWebProxy;
            if (graphOptions.ProxyUseDefaultCredentials)
            {
                proxy.Credentials = CredentialCache.DefaultNetworkCredentials;
            }
        }

        if (proxy != null)
        {
            handler.UseProxy = true;
            handler.Proxy = proxy;
        }
    }

    var msalHttpClient = new HttpClient(handler, disposeHandler: false);

    var confidentialClient = ConfidentialClientApplicationBuilder
        .Create(graphOptions.ClientId)
        .WithTenantId(graphOptions.TenantId)
        .WithClientSecret(graphOptions.ClientSecret)
        .WithHttpClientFactory(new MsalHttpClientFactory(msalHttpClient))
        .Build();

    var scopes = graphOptions.Scopes?.Length > 0
        ? graphOptions.Scopes
        : new[] { "https://graph.microsoft.com/.default" };

    var authProvider = new GraphAuthProvider(confidentialClient, scopes);
    var httpProvider = new MSGraph.HttpProvider(handler, disposeHandler: false);
    return new MSGraph.GraphServiceClient(authProvider, httpProvider);
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

public partial class Program
{
}
