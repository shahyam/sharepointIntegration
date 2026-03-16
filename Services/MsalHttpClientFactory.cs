using System.Net.Http;
using Microsoft.Identity.Client;

namespace sharepointIntegration.Services;

public sealed class MsalHttpClientFactory : IMsalHttpClientFactory
{
    private readonly HttpClient _client;

    public MsalHttpClientFactory(HttpClient client)
    {
        _client = client;
    }

    public HttpClient GetHttpClient() => _client;
}
