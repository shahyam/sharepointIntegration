using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace sharepointIntegration.Services;

public class GraphAuthProvider : IAuthenticationProvider
{
    private readonly IConfidentialClientApplication _confidentialClient;
    private readonly string[] _scopes;

    public GraphAuthProvider(IConfidentialClientApplication confidentialClient, string[] scopes)
    {
        _confidentialClient = confidentialClient;
        _scopes = scopes;
    }

    public async Task AuthenticateRequestAsync(HttpRequestMessage request)
    {
        var authResult = await _confidentialClient
            .AcquireTokenForClient(_scopes)
            .ExecuteAsync();

        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
    }
}
