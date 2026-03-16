namespace sharepointIntegration.Options;

public class GraphOptions
{
    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public string[] Scopes { get; set; } = Array.Empty<string>();
    public string? ProxyUrl { get; set; }
    public string? ProxyUsername { get; set; }
    public string? ProxyPassword { get; set; }
    public string? ProxyDomain { get; set; }
    public bool ProxyUseDefaultCredentials { get; set; }
}
