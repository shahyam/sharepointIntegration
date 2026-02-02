using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using System.Net;

namespace sharepointIntegration.Controllers;

[ApiController]
[Route("[controller]")]
public class GraphController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<GraphController> _logger;

    public GraphController(GraphServiceClient graphClient, ILogger<GraphController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    [HttpGet("me")]
    public async Task<IActionResult> GetMe()
    {
        try
        {
            var me = await _graphClient.Me
                .Request()
                .Select("id,displayName,userPrincipalName")
                .GetAsync();

            return Ok(new
            {
                me.Id,
                me.DisplayName,
                me.UserPrincipalName
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/me");
        }
    }

    [HttpGet("drive/root")]
    public async Task<IActionResult> GetDriveRoot()
    {
        try
        {
            var driveRoot = await _graphClient.Me.Drive.Root
                .Request()
                .Select("id,name,webUrl")
                .GetAsync();

            return Ok(new
            {
                driveRoot.Id,
                driveRoot.Name,
                driveRoot.WebUrl
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/drive/root");
        }
    }

    [HttpGet("users/{userId}")]
    public async Task<IActionResult> GetUser(string userId)
    {
        try
        {
            var user = await _graphClient.Users[userId]
                .Request()
                .Select("id,displayName,userPrincipalName")
                .GetAsync();

            return Ok(new
            {
                user.Id,
                user.DisplayName,
                user.UserPrincipalName
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/users/{userId}");
        }
    }

    [HttpGet("users/{userId}/drive/root")]
    public async Task<IActionResult> GetUserDriveRoot(string userId)
    {
        try
        {
            var driveRoot = await _graphClient.Users[userId].Drive.Root
                .Request()
                .Select("id,name,webUrl")
                .GetAsync();

            return Ok(new
            {
                driveRoot.Id,
                driveRoot.Name,
                driveRoot.WebUrl
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/users/{userId}/drive/root");
        }
    }

    [HttpGet("sites/{siteId}/drives/{driveId}/folders")]
    public async Task<IActionResult> GetSiteDriveFolder(
        string siteId,
        string driveId,
        [FromQuery] string folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            return BadRequest(new { message = "folderPath query parameter is required." });
        }

        try
        {
            var folder = await _graphClient
                .Drives[driveId]
                .Root
                .ItemWithPath(folderPath)
                .Request()
                .Select("id,name,webUrl,folder")
                .GetAsync();

            return Ok(new
            {
                folder.Id,
                folder.Name,
                folder.WebUrl,
                IsFolder = folder.Folder != null
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/sites/{siteId}/drives/{driveId}/folders?folderPath=");
        }
    }

    private IActionResult HandleGraphException(ServiceException ex, string operation)
    {
        _logger.LogError(ex, "Graph request failed for {Operation}. Code: {Code}, Message: {Message}",
            operation,
            ex.Error?.Code ?? "Unknown",
            ex.Error?.Message ?? ex.Message);

        var details = new ProblemDetails
        {
            Title = "Microsoft Graph request failed",
            Detail = ex.Error?.Message ?? ex.Message,
            Status = (int)HttpStatusCode.BadGateway
        };

        details.Extensions["code"] = ex.Error?.Code;
        details.Extensions["operation"] = operation;

        return StatusCode(details.Status.Value, details);
    }
}