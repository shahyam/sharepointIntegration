using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using System.Net;
using System.IO;
using System.Net.Http;
using System.Threading;

namespace sharepointIntegration.Controllers;

[ApiController]
[Route("[controller]")]
public class GraphController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<GraphController> _logger;
    private readonly string userId = "shyamendrashah@zirconsam.onmicrosoft.com";
    private static readonly long MaxUploadBytes = 50 * 1024 * 1024;
    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".doc",
        ".docx",
        ".xls",
        ".xlsx",
        ".pdf",
        ".eml",
        ".msg"
    };
    private static readonly HashSet<string> AllowedRoles = new(StringComparer.OrdinalIgnoreCase)
    {
        "read",
        "write"
    };

    public GraphController(GraphServiceClient graphClient, ILogger<GraphController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

  
  

    [HttpGet("users")]
    /// <summary>Gets basic user info for the configured userId.</summary>
    /// <remarks>No parameters.</remarks>
    public async Task<IActionResult> GetUser()
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


    

    [HttpGet("users/drive/root")]
    /// <summary>Gets the root drive for the configured userId.</summary>
    /// <remarks>No parameters.</remarks>
    public async Task<IActionResult> GetUserDriveRoot()
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

    [HttpGet("users/drive/root/folders")]
    /// <summary>Lists folders under the user's drive root.</summary>
    /// <remarks>No parameters.</remarks>
    public async Task<IActionResult> GetUserDriveRootFolders()
    {
        try
        {
            var items = await _graphClient.Users[userId].Drive.Root.Children
                .Request()
                .Select("id,name,webUrl,folder")
                .GetAsync();

            var folders = items.CurrentPage
                .Where(item => item.Folder != null)
                .Select(item => new
                {
                    item.Id,
                    item.Name,
                    item.WebUrl
                });

            return Ok(folders);
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/users/drive/root/folders");
        }
    }

    [HttpPost("users/drive/root/folders")]
    /// <summary>Creates a folder under the user's drive root.</summary>
    /// <param name="folderName">Query parameter. Required.</param>
    public async Task<IActionResult> CreateUserDriveRootFolder([FromQuery] string folderName)
    {
        if (string.IsNullOrWhiteSpace(folderName))
        {
            return BadRequest(new { message = "folderName query parameter is required." });
        }

        try
        {
            var driveItem = new DriveItem
            {
                Name = folderName,
                Folder = new Folder(),
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = "rename"
                }
            };

            var createdFolder = await _graphClient.Users[userId].Drive.Root.Children
                .Request()
                .AddAsync(driveItem);

            return Ok(new
            {
                createdFolder.Id,
                createdFolder.Name,
                createdFolder.WebUrl
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "POST /graph/users/drive/root/folders?folderName=");
        }
    }

    [HttpPost("users/drive/root/folders/upload")]
    [RequestSizeLimit(50 * 1024 * 1024)]
    /// <summary>Uploads a file to a folder path under the user's drive root.</summary>
    /// <param name="folderPath">Query parameter. Required.</param>
    /// <param name="file">Form file. Required.</param>
    /// <remarks>Limits: Max 50 MB; allowed extensions are doc, docx, xls, xlsx, pdf, eml, msg.</remarks>
    public async Task<IActionResult> UploadFileToFolder([FromQuery] string folderPath, IFormFile file)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            return BadRequest(new { message = "folderPath query parameter is required." });
        }

        if (file == null || file.Length == 0)
        {
            return BadRequest(new { message = "File is required." });
        }

        if (file.Length > MaxUploadBytes)
        {
            return BadRequest(new { message = "File size exceeds 50 MB limit." });
        }

        var extension = Path.GetExtension(file.FileName);
        if (string.IsNullOrWhiteSpace(extension) || !AllowedExtensions.Contains(extension))
        {
            return BadRequest(new
            {
                message = "Only Word, Excel, PDF, or email files are allowed.",
                allowedExtensions = AllowedExtensions
            });
        }

        try
        {
            var trimmedFolderPath = folderPath.Trim('/');
            var targetPath = string.IsNullOrWhiteSpace(trimmedFolderPath)
                ? file.FileName
                : $"{trimmedFolderPath}/{file.FileName}";

            await using var stream = file.OpenReadStream();
            var uploadedItem = await _graphClient.Users[userId].Drive.Root
                .ItemWithPath(targetPath)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);

            return Ok(new
            {
                uploadedItem.Id,
                uploadedItem.Name,
                uploadedItem.WebUrl,
                uploadedItem.Size
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "POST /graph/users/drive/root/folders/upload?folderPath=");
        }
    }

    [HttpPost("users/drive/root/folders/secure")]
    /// <summary>Creates a folder (or uses existing) and applies sharing permissions.</summary>
    /// <param name="request">Body. FolderName required, FolderPath optional, Recipients required, Role optional (default "read"),
    /// RemoveExistingPermissions optional (default true), Message optional.</param>
    public async Task<IActionResult> 
        CreateSecureFolder([FromBody] SecureFolderRequest request)
    {
        if (request == null)
        {

            return BadRequest(new { message = "Request body is required." });
        }

        if (string.IsNullOrWhiteSpace(request.FolderName))
        {
            return BadRequest(new { message = "FolderName is required." });
        }

        if (request.Recipients == null || request.Recipients.Count == 0)
        {
            return BadRequest(new { message = "At least one recipient is required." });
        }

        var role = string.IsNullOrWhiteSpace(request.Role) ? "read" : request.Role.Trim();
        if (!AllowedRoles.Contains(role))
        {
            return BadRequest(new { message = "Role must be 'read' or 'write'." });
        }

        try
        {
            var trimmedFolderPath = request.FolderPath?.Trim('/');
            var targetPath = string.IsNullOrWhiteSpace(trimmedFolderPath)
                ? request.FolderName
                : $"{trimmedFolderPath}/{request.FolderName}";

            // Check if folder already exists
            DriveItem? createdFolder = null;
            try
            {
                createdFolder = await _graphClient.Users[userId].Drive.Root
                    .ItemWithPath(targetPath)
                    .Request()
                    .GetAsync();
                
                // If folder doesn't exist or is not a folder, we'll create it
                if (createdFolder?.Folder == null)
                {
                    createdFolder = null;
                }
            }
            catch (ServiceException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                // Folder doesn't exist, we'll create it below
                createdFolder = null;
            }

            // Create folder only if it doesn't exist
            if (createdFolder == null)
            {
                var driveItem = new DriveItem
                {
                    Name = request.FolderName,
                    Folder = new Folder(),
                    AdditionalData = new Dictionary<string, object>
                    {
                        ["@microsoft.graph.conflictBehavior"] = "fail"
                    }
                };

                if (string.IsNullOrWhiteSpace(trimmedFolderPath))
                {
                    createdFolder = await _graphClient.Users[userId].Drive.Root
                        .Children
                        .Request()
                        .AddAsync(driveItem);
                }
                else
                {
                    createdFolder = await _graphClient.Users[userId].Drive.Root
                        .ItemWithPath(trimmedFolderPath)
                        .Children
                        .Request()
                        .AddAsync(driveItem);
                }
            }

            var recipients = request.Recipients
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .Select(email => new DriveRecipient { Email = email.Trim() })
                .ToList();

            if (recipients.Count == 0)
            {
                return BadRequest(new { message = "At least one valid recipient email is required." });
            }

            // Remove unwanted permissions BEFORE adding new ones
            if (request.RemoveExistingPermissions)
            {
                var allowedUserIds = await ResolveRecipientIdsAsync(recipients.Select(r => r.Email));
                // Add the owner's ID to the allowed list
                var owner = await _graphClient.Users[userId]
                    .Request()
                    .Select("id")
                    .GetAsync();
                if (!string.IsNullOrWhiteSpace(owner.Id))
                {
                    allowedUserIds.Add(owner.Id);
                }
                await RemoveUnwantedPermissions(createdFolder.Id, allowedUserIds);
            }

            // Now add the new permissions
            var inviteRoles = new List<string> { role };
            await SendInviteAsync(createdFolder.Id, recipients, inviteRoles, request.Message);

            return Ok(new
            {
                createdFolder.Id,
                createdFolder.Name,
                createdFolder.WebUrl
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "POST /graph/users/drive/root/folders/secure");
        }
    }

    [HttpPost("users/drive/items/{itemId}/permissions")]
    /// <summary>Adds sharing permissions to an existing drive item.</summary>
    /// <param name="itemId">Route parameter. Required.</param>
    /// <param name="request">Body. Recipients required, Role optional (default "read"), Message optional.</param>
    public async Task<IActionResult> AddFolderPermissions(
        string itemId,
        [FromBody] AddPermissionsRequest request)
    {
        if (string.IsNullOrWhiteSpace(itemId))
        {
            return BadRequest(new { message = "itemId is required." });
        }

        if (request == null)
        {
            return BadRequest(new { message = "Request body is required." });
        }

        if (request.Recipients == null || request.Recipients.Count == 0)
        {
            return BadRequest(new { message = "At least one recipient is required." });
        }

        var role = string.IsNullOrWhiteSpace(request.Role) ? "read" : request.Role.Trim();
        if (!AllowedRoles.Contains(role))
        {
            return BadRequest(new { message = "Role must be 'read' or 'write'." });
        }

        try
        {
            var recipients = request.Recipients
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .Select(email => new DriveRecipient { Email = email.Trim() })
                .ToList();

            if (recipients.Count == 0)
            {
                return BadRequest(new { message = "At least one valid recipient email is required." });
            }

            var inviteRoles = new List<string> { role };
            await SendInviteAsync(itemId, recipients, inviteRoles, request.Message);

            return Ok(new
            {
                message = "Permissions added successfully",
                itemId,
                recipientsCount = recipients.Count,
                role
            });
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "POST /graph/users/drive/items/{itemId}/permissions");
        }
    }

    [HttpGet("users/drive/items/download")]
    /// <summary>Searches by name and downloads the first match. If the match is a folder, returns a zip.</summary>
    /// <param name="q">Query parameter. Required.</param>
    /// <remarks>No defaults; search uses Microsoft Graph search semantics.</remarks>
    public async Task<IActionResult> DownloadDriveItem([FromQuery] string q)
    {
        if (string.IsNullOrWhiteSpace(q))
        {
            return BadRequest(new { message = "q query parameter is required." });
        }

        try
        {
            var results = await _graphClient.Users[userId].Drive.Root
                .Search(q)
                .Request()
                .Select("id,name,webUrl,file,folder,parentReference,remoteItem")
                .GetAsync();

            var firstMatch = results.CurrentPage.FirstOrDefault();
            if (firstMatch == null)
            {
                return NotFound(new { message = "No matching item found for the provided query." });
            }

            var effectiveItemId = firstMatch.RemoteItem?.Id ?? firstMatch.Id;
            var effectiveParent = firstMatch.RemoteItem?.ParentReference ?? firstMatch.ParentReference;
            var driveId = effectiveParent?.DriveId;
            var driveItems = string.IsNullOrWhiteSpace(driveId)
                ? _graphClient.Users[userId].Drive.Items
                : _graphClient.Drives[driveId].Items;

            var driveItem = await driveItems[effectiveItemId]
                .Request()
                .Select("id,name,file,folder")
                .GetAsync();

            if (driveItem.Folder != null)
            {
                var options = new List<Option> { new QueryOption("format", "zip") };
                var zipStream = await driveItems[effectiveItemId]
                    .Content
                    .Request(options)
                    .GetAsync();

                return File(zipStream, "application/zip", $"{driveItem.Name}.zip");
            }

            var contentStream = await driveItems[effectiveItemId]
                .Content
                .Request()
                .GetAsync();

            var contentType = driveItem.File?.MimeType ?? "application/octet-stream";

            return File(contentStream, contentType, driveItem.Name);
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/users/drive/items/download?q=");
        }
    }

    [HttpGet("users/drive/search")]
    /// <summary>Searches the user's drive and returns matching file item IDs.</summary>
    /// <param name="q">Query parameter. Required.</param>
    /// <remarks>No defaults; search uses Microsoft Graph search semantics.</remarks>
    public async Task<IActionResult> SearchDriveItems([FromQuery] string q)
    {
        if (string.IsNullOrWhiteSpace(q))
        {
            return BadRequest(new { message = "q query parameter is required." });
        }

        try
        {
            var results = await _graphClient.Users[userId].Drive.Root
                .Search(q)
                .Request()
                .Select("id,name,webUrl,file,folder")
                .GetAsync();

            var items = results.CurrentPage
                .Where(item => item.File != null)
                .Select(item => new
                {
                    item.Id,
                    item.Name,
                    item.WebUrl
                });

            return Ok(items);
        }
        catch (ServiceException ex)
        {
            return HandleGraphException(ex, "GET /graph/users/drive/search?q=");
        }
    }

    private async Task<HashSet<string>> ResolveRecipientIdsAsync(IEnumerable<string?> emails)
    {
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var email in emails.Where(email => !string.IsNullOrWhiteSpace(email)))
        {
            var user = await _graphClient.Users[email!]
                .Request()
                .Select("id")
                .GetAsync();

            if (!string.IsNullOrWhiteSpace(user.Id))
            {
                ids.Add(user.Id);
            }
        }

        return ids;
    }

    private async Task SendInviteAsync(string itemId, IEnumerable<DriveRecipient> recipients, IEnumerable<string> roles, string? message)
    {
        var requestUrl = $"{_graphClient.BaseUrl}/users/{userId}/drive/items/{itemId}/invite";
        var inviteRequest = new BaseRequest(requestUrl, _graphClient, null)
        {
            Method = Microsoft.Graph.HttpMethods.POST,
            ContentType = "application/json"
        };

        var payload = new
        {
            recipients = recipients.Select(r => new { email = r.Email }).ToList(),
            requireSignIn = true,
            sendInvitation = true,
            roles = roles.ToList(),
            message
        };

        await inviteRequest.SendAsync<Permission>(payload, CancellationToken.None, HttpCompletionOption.ResponseContentRead);
    }

    private async Task RemoveUnwantedPermissions(string itemId, HashSet<string> allowedUserIds)
    {
        var permissions = await _graphClient.Users[userId].Drive.Items[itemId]
            .Permissions
            .Request()
            .GetAsync();

        foreach (var permission in permissions.CurrentPage)
        {
            if (permission.Roles != null && permission.Roles.Contains("owner"))
            {
                continue;
            }

            var grantedUserId = permission.GrantedToV2?.User?.Id ?? permission.GrantedTo?.User?.Id;

            if (string.IsNullOrWhiteSpace(grantedUserId) || !allowedUserIds.Contains(grantedUserId))
            {
                await _graphClient.Users[userId].Drive.Items[itemId]
                    .Permissions[permission.Id]
                    .Request()
                    .DeleteAsync();
            }
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

    public class SecureFolderRequest
    {
        public string FolderName { get; set; } = string.Empty;
        public string? FolderPath { get; set; }
        public List<string> Recipients { get; set; } = new();
        public string? Role { get; set; }
        public bool RemoveExistingPermissions { get; set; } = true;
        public string? Message { get; set; }
    }

    public class AddPermissionsRequest
    {
        public List<string> Recipients { get; set; } = new();
        public string? Role { get; set; }
        public string? Message { get; set; }
    }
}