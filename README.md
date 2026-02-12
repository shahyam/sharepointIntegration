# SharePoint Integration API (.NET 10)

This API integrates with Microsoft Graph using **Microsoft.Graph SDK v4.54** and client-credential authentication for SharePoint/OneDrive operations.

## Technologies

- **.NET 10.0**
- **Microsoft.Graph 4.54.0**
- **Microsoft.Identity.Client 4.66.0**
- **ASP.NET Core Web API**
- **Swagger/OpenAPI**

## Configure

Update the `Graph` section in `appsettings.json` or `appsettings.Development.json`:

```json
{
  "Graph": {
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret",
    "Scopes": ["https://graph.microsoft.com/.default"]
  }
}
```

## Azure App Registration Steps

1. Go to **Azure Portal** → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name the app (e.g., `SharePointIntegrationApi`) and select **Single tenant**
3. Click **Register**
4. Copy **Application (client) ID** and **Directory (tenant) ID** from the overview page
5. Create a secret: **Certificates & secrets** → **New client secret** → copy the secret value
6. Add permissions: **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
7. Add required permissions:
   - `User.Read.All` - Read all users' profiles
   - `Sites.ReadWrite.All` - Read and write to all SharePoint sites
   - `Files.ReadWrite.All` - Read and write files in all site collections
8. Click **Grant admin consent for [your tenant]**

## SharePoint Online License Requirements

Drive endpoints require SharePoint Online licensing:

1. In the **Microsoft 365 admin center**, go to **Billing → Licenses** and verify a plan that includes **SharePoint Online**
2. Assign the license to at least one user (the user configured in `userId`)
3. Ensure the user's **OneDrive** is provisioned (first sign-in or provision via admin center)
4. Confirm app permissions include `Sites.ReadWrite.All` and `Files.ReadWrite.All` with admin consent

## API Endpoints

### User Information
- **GET** `/graph/users` - Get the configured user's profile (ID, display name, UPN)

### Drive Root Operations
- **GET** `/graph/users/drive/root` - Get the user's OneDrive root folder information
- **GET** `/graph/users/drive/root/folders` - List all folders in the user's OneDrive root

### Folder Management
- **POST** `/graph/users/drive/root/folders?folderName={name}` - Create a new folder in the user's OneDrive root
  - Automatically renames if folder exists
  - Returns folder ID, name, and web URL

### File Upload
- **POST** `/graph/users/drive/root/folders/upload?folderPath={path}` - Upload a file to a specific folder
  - **Query Parameters:**
    - `folderPath`: Target folder path (e.g., "Reports/2024")
  - **Body:** Form-data with file
  - **Constraints:**
    - Max file size: 50 MB
    - Allowed extensions: .doc, .docx, .xls, .xlsx, .pdf, .eml, .msg
  - Returns uploaded file ID, name, web URL, and size

### Secure Folder (Advanced Permission Management)
- **POST** `/graph/users/drive/root/folders/secure` - Create a folder with restricted access
  - **Request Body:**
    ```json
    {
      "folderName": "Confidential Reports",
      "folderPath": "Projects/2024",
      "recipients": ["user1@domain.com", "user2@domain.com"],
      "role": "read",
      "removeExistingPermissions": true,
      "message": "You have been granted access to this folder"
    }
    ```
  - **Parameters:**
    - `folderName` (required): Name of the folder to create
    - `folderPath` (optional): Parent path where folder will be created
    - `recipients` (required): Array of email addresses to grant access to
    - `role` (optional): Either "read" or "write" (defaults to "read")
    - `removeExistingPermissions` (optional): Remove all non-owner permissions before adding new ones (defaults to true)
    - `message` (optional): Custom message to include with the permission grant

### Permission Management
- **POST** `/graph/users/drive/items/{itemId}/permissions` - Add permissions to an existing folder/file
  - **Path Parameters:**
    - `itemId`: The ID of the folder or file
  - **Request Body:**
    ```json
    {
      "recipients": ["user@domain.com"],
      "role": "read",
      "message": "Optional message"
    }
    ```
  - **Parameters:**
    - `recipients` (required): Array of email addresses
    - `role` (optional): "read" or "write" (defaults to "read")
    - `message` (optional): Custom message

## How Secure Folder Works

The secure folder endpoint (`POST /graph/users/drive/root/folders/secure`) provides advanced permission management:

### Process Flow

1. **Check for Existing Folder**
   - First checks if the folder already exists at the specified path
   - If it exists and is a folder, uses the existing folder
   - If it doesn't exist, creates a new folder with conflict behavior set to "fail" to prevent duplicates

2. **Remove Existing Permissions** (if `removeExistingPermissions` is true)
   - Lists all current permissions on the folder
   - Identifies which users are in the allowed recipients list
   - Adds the folder owner to the allowed list to prevent removing their access
   - Removes all permissions that don't belong to the owner or specified recipients
   - This ensures ONLY the specified users have access

3. **Grant New Permissions**
   - Sends invitations to all specified recipients
   - Grants either "read" or "write" access based on the role parameter
   - Uses Microsoft Graph's invite API to properly set permissions
   - Does not send email notifications (`sendInvitation: false`)

### Permission Levels

- **read**: Recipients can view and download files
- **write**: Recipients can view, download, upload, and modify files

### Security Features

- **Explicit Permission Control**: When `removeExistingPermissions` is true, only specified recipients and the owner retain access
- **No Duplicate Folders**: Uses conflict behavior "fail" to prevent creating duplicate folders
- **Owner Protection**: Automatically preserves the folder owner's permissions
- **Input Validation**: Validates email addresses and roles before processing
- **Error Handling**: Provides detailed error messages for troubleshooting

### Example Usage

Create a folder only accessible by specific users:
```bash
POST /graph/users/drive/root/folders/secure
Content-Type: application/json

{
  "folderName": "Q4 Financial Reports",
  "recipients": ["cfo@company.com", "accountant@company.com"],
  "role": "read",
  "removeExistingPermissions": true
}
```

### Manual Alternative

If you prefer to manage permissions manually using Graph API:

1. **Create the folder**: `POST /users/{userId}/drive/root/children`
2. **List current permissions**: `GET /users/{userId}/drive/items/{itemId}/permissions`
3. **Remove unwanted permissions**: `DELETE /users/{userId}/drive/items/{itemId}/permissions/{permissionId}`
4. **Invite specific users**: `POST /users/{userId}/drive/items/{itemId}/invite`

## Required Microsoft Graph Permissions

**Application Permissions** (requires admin consent):
- `User.Read.All` - Read user profiles to resolve email addresses
- `Files.ReadWrite.All` - Full read/write access to files
- `Sites.ReadWrite.All` - Full read/write access to SharePoint sites

## Swagger Documentation

Swagger UI is available at `/swagger` when the application is running. Provides interactive API documentation and testing interface.

## Running the Application

```bash
# Restore packages
dotnet restore

# Build the project
dotnet build

# Run the application
dotnet run
```

The API will be available at `https://localhost:5001` or `http://localhost:5000`.

## Notes

- Uses confidential client with client credentials flow (app-only authentication)
- All operations are performed in the context of the configured user
- Requires proper Azure AD app permissions with admin consent
- OneDrive must be provisioned for the configured user
- File size limit for uploads is 50 MB (configurable in code)
