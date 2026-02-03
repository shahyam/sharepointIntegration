# SharePoint Integration (Microsoft Graph SDK v1)

This API sample integrates with Microsoft Graph using the **Microsoft.Graph SDK v1.0** and client-credential authentication.

## Configure

Update the `Graph` section in `appsettings.json` or `appsettings.Development.json`:

- `TenantId`
- `ClientId`
- `ClientSecret`
- `Scopes` (defaults to `https://graph.microsoft.com/.default`)

## Azure app registration steps

1. Go to **Azure Portal** → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name the app (e.g., `SharePointIntegrationApi`) and select **Single tenant** (typical).
3. Click **Register**.
4. Copy **Application (client) ID** and **Directory (tenant) ID** from the overview page.
5. Create a secret: **Certificates & secrets** → **New client secret** → copy the secret value.
6. Add permissions: **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.
7. Add required permissions (e.g., `User.Read.All`, `Sites.Read.All`, `Files.Read.All`).
8. Click **Grant admin consent**.

> Note: `/graph/me` requires delegated permissions. With client-credentials flow, use `/users/{id}` or `/graph/drive/root` instead, or switch to delegated auth if you need a signed-in user.

## SharePoint Online (SPO) license steps

Drive endpoints require SharePoint Online licensing. If you see "Tenant does not have a SPO license", complete the steps below:

1. In the **Microsoft 365 admin center**, go to **Billing → Licenses** and verify a plan that includes **SharePoint Online**.
2. Assign the license to at least one user (the user you call `/graph/users/{userId}` for).
3. Ensure the user's **OneDrive** is provisioned (first sign-in or use the admin center to provision).
4. Confirm app permissions include `Sites.Read.All` and `Files.Read.All` (or higher) and admin consent is granted.

After that, `/graph/users/{userId}/drive/root` should succeed for the licensed user.

## Endpoints

- `GET /graph/me` – returns the signed-in application context user profile.
- `GET /graph/drive/root` – returns the root drive for the user.
- `GET /graph/users/{userId}` – returns a user by ID or UPN (app-only compatible).
- `GET /graph/users/{userId}/drive/root` – returns a user drive root (app-only compatible).
- `GET /graph/users/drive/root/folders` – lists folders under the configured user's drive root.
- `POST /graph/users/drive/root/folders?folderName=Reports` – creates a folder under the configured user's drive root.
- `POST /graph/users/drive/root/folders/upload?folderPath=Shared Documents/Reports` – uploads a Word/Excel/PDF/email file (max 50 MB).
- `POST /graph/users/drive/root/folders/secure` – creates a folder and grants access only to specified users.
- `GET /graph/sites/{siteId}/drives/{driveId}/folders?folderPath=/Shared Documents/Reports` – returns a folder from a SharePoint document library (app-only compatible).

## Folder access control

The `POST /graph/users/drive/root/folders/secure` endpoint creates a folder and restricts access to the recipient list you provide:

- The folder is created in the user's drive (OneDrive) under `folderPath`.
- Recipients are granted either `read` or `write` access.
- When `removeExistingPermissions` is `true`, non-owner permissions are removed so only the specified users retain access.

### Required permissions

To manage folder permissions, your app must have **application** permissions:

- `Files.ReadWrite.All`
- `Sites.ReadWrite.All`

Grant **admin consent** after adding these permissions.

## Swagger

Swagger UI is available at `/swagger` once the app is running.

## Notes

The app uses a confidential client with client credentials flow, which requires app permissions granted in Azure AD.
