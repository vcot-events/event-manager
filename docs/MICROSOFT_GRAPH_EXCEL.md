# Microsoft Graph Excel API Setup

Microsoft Graph does not use a single "Excel API key." You register an app in Microsoft Entra ID, authorize it to OneDrive/SharePoint, and exchange OAuth credentials for short-lived access tokens.

## 1. Register the app

1. Go to the Microsoft Entra admin center and open **Identity > Applications > App registrations**.
2. Choose **New registration**.
3. Name it something like `TritonAI Event Manager`.
4. Use a redirect URI appropriate for the app:
   - local prototype: `http://localhost:3000/auth/callback`
   - hosted app: the exact HTTPS callback URL
5. Save the **Application (client) ID** and **Directory (tenant) ID**.
6. Under **Certificates & secrets**, create a client secret if using a server-side confidential app.

## 2. Request delegated permissions

For files owned by your UCSD account and shared files you can access, request:

- `User.Read`
- `Files.ReadWrite.All`
- `Sites.ReadWrite.All`

For a first prototype, delegated permissions are safer than application permissions because the agent acts as the signed-in user. UCSD may require admin approval for one or more scopes.

## 3. Get an access token

Use OAuth 2.0 authorization-code flow with PKCE:

```text
https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/authorize
https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token
scope=https://graph.microsoft.com/Files.ReadWrite.All offline_access
```

Store the refresh token securely and request new access tokens as needed.

## 4. Excel endpoints

Find a file by OneDrive path:

```http
GET https://graph.microsoft.com/v1.0/me/drive/root:/Event Manager/DFP Staff Check-In 2026.xlsx
```

Read a worksheet range:

```http
GET https://graph.microsoft.com/v1.0/me/drive/items/{item-id}/workbook/worksheets('Sep 15')/range(address='A1:G25')
```

Write a range:

```http
PATCH https://graph.microsoft.com/v1.0/me/drive/items/{item-id}/workbook/worksheets('Sep 15')/range(address='A4:D25')
Content-Type: application/json

{
  "values": [
    ["Auzenne", "Keiara", "kauzenne@ucsd.edu", null],
    ["Barasi", "Ramadan", "rbarasi@ucsd.edu", null]
  ]
}
```

For files in SharePoint/Teams rather than personal OneDrive, first resolve the site and drive, then use the same workbook endpoints with that `drive-id` and `item-id`.

## 5. MVP access request

Ask UCSD IT/your Microsoft admin for:

1. Permission to register an Entra app for `vcotevents@ucsd.edu` or a service account.
2. Delegated `Files.ReadWrite.All` and `Sites.ReadWrite.All` access.
3. Access to the exact SharePoint/OneDrive folder containing event files.
4. Confirmation whether the workbook is in personal OneDrive, a shared OneDrive, or SharePoint.
