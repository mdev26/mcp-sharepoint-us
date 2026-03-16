# SharePoint MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

MCP Server for Microsoft SharePoint using modern Azure AD (MSAL) authentication.

## Prerequisites

### Azure AD App Registration

1. Go to **Azure Portal** → **Azure Active Directory** → **App registrations** → **New registration**
   - Name: anything you like
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: leave empty

2. From the **Overview** page, note:
   - **Application (client) ID** → `SHP_ID_APP`
   - **Directory (tenant) ID** → `SHP_TENANT_ID`

3. **Certificates & secrets** → **New client secret**
   - Save the **Value** immediately (you won't see it again) → `SHP_ID_APP_SECRET`

4. **API permissions** → **Add a permission** → **SharePoint** → **Application permissions**
   - Add `Sites.ReadWrite.All` (or `Sites.Read.All` for read-only)
   - Click **Grant admin consent** — the status must show a green checkmark

5. Get your SharePoint site URL (e.g. `https://contoso.sharepoint.com/sites/yoursite`) → `SHP_SITE_URL`
   - Do NOT include a trailing slash

### Security Best Practices

- Use `Sites.Read.All` if you only need read access (principle of least privilege)
- Set a calendar reminder to rotate client secrets before expiry
- Never commit `.env` files to Git — use environment variables or a secrets manager

## Installation

```bash
pip install mcp-sharepoint-us
```

Or from source:

```bash
git clone https://github.com/mdev26/mcp-sharepoint-us.git
cd mcp-sharepoint-us
pip install -e .
```

## Configuration

```bash
# Required
SHP_TENANT_ID=your-tenant-id-guid
SHP_ID_APP=your-client-id-guid
SHP_ID_APP_SECRET=your-client-secret
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site

# Optional
SHP_DOC_LIBRARY=Shared Documents   # default
SHP_AUTH_METHOD=msal               # options: msal (default), certificate, legacy
```

### Certificate-based authentication (optional)

```bash
SHP_AUTH_METHOD=certificate
SHP_CERT_PATH=/path/to/certificate.pem
SHP_CERT_THUMBPRINT=your-cert-thumbprint
```

## Claude Desktop Integration

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "python",
      "args": ["-m", "mcp_sharepoint"],
      "env": {
        "SHP_TENANT_ID": "your-tenant-id",
        "SHP_ID_APP": "your-client-id",
        "SHP_ID_APP_SECRET": "your-client-secret",
        "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
        "SHP_AUTH_METHOD": "msal"
      }
    }
  }
}
```

### Using uvx

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "uvx",
      "args": ["mcp-sharepoint-us"],
      "env": {
        "SHP_TENANT_ID": "your-tenant-id",
        "SHP_ID_APP": "your-client-id",
        "SHP_ID_APP_SECRET": "your-client-secret",
        "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site"
      }
    }
  }
}
```

## Available Tools

| Tool | Description |
|------|-------------|
| `Test_Connection` | Verify authentication and connection |
| `List_SharePoint_Documents` | List documents in a folder |
| `Get_Document_Content` | Read document content (supports .docx, .pptx, .xlsx, .pdf) |
| `Upload_Document` | Upload a new document |
| `Update_Document` | Update an existing document |
| `Delete_Document` | Delete a document |
| `List_SharePoint_Folders` | List folders |
| `Create_Folder` | Create a new folder |
| `Delete_Folder` | Delete an empty folder |
| `Get_SharePoint_Tree` | Get recursive folder structure |
| `Create_Word_Document` | Create a formatted .docx and upload to SharePoint |
| `Edit_Word_Document` | Find/replace or section-replace content in a .docx |
| `Create_PowerPoint` | Create a .pptx and upload to SharePoint |

## Troubleshooting

### Enable debug logging

```bash
LOGLEVEL=DEBUG python -m mcp_sharepoint
```

### "Acquire app-only access token failed"

- Ensure `SHP_TENANT_ID` is set and correct (Azure Portal → Azure AD → Overview → Tenant ID)
- Ensure `SHP_AUTH_METHOD=msal` (or leave unset — msal is the default)
- Verify admin consent is granted in Azure Portal (green checkmarks on API permissions)
- After granting permissions, wait 5–10 minutes for propagation

### "403 Forbidden" / "Access denied"

- Permissions must be **Application** permissions, not Delegated
- Admin consent must be granted
- The site URL must exactly match the SharePoint site (no trailing slash)

### "Invalid client secret"

- Copy the secret **Value**, not the secret ID
- Check for extra spaces in your `.env` file
- The secret may have expired — create a new one

### Connection reset / firewall issues

If authentication succeeds but Graph API calls fail (connection reset during TLS), the endpoint `graph.microsoft.us` (US Government) or `graph.microsoft.com` (commercial) may be blocked by a firewall using deep packet inspection. Ask your network team to whitelist the endpoint on port 443. For proxy environments:

```bash
export HTTPS_PROXY=http://proxy.company.com:8080
```

## License

MIT License
