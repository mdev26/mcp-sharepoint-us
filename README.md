# SharePoint MCP Server - Updated with Modern Authentication

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A comprehensive MCP Server for seamless integration with Microsoft SharePoint, **now with modern Azure AD authentication support** to work with new tenants.

## 🆕 What's New in v2.0

This is an updated fork of the original [mcp-sharepoint](https://github.com/Sofias-ai/mcp-sharepoint) that fixes the `Acquire app-only access token failed` error by implementing modern Azure AD authentication methods.

### Key Updates

- ✅ **Modern MSAL Authentication**: Uses Microsoft Authentication Library (MSAL) for Azure AD
- ✅ **Multiple Auth Methods**: Supports MSAL, certificate-based, and legacy authentication
- ✅ **New Tenant Compatible**: Works out-of-the-box with new Microsoft 365 tenants
- ✅ **Better Error Messages**: Clear guidance when authentication fails
- ✅ **Tenant ID Required**: Properly implements Azure AD app-only authentication

## 🔧 The Problem This Fixes

The original mcp-sharepoint used the deprecated ACS (Azure Access Control Service) authentication method with `with_client_credentials()`. This fails on new tenants with:

```
ValueError: Acquire app-only access token failed
```

**Why this happens:**
- Microsoft disabled ACS app-only authentication by default for new tenants
- The old method doesn't include tenant ID in the authentication flow
- Modern Azure AD authentication requires MSAL or certificate-based auth

**This update solves it by:**
- Using MSAL (Microsoft Authentication Library) by default
- Properly passing tenant ID to Azure AD
- Supporting multiple modern authentication methods
- Falling back gracefully with helpful error messages

## 📋 Prerequisites

### Azure AD App Registration

You need to register an application in Azure AD with the following:

1. **Go to Azure Portal** → Azure Active Directory → App registrations
2. **Create new registration**:
   - Name: "SharePoint MCP Server" (or your choice)
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: Not needed for app-only auth

3. **Configure API Permissions**:
   - Click "API permissions" → "Add a permission"
   - Choose "SharePoint" → "Application permissions"
   - Add these permissions:
     - `Sites.Read.All` (for read operations)
     - `Sites.ReadWrite.All` (for write operations)
   - **Important**: Click "Grant admin consent" for your organization

4. **Create Client Secret**:
   - Go to "Certificates & secrets"
   - Click "New client secret"
   - Choose expiration (recommend 24 months)
   - **Save the secret value immediately** (you won't see it again)

5. **Note these values**:
   - Application (client) ID
   - Directory (tenant) ID
   - Client secret value

### Required Information

You'll need:
- **Tenant ID**: Your Azure AD tenant ID (GUID)
- **Client ID**: Your Azure AD application ID (GUID)
- **Client Secret**: The secret value you created
- **Site URL**: Your SharePoint site URL (e.g., `https://contoso.sharepoint.com/sites/yoursite`)

## 🚀 Installation

### From Source (Recommended for this fork)

```bash
# Clone this repository
git clone https://github.com/your-username/mcp-sharepoint-updated.git
cd mcp-sharepoint-updated

# Install in development mode
pip install -e .
```

### Using uv (Alternative)

```bash
uv pip install -e .
```

## ⚙️ Configuration

### Environment Variables

Create a `.env` file or set these environment variables:

```bash
# Required - Modern Authentication
SHP_TENANT_ID=your-tenant-id-guid
SHP_ID_APP=your-client-id-guid
SHP_ID_APP_SECRET=your-client-secret
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site

# Optional
SHP_DOC_LIBRARY=Shared Documents
SHP_AUTH_METHOD=msal  # Options: msal (default), certificate, legacy

# For certificate-based authentication (optional)
# SHP_CERT_PATH=/path/to/certificate.pem
# SHP_CERT_THUMBPRINT=your-cert-thumbprint
```

### Authentication Methods

The server supports three authentication methods:

#### 1. MSAL (Recommended - Default)
Modern Azure AD authentication using MSAL. Works with new tenants.
```bash
SHP_AUTH_METHOD=msal
```

#### 2. Certificate-Based
For organizations requiring certificate authentication.
```bash
SHP_AUTH_METHOD=certificate
SHP_CERT_PATH=/path/to/cert.pem
SHP_CERT_THUMBPRINT=your-thumbprint
```

#### 3. Legacy (Deprecated)
Old ACS authentication. Only use if you have an older tenant with ACS enabled.
```bash
SHP_AUTH_METHOD=legacy
```

## 🔌 Claude Desktop Integration

### Windows
Edit: `%APPDATA%\Claude\claude_desktop_config.json`

### macOS
Edit: `~/Library/Application Support/Claude/claude_desktop_config.json`

### Configuration

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
        "SHP_DOC_LIBRARY": "Shared Documents",
        "SHP_AUTH_METHOD": "msal"
      }
    }
  }
}
```

### Using uvx (Alternative)

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "uvx",
      "args": ["mcp-sharepoint"],
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

## 🧪 Testing the Connection

After installation, you can test the connection:

```bash
# Set your environment variables first
export SHP_TENANT_ID=your-tenant-id
export SHP_ID_APP=your-client-id
export SHP_ID_APP_SECRET=your-secret
export SHP_SITE_URL=https://your-site.sharepoint.com/sites/yoursite

# Run the server
python -m mcp_sharepoint
```

In Claude Desktop, you can use the "Test_Connection" tool to verify everything is working.

## 🛠️ Available Tools

The server provides these tools for SharePoint operations:

### Document Management
- `List_SharePoint_Documents` - List all documents in a folder
- `Get_Document_Content` - Read document content
- `Upload_Document` - Upload new documents
- `Update_Document` - Update existing documents
- `Delete_Document` - Delete documents

### Folder Management
- `List_SharePoint_Folders` - List folders in a directory
- `Create_Folder` - Create new folders
- `Delete_Folder` - Delete empty folders
- `Get_SharePoint_Tree` - Get recursive folder structure

### Utilities
- `Test_Connection` - Test authentication and connection

## 🔍 Troubleshooting

### "Acquire app-only access token failed"

**This is the error we're fixing!** If you still see this:

1. **Check you're using MSAL**:
   ```bash
   SHP_AUTH_METHOD=msal  # Make sure this is set
   ```

2. **Verify tenant ID is correct**:
   - Go to Azure Portal → Azure Active Directory → Overview
   - Copy the "Tenant ID" GUID
   - Ensure it matches `SHP_TENANT_ID`

3. **Check API permissions**:
   - Azure Portal → Your App → API permissions
   - Ensure SharePoint permissions are granted
   - Click "Grant admin consent"

4. **Wait for permission propagation**:
   - After granting permissions, wait 5-10 minutes
   - Try again

### "Authentication failed"

1. **Verify credentials**:
   - Client ID is correct (from App registrations → Overview)
   - Client secret is correct and not expired
   - Tenant ID is correct

2. **Check site URL**:
   - Must be exact SharePoint site URL
   - Should NOT end with a slash
   - Example: `https://contoso.sharepoint.com/sites/marketing`

3. **Verify network access**:
   - Ensure you can reach `login.microsoftonline.com` on port 443
   - Check firewall/proxy settings

### "Access denied" / "403 Forbidden"

1. **Check SharePoint permissions**:
   - API permissions must be **Application** permissions, not Delegated
   - Need `Sites.Read.All` or `Sites.ReadWrite.All`
   - Admin consent must be granted

2. **Verify site access**:
   - The app must have access to the specific SharePoint site
   - May need to grant site permissions separately

## 📚 Migration from Original mcp-sharepoint

If you're migrating from the original version:

1. **Add new environment variable**:
   ```bash
   SHP_TENANT_ID=your-tenant-id  # This is new!
   ```

2. **Optional: Set auth method explicitly**:
   ```bash
   SHP_AUTH_METHOD=msal  # Explicitly use modern auth
   ```

3. **Update your Claude Desktop config**:
   - Add `SHP_TENANT_ID` to the env section
   - Optionally add `SHP_AUTH_METHOD: "msal"`

4. **Restart Claude Desktop** to load the new configuration

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Original [mcp-sharepoint](https://github.com/Sofias-ai/mcp-sharepoint) by sofias tech
- [Office365-REST-Python-Client](https://github.com/vgrem/Office365-REST-Python-Client) library
- Microsoft Authentication Library (MSAL) for Python

## 📞 Support

If you encounter issues:

1. Check the [Troubleshooting](#-troubleshooting) section
2. Review [Azure AD App Setup](#azure-ad-app-registration)
3. Open an issue with:
   - Error message (sanitize any secrets!)
   - Your environment (Python version, OS)
   - Authentication method being used
   - Whether it's a new or existing tenant

---

**Note**: This is an updated version that solves the authentication issues with new Microsoft 365 tenants. The original version can be found at [Sofias-ai/mcp-sharepoint](https://github.com/Sofias-ai/mcp-sharepoint).
