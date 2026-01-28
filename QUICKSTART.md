# Quick Start Guide

Get up and running with SharePoint MCP Server in 10 minutes.

## TL;DR - Minimal Setup

```bash
# 1. Clone and install
git clone <your-repo-url>
cd mcp-sharepoint-updated
pip install -e .

# 2. Set environment variables
export SHP_TENANT_ID="your-tenant-id"
export SHP_ID_APP="your-client-id"
export SHP_ID_APP_SECRET="your-client-secret"
export SHP_SITE_URL="https://your-site.sharepoint.com/sites/yoursite"

# 3. Test it
python -m mcp_sharepoint
```

## Step 1: Azure AD Setup (5 minutes)

### Get Your Credentials

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Name it "SharePoint MCP Server", click **Register**
4. Copy these values:
   - **Application (client) ID** → This is `SHP_ID_APP`
   - **Directory (tenant) ID** → This is `SHP_TENANT_ID`

### Create Secret

1. Go to **Certificates & secrets** → **New client secret**
2. Set expiration to 24 months
3. Copy the **Value** immediately → This is `SHP_ID_APP_SECRET`

### Grant Permissions

1. Go to **API permissions** → **Add a permission**
2. Select **SharePoint** → **Application permissions**
3. Add `Sites.ReadWrite.All`
4. Click **Grant admin consent**

👉 **Detailed guide**: See [AZURE_PORTAL_GUIDE.md](AZURE_PORTAL_GUIDE.md)

## Step 2: Get SharePoint Site URL

1. Open your SharePoint site in a browser
2. Copy the URL (e.g., `https://contoso.sharepoint.com/sites/marketing`)
3. This is your `SHP_SITE_URL`

## Step 3: Install

```bash
# Clone the repository
git clone <your-repo-url>
cd mcp-sharepoint-updated

# Install
pip install -e .
```

## Step 4: Configure

Create a `.env` file:

```bash
SHP_TENANT_ID=12345678-1234-1234-1234-123456789abc
SHP_ID_APP=87654321-4321-4321-4321-abcdef123456
SHP_ID_APP_SECRET=your-super-secret-value-here
SHP_SITE_URL=https://contoso.sharepoint.com/sites/yoursite
SHP_AUTH_METHOD=msal
```

Or export as environment variables:

```bash
export SHP_TENANT_ID="12345678-1234-1234-1234-123456789abc"
export SHP_ID_APP="87654321-4321-4321-4321-abcdef123456"
export SHP_ID_APP_SECRET="your-super-secret-value-here"
export SHP_SITE_URL="https://contoso.sharepoint.com/sites/yoursite"
export SHP_AUTH_METHOD="msal"
```

## Step 5: Test Connection

```bash
python -m mcp_sharepoint
```

In Claude Desktop, use the "Test_Connection" tool. You should see:
```
✓ Successfully connected to SharePoint!
Site Title: Your Site Name
Authentication Method: MSAL
```

## Step 6: Claude Desktop Integration

Edit Claude Desktop config:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

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
        "SHP_SITE_URL": "https://your-site.sharepoint.com/sites/yoursite",
        "SHP_AUTH_METHOD": "msal"
      }
    }
  }
}
```

Restart Claude Desktop.

## Common Issues

### "Acquire app-only access token failed"

**Fix**: This is exactly what we're solving! Make sure:
- `SHP_TENANT_ID` is set correctly
- `SHP_AUTH_METHOD=msal` (or leave it unset, msal is default)
- You've granted admin consent in Azure Portal

### "403 Forbidden" / "Access denied"

**Fix**: 
- Make sure you used **Application permissions**, not Delegated
- Click "Grant admin consent" in Azure Portal
- Wait 5-10 minutes for permissions to propagate

### "Invalid client secret"

**Fix**:
- Make sure you copied the secret **value**, not the secret ID
- Check for extra spaces in your `.env` file
- The secret might have expired - create a new one

## Next Steps

Once connected, try these commands in Claude:

```
1. "List all documents in my SharePoint library"
2. "Show me the folder structure"
3. "Upload a new document called test.txt with content 'Hello World'"
4. "Read the content of test.txt"
```

## Need Help?

- 📖 Full documentation: [README.md](README.md)
- 🔧 Azure setup details: [AZURE_PORTAL_GUIDE.md](AZURE_PORTAL_GUIDE.md)
- 🐛 Open an issue on GitHub
