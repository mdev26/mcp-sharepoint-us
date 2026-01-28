# Migration Guide from Original mcp-sharepoint

This guide helps you migrate from the original [Sofias-ai/mcp-sharepoint](https://github.com/Sofias-ai/mcp-sharepoint) to this updated version with modern authentication.

## Why Migrate?

The original version uses deprecated ACS (Azure Access Control Service) authentication, which:
- ❌ Doesn't work on new Microsoft 365 tenants
- ❌ Fails with "Acquire app-only access token failed" error
- ❌ Will be completely disabled by Microsoft in the future

This updated version:
- ✅ Uses modern MSAL (Microsoft Authentication Library)
- ✅ Works on all tenants (new and existing)
- ✅ Future-proof and actively maintained by Microsoft
- ✅ More secure authentication flow

## What's Changed?

### New Requirement: Tenant ID

The biggest change is that you now **must** provide your Azure AD tenant ID:

**Before (Original)**:
```bash
SHP_ID_APP=your-client-id
SHP_ID_APP_SECRET=your-client-secret
SHP_SITE_URL=https://your-site.sharepoint.com
SHP_DOC_LIBRARY=Shared Documents
```

**After (Updated)**:
```bash
SHP_TENANT_ID=your-tenant-id  # ← NEW! Required!
SHP_ID_APP=your-client-id
SHP_ID_APP_SECRET=your-client-secret
SHP_SITE_URL=https://your-site.sharepoint.com
SHP_DOC_LIBRARY=Shared Documents
SHP_AUTH_METHOD=msal  # ← NEW! Optional (default: msal)
```

### New Authentication Methods

The updated version supports multiple authentication methods:

| Method | When to Use | Setting |
|--------|-------------|---------|
| **MSAL** (default) | New tenants, modern setup | `SHP_AUTH_METHOD=msal` |
| **Certificate** | Enterprise requirements | `SHP_AUTH_METHOD=certificate` |
| **Legacy** | Old tenants with ACS enabled | `SHP_AUTH_METHOD=legacy` |

## Migration Steps

### Step 1: Get Your Tenant ID

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **Overview**
3. Copy the **Tenant ID** (it's a GUID)
   - Example: `12345678-1234-1234-1234-123456789abc`

### Step 2: Verify Azure AD Permissions

Make sure your app has the correct permissions:

1. Go to **Azure Active Directory** → **App registrations**
2. Find your existing app
3. Click **API permissions**
4. Verify you have:
   - **SharePoint** → **Application permissions** → **Sites.ReadWrite.All** (or Sites.Read.All)
   - Status should show "✓ Granted"

If not configured correctly:
- Follow the [Azure Portal Guide](AZURE_PORTAL_GUIDE.md)
- Make sure to grant admin consent

### Step 3: Update Your Configuration

#### Option A: Using .env file

Update your `.env` file:

```bash
# Add this new line
SHP_TENANT_ID=your-tenant-id-from-step-1

# Keep your existing values
SHP_ID_APP=your-existing-client-id
SHP_ID_APP_SECRET=your-existing-client-secret
SHP_SITE_URL=https://your-site.sharepoint.com/sites/yoursite
SHP_DOC_LIBRARY=Shared Documents

# Add this for explicit modern auth
SHP_AUTH_METHOD=msal
```

#### Option B: Claude Desktop Config

Update `claude_desktop_config.json`:

**Before**:
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "mcp-sharepoint",
      "env": {
        "SHP_ID_APP": "your-client-id",
        "SHP_ID_APP_SECRET": "your-client-secret",
        "SHP_SITE_URL": "https://your-site.sharepoint.com"
      }
    }
  }
}
```

**After**:
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "python",
      "args": ["-m", "mcp_sharepoint"],
      "env": {
        "SHP_TENANT_ID": "your-tenant-id",  // ← ADD THIS
        "SHP_ID_APP": "your-client-id",
        "SHP_ID_APP_SECRET": "your-client-secret",
        "SHP_SITE_URL": "https://your-site.sharepoint.com",
        "SHP_AUTH_METHOD": "msal"  // ← ADD THIS (optional)
      }
    }
  }
}
```

### Step 4: Install Updated Version

```bash
# Uninstall old version (if installed via pip)
pip uninstall mcp-sharepoint

# Clone and install updated version
git clone <your-updated-repo-url>
cd mcp-sharepoint-updated
pip install -e .
```

### Step 5: Test the Migration

```bash
# Test with environment variables
export SHP_TENANT_ID="your-tenant-id"
export SHP_ID_APP="your-client-id"
export SHP_ID_APP_SECRET="your-secret"
export SHP_SITE_URL="https://your-site.sharepoint.com/sites/yoursite"

python -m mcp_sharepoint
```

Use the "Test_Connection" tool in Claude Desktop. You should see:
```
✓ Successfully connected to SharePoint!
Authentication Method: MSAL
```

## Troubleshooting Migration

### Still Getting "Acquire app-only access token failed"?

**Checklist**:
- [ ] Added `SHP_TENANT_ID` to your configuration
- [ ] Set `SHP_AUTH_METHOD=msal` (or left it unset, as msal is default)
- [ ] Verified tenant ID is correct in Azure Portal
- [ ] Granted admin consent for API permissions
- [ ] Waited 5-10 minutes after changing permissions

### "Legacy" Mode for Backwards Compatibility

If you absolutely need to use the old authentication method temporarily:

```bash
SHP_AUTH_METHOD=legacy
```

**Warning**: This may not work on new tenants and is deprecated. Use only as a temporary measure.

### Certificate Authentication for Enterprises

If your organization requires certificate-based authentication:

```bash
SHP_AUTH_METHOD=certificate
SHP_CERT_PATH=/path/to/certificate.pem
SHP_CERT_THUMBPRINT=your-cert-thumbprint
```

See [Azure Portal Guide](AZURE_PORTAL_GUIDE.md) for certificate setup.

## Feature Comparison

All features from the original version are preserved:

| Feature | Original | Updated | Notes |
|---------|----------|---------|-------|
| List folders | ✅ | ✅ | Identical |
| List documents | ✅ | ✅ | Identical |
| Read documents | ✅ | ✅ | Identical |
| Upload documents | ✅ | ✅ | Identical |
| Update documents | ✅ | ✅ | Identical |
| Delete documents | ✅ | ✅ | Identical |
| Create folders | ✅ | ✅ | Identical |
| Delete folders | ✅ | ✅ | Identical |
| Get tree view | ✅ | ✅ | Identical |
| Test connection | ❌ | ✅ | **New!** |
| Modern auth | ❌ | ✅ | **New!** |
| Multi-auth support | ❌ | ✅ | **New!** |

## Breaking Changes

### Authentication Method

- **Before**: Used `with_client_credentials()` (ACS)
- **After**: Uses `with_access_token()` with MSAL (Azure AD)

This is handled automatically - you just need to provide `SHP_TENANT_ID`.

### Environment Variables

- **New required**: `SHP_TENANT_ID`
- **New optional**: `SHP_AUTH_METHOD`
- All other variables remain the same

### Installation

- **Before**: `pip install mcp-sharepoint`
- **After**: Install from source (until published to PyPI)

## Rollback Plan

If you need to rollback to the original version:

```bash
# Uninstall updated version
pip uninstall mcp-sharepoint

# Reinstall original version
pip install mcp-sharepoint

# Remove SHP_TENANT_ID from your config
# Restore original claude_desktop_config.json
```

**Note**: Rollback will only work if your tenant still supports legacy ACS authentication.

## Need Help?

- 📖 [Full README](README.md)
- 🔧 [Azure Portal Setup Guide](AZURE_PORTAL_GUIDE.md)
- 🚀 [Quick Start Guide](QUICKSTART.md)
- 🐛 [Open an issue](https://github.com/your-repo/issues)

## Timeline

- **Now**: Both versions can coexist (if your tenant supports legacy auth)
- **Soon**: Microsoft will fully disable ACS authentication
- **Future**: Only modern authentication will work

Migrate now to avoid service interruption!
