# Azure Portal Setup Guide

This guide walks you through setting up an Azure AD application for SharePoint MCP Server with modern authentication.

## Prerequisites

- Access to Azure Portal with admin rights
- Permission to register applications in Azure AD
- Permission to grant admin consent for SharePoint permissions

## Step-by-Step Setup

### Step 1: Navigate to Azure Portal

1. Go to [https://portal.azure.com](https://portal.azure.com)
2. Sign in with your Microsoft account
3. Navigate to **Azure Active Directory** (you can search for it in the top search bar)

### Step 2: Register a New Application

1. In the left sidebar, click **App registrations**
2. Click **+ New registration** at the top
3. Fill in the registration form:
   - **Name**: `SharePoint MCP Server` (or any name you prefer)
   - **Supported account types**: Select **"Accounts in this organizational directory only (Single tenant)"**
   - **Redirect URI**: Leave empty (not needed for app-only authentication)
4. Click **Register**

### Step 3: Note Your Application IDs

After registration, you'll see the Overview page. **Save these values**:

1. **Application (client) ID**: This is your `SHP_ID_APP`
   - Format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - Example: `a1b2c3d4-e5f6-1234-5678-9abcdef01234`

2. **Directory (tenant) ID**: This is your `SHP_TENANT_ID`
   - Format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - Found under "Directory (tenant) ID" on the Overview page

### Step 4: Create a Client Secret

1. In the left sidebar, click **Certificates & secrets**
2. Under "Client secrets", click **+ New client secret**
3. Configure the secret:
   - **Description**: `MCP Server Secret` (or any description)
   - **Expires**: Choose an expiration period
     - Recommended: **24 months** for production
     - For testing: **3 months** is fine
4. Click **Add**
5. **IMPORTANT**: Copy the **Value** immediately!
   - This is your `SHP_ID_APP_SECRET`
   - You won't be able to see it again after you navigate away
   - Store it securely (e.g., in a password manager)

### Step 5: Configure API Permissions

This is the most important step for fixing the authentication issue!

1. In the left sidebar, click **API permissions**
2. Click **+ Add a permission**
3. In the "Request API permissions" panel:
   - Click **APIs my organization uses**
   - Search for **SharePoint**
   - Click **SharePoint** in the results

4. Select **Application permissions** (NOT Delegated permissions!)
   - This is crucial for app-only authentication

5. Under "Application permissions", expand **Sites**:
   - Check `Sites.Read.All` (if you only need read access)
   - OR check `Sites.ReadWrite.All` (if you need read and write access)

6. Click **Add permissions**

7. **Grant Admin Consent** (Critical!)
   - After adding permissions, you'll see them listed
   - Click the **"Grant admin consent for [Your Organization]"** button
   - Click **Yes** in the confirmation dialog
   - The status should change to green checkmarks

### Step 6: Verify Your Configuration

Your API permissions should look like this:

| API / Permissions name | Type | Status |
|------------------------|------|--------|
| SharePoint Sites.Read.All or Sites.ReadWrite.All | Application | ✓ Granted for [Your Org] |

The green checkmark is essential!

### Step 7: Get Your SharePoint Site URL

1. Open SharePoint in a browser
2. Navigate to the site you want to use
3. Copy the URL from the address bar
4. The URL should look like:
   - `https://contoso.sharepoint.com/sites/marketing`
   - `https://contoso.sharepoint.com/sites/team-site`
5. This is your `SHP_SITE_URL`

**Important**: 
- Do NOT include a trailing slash
- Use the full site URL, not just the domain

## Configuration Summary

After completing these steps, you should have:

| Variable | Value | Where to Find It |
|----------|-------|------------------|
| `SHP_TENANT_ID` | GUID | Azure AD → App registrations → Your app → Overview → "Directory (tenant) ID" |
| `SHP_ID_APP` | GUID | Azure AD → App registrations → Your app → Overview → "Application (client) ID" |
| `SHP_ID_APP_SECRET` | Secret string | Azure AD → App registrations → Your app → Certificates & secrets → Client secrets (created in Step 4) |
| `SHP_SITE_URL` | URL | SharePoint site URL (e.g., https://contoso.sharepoint.com/sites/yoursite) |

## Testing Your Setup

1. Create a `.env` file with your values:
```bash
SHP_TENANT_ID=your-tenant-id
SHP_ID_APP=your-client-id
SHP_ID_APP_SECRET=your-client-secret
SHP_SITE_URL=https://your-site.sharepoint.com/sites/yoursite
SHP_AUTH_METHOD=msal
```

2. Run the test:
```bash
python -m mcp_sharepoint
```

3. In Claude Desktop, use the "Test_Connection" tool

## Troubleshooting

### "Grant admin consent" button is grayed out

**Problem**: You don't have admin rights in Azure AD.

**Solutions**:
1. Ask your IT administrator to grant consent
2. Or ask them to give you the "Application Administrator" role in Azure AD

### Permissions not showing green checkmark

**Problem**: Admin consent wasn't granted properly.

**Solution**:
1. Remove and re-add the permissions
2. Make sure to click "Grant admin consent" again
3. Wait 5-10 minutes for changes to propagate

### "Application permissions only" vs "Delegated permissions"

**Important**: For app-only authentication (which this MCP server uses), you MUST use **Application permissions**, not Delegated permissions.

- **Application permissions**: App acts on its own behalf (what we need)
- **Delegated permissions**: App acts on behalf of a signed-in user (not suitable)

### Need both Read and Write access?

If you need to:
- Read documents: `Sites.Read.All` is sufficient
- Write/upload documents: Use `Sites.ReadWrite.All` instead
- Delete documents: Use `Sites.ReadWrite.All`

You can always change permissions later:
1. Go back to API permissions
2. Remove the old permission
3. Add the new one
4. Grant admin consent again

## Security Best Practices

1. **Limit Secret Expiration**:
   - Shorter expiration = more secure
   - Set calendar reminders to rotate secrets before expiry

2. **Principle of Least Privilege**:
   - Only request the permissions you actually need
   - Use `Sites.Read.All` if you only need read access

3. **Secure Secret Storage**:
   - Never commit `.env` files to Git
   - Use environment variables or secret management systems in production
   - Consider using Azure Key Vault for production deployments

4. **Regular Audits**:
   - Review which apps have access to SharePoint
   - Remove unused app registrations
   - Rotate secrets regularly

## Certificate-Based Authentication (Advanced)

If your organization requires certificate-based authentication:

1. In **Certificates & secrets**, click the **Certificates** tab
2. Click **Upload certificate**
3. Upload your `.cer` or `.pem` file
4. Note the **Thumbprint** value
5. Set environment variables:
   ```bash
   SHP_AUTH_METHOD=certificate
   SHP_CERT_PATH=/path/to/your/certificate.pem
   SHP_CERT_THUMBPRINT=your-thumbprint-here
   ```

## Need Help?

If you're stuck:

1. Verify each step carefully
2. Double-check all GUIDs and secrets are copied correctly
3. Ensure there are no extra spaces in your `.env` file
4. Wait 5-10 minutes after granting permissions
5. Check the [README.md](README.md) troubleshooting section

## Additional Resources

- [Microsoft: Register an application with Azure AD](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft: SharePoint app-only authentication](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs)
- [MSAL Python Documentation](https://msal-python.readthedocs.io/)
