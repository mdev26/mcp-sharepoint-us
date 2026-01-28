# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2025-01-28

### Added

- **Modern MSAL Authentication**: Primary authentication method using Microsoft Authentication Library (MSAL)
- **Certificate-Based Authentication**: Support for certificate-based app-only authentication
- **Multi-Authentication Support**: Choose between MSAL, certificate, or legacy authentication methods
- **Tenant ID Support**: Required `SHP_TENANT_ID` environment variable for proper Azure AD authentication
- **Authentication Method Selection**: New `SHP_AUTH_METHOD` environment variable to select auth method
- **Test Connection Tool**: New tool to verify SharePoint connection and authentication
- **Better Error Messages**: Clear, actionable error messages for authentication failures
- **Comprehensive Documentation**: 
  - Detailed Azure Portal setup guide
  - Quick start guide
  - Migration guide from v1.x
  - Troubleshooting sections
- **Certificate Support**: Environment variables for certificate-based authentication
- **Improved Logging**: Better logging for debugging authentication issues

### Changed

- **BREAKING**: Tenant ID is now required (`SHP_TENANT_ID`)
- **Default Authentication**: Changed from ACS (legacy) to MSAL (modern) by default
- **Authentication Flow**: Now uses token-based authentication via MSAL instead of client credentials
- **Dependency Updates**: Updated to latest MSAL and office365-rest-python-client versions
- **Error Handling**: Improved error messages with troubleshooting hints

### Fixed

- **"Acquire app-only access token failed" Error**: Fixed by implementing modern Azure AD authentication
- **New Tenant Compatibility**: Now works with tenants where ACS app-only is disabled
- **Authentication Failures**: Better handling of various authentication failure scenarios

### Deprecated

- **Legacy ACS Authentication**: Still available via `SHP_AUTH_METHOD=legacy` but deprecated
- Will be removed in version 3.0.0 when Microsoft fully disables ACS

### Security

- **Modern Authentication**: More secure authentication flow using Azure AD
- **Token-Based Auth**: Uses MSAL for proper token acquisition and management
- **Certificate Support**: Option to use certificate-based authentication for enhanced security

## [1.0.0] - 2024-XX-XX (Original Version)

### Features from Original Version

- List SharePoint folders
- List SharePoint documents
- Get document content (with text extraction from PDF, Word, Excel)
- Upload documents (text and binary)
- Update documents
- Delete documents
- Create folders
- Delete folders
- Get folder tree structure
- Support for multiple file types
- Base64 encoding for binary files

### Original Authentication

- Used ACS (Azure Access Control Service) authentication
- Required only client ID and client secret
- Did not require tenant ID

---

## Migration Notes

### From 1.x to 2.x

**Required Changes**:
1. Add `SHP_TENANT_ID` to your environment configuration
2. Optionally set `SHP_AUTH_METHOD=msal` (or leave unset for default)
3. Verify Azure AD app permissions are configured correctly
4. Grant admin consent in Azure Portal

**Optional Changes**:
- Use certificate-based authentication for enhanced security
- Review and update API permissions if needed

See [MIGRATION_GUIDE.md](MIGRATION_GUIDE.md) for detailed migration instructions.

## Compatibility

### Version 2.0.0
- **Python**: 3.10, 3.11, 3.12
- **office365-rest-python-client**: >=2.5.0
- **msal**: >=1.24.0
- **mcp**: >=1.0.0
- **Microsoft 365**: All tenants (new and existing)
- **SharePoint**: SharePoint Online

### Version 1.0.0
- **Python**: 3.10+
- **Microsoft 365**: Existing tenants with ACS enabled only
- **SharePoint**: SharePoint Online

## Known Issues

### Version 2.0.0
- None currently known

### Version 1.0.0
- ❌ Fails on new tenants with "Acquire app-only access token failed"
- ❌ Uses deprecated ACS authentication
- ❌ May stop working when Microsoft fully disables ACS

## Roadmap

### Version 2.1.0 (Planned)
- [ ] Interactive authentication support
- [ ] Improved file upload for large files
- [ ] Batch operations support
- [ ] SharePoint list operations
- [ ] Site collection management

### Version 3.0.0 (Future)
- [ ] Remove legacy ACS authentication support
- [ ] Require Python 3.11+
- [ ] Enhanced error handling
- [ ] Performance optimizations

## Contributing

We welcome contributions! Please see the contributing guidelines in the README.

## Credits

- Original [mcp-sharepoint](https://github.com/Sofias-ai/mcp-sharepoint) by sofias tech
- Updated version with modern authentication by community contributors
- Built using [Office365-REST-Python-Client](https://github.com/vgrem/Office365-REST-Python-Client)
- Authentication via [MSAL Python](https://github.com/AzureAD/microsoft-authentication-library-for-python)
