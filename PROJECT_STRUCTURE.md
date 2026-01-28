# Project Structure

```
mcp-sharepoint-updated/
│
├── src/
│   └── mcp_sharepoint/
│       ├── __init__.py          # Main MCP server implementation
│       ├── __main__.py          # Entry point for python -m mcp_sharepoint
│       └── auth.py              # Authentication module (MSAL, Certificate, Legacy)
│
├── Documentation/
│   ├── README.md               # Main documentation
│   ├── QUICKSTART.md           # 10-minute setup guide
│   ├── AZURE_PORTAL_GUIDE.md   # Detailed Azure AD setup
│   ├── MIGRATION_GUIDE.md      # v1 to v2 migration guide
│   ├── CHANGELOG.md            # Version history
│   └── SUMMARY.md              # Technical summary
│
├── Configuration Files/
│   ├── pyproject.toml          # Package configuration
│   ├── requirements.txt        # Pip dependencies
│   ├── .env.example            # Environment variable template
│   ├── .gitignore              # Git ignore rules
│   └── LICENSE                 # MIT License
│
├── Testing/
│   └── test_connection.py      # Configuration & connection test script
│
└── README.md                   # You are here!
```

## File Purposes

### Core Implementation

**src/mcp_sharepoint/__init__.py**
- Main MCP server implementation
- All 10 SharePoint tools
- Uses modern authentication
- Backwards compatible with original features

**src/mcp_sharepoint/auth.py**
- SharePointAuthenticator class
- MSAL authentication (default)
- Certificate-based authentication
- Legacy ACS authentication (deprecated)
- Factory functions for easy use

**src/mcp_sharepoint/__main__.py**
- Package entry point
- Enables `python -m mcp_sharepoint` execution

### Documentation (Start Here!)

**README.md** ⭐ START HERE
- Complete project documentation
- Setup instructions
- Troubleshooting guide
- All features explained

**QUICKSTART.md** - For the impatient
- 10-minute setup
- Minimal instructions
- Get running fast

**AZURE_PORTAL_GUIDE.md** - Azure AD setup
- Step-by-step Azure Portal walkthrough
- Screenshot-level detail
- Permission configuration
- Security best practices

**MIGRATION_GUIDE.md** - Upgrading from v1
- What changed
- Migration steps
- Backwards compatibility
- Rollback instructions

**CHANGELOG.md** - Version history
- All changes documented
- Migration notes between versions
- Known issues
- Future roadmap

**SUMMARY.md** - Technical overview
- Architecture explanation
- Code changes
- Implementation details
- Developer reference

### Configuration

**pyproject.toml**
- Package metadata
- Dependencies (MSAL, office365, mcp, etc.)
- Build configuration
- Entry points

**requirements.txt**
- Pip-installable dependencies
- Alternative to pyproject.toml
- For traditional pip workflows

**.env.example**
- Template for environment variables
- Copy to `.env` and fill in your values
- Documents all configuration options

**.gitignore**
- Prevents committing sensitive files
- Standard Python ignores
- `.env` excluded from git

**LICENSE**
- MIT License
- Same as original project

### Testing

**test_connection.py** ⭐ RUN THIS FIRST
- Validates configuration
- Tests SharePoint connection
- Checks all prerequisites
- Provides troubleshooting guidance

## Quick Reference

### Installation
```bash
git clone <your-repo>
cd mcp-sharepoint-updated
pip install -e .
```

### Configuration
```bash
cp .env.example .env
# Edit .env with your values
```

### Testing
```bash
python test_connection.py
```

### Usage with Claude
See README.md section on Claude Desktop integration

## What to Read First

1. **Getting Started**: README.md → Overview & Features
2. **Azure Setup**: AZURE_PORTAL_GUIDE.md → Get credentials
3. **Quick Setup**: QUICKSTART.md → Fast installation
4. **Testing**: Run test_connection.py
5. **Integration**: README.md → Claude Desktop config

## What to Read for Specific Needs

| Need | Read This |
|------|-----------|
| First time setup | QUICKSTART.md |
| Azure AD configuration | AZURE_PORTAL_GUIDE.md |
| Migrating from v1 | MIGRATION_GUIDE.md |
| Understanding changes | CHANGELOG.md, SUMMARY.md |
| Troubleshooting | README.md (Troubleshooting section) |
| Technical details | SUMMARY.md |
| API reference | README.md (Available Tools section) |

## Key Features

✅ Modern MSAL authentication (fixes "Acquire app-only access token failed")
✅ Works with new Microsoft 365 tenants
✅ Multiple authentication methods
✅ All original features preserved
✅ Comprehensive documentation
✅ Easy migration from v1
✅ Configuration testing script
✅ Detailed troubleshooting

## Support

- 📖 Documentation in this repository
- 🧪 Test your setup with test_connection.py
- 🐛 Open an issue on GitHub
- 💬 Check discussions for Q&A

## Contributing

Contributions welcome! See README.md for guidelines.

## License

MIT License - See LICENSE file
