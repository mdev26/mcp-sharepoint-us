#!/usr/bin/env python3
"""
SharePoint MCP Server - Configuration Test Script

Run this script to verify your SharePoint connection and authentication
before integrating with Claude Desktop.

Usage:
    python test_connection.py

Requirements:
    - Environment variables set (or .env file)
    - Azure AD app configured with proper permissions
"""

import os
import sys
from dotenv import load_dotenv

# Load environment variables from .env file if present
load_dotenv()

def print_header(text):
    """Print a formatted header"""
    print(f"\n{'='*60}")
    print(f"  {text}")
    print(f"{'='*60}\n")

def print_success(text):
    """Print success message"""
    print(f"✓ {text}")

def print_error(text):
    """Print error message"""
    print(f"✗ {text}")

def print_warning(text):
    """Print warning message"""
    print(f"⚠ {text}")

def check_environment_variables():
    """Check if required environment variables are set"""
    print_header("Checking Environment Variables")
    
    required_vars = {
        'SHP_TENANT_ID': 'Azure AD Tenant ID',
        'SHP_ID_APP': 'Azure AD Application (Client) ID',
        'SHP_ID_APP_SECRET': 'Azure AD Client Secret',
        'SHP_SITE_URL': 'SharePoint Site URL'
    }
    
    optional_vars = {
        'SHP_DOC_LIBRARY': 'Document Library Path (default: "Shared Documents")',
        'SHP_AUTH_METHOD': 'Authentication Method (default: "msal")',
    }
    
    all_good = True
    
    # Check required variables
    for var, description in required_vars.items():
        value = os.getenv(var)
        if value:
            # Mask sensitive values
            if 'SECRET' in var:
                masked_value = f"{value[:4]}...{value[-4:]}" if len(value) > 8 else "****"
                print_success(f"{var:25} {description:40} {masked_value}")
            else:
                print_success(f"{var:25} {description:40} {value}")
        else:
            print_error(f"{var:25} {description:40} NOT SET")
            all_good = False
    
    # Check optional variables
    print("\nOptional Variables:")
    for var, description in optional_vars.items():
        value = os.getenv(var)
        if value:
            print_success(f"{var:25} {description:40} {value}")
        else:
            print(f"  {var:25} {description:40} (using default)")
    
    return all_good

def test_imports():
    """Test if required packages are installed"""
    print_header("Checking Python Packages")
    
    packages = {
        'mcp': 'MCP SDK',
        'office365': 'Office365 REST Python Client',
        'msal': 'Microsoft Authentication Library',
        'pydantic': 'Pydantic',
        'dotenv': 'Python Dotenv'
    }
    
    all_good = True
    
    for package, description in packages.items():
        try:
            __import__(package)
            print_success(f"{package:20} {description}")
        except ImportError:
            print_error(f"{package:20} {description} - NOT INSTALLED")
            all_good = False
    
    return all_good

def test_sharepoint_connection():
    """Test actual SharePoint connection"""
    print_header("Testing SharePoint Connection")
    
    try:
        from mcp_sharepoint.auth import create_sharepoint_context
        
        print("Attempting to connect to SharePoint...")
        ctx = create_sharepoint_context()
        
        print("Fetching web information...")
        web = ctx.web.get().execute_query()
        
        print_success("Successfully connected to SharePoint!")
        print(f"\n  Site Title: {web.title}")
        print(f"  Site URL:   {web.url}")
        print(f"  Auth Method: {os.getenv('SHP_AUTH_METHOD', 'msal').upper()}")
        
        return True
        
    except ImportError as e:
        print_error("Failed to import mcp_sharepoint module")
        print(f"  Error: {e}")
        print("  Make sure you've installed the package: pip install -e .")
        return False
        
    except ValueError as e:
        print_error("Configuration Error")
        print(f"  Error: {e}")
        if "Missing required environment variables" in str(e):
            print("\n  Fix: Set all required environment variables")
        elif "Failed to acquire token" in str(e):
            print("\n  Possible fixes:")
            print("    1. Verify your SHP_TENANT_ID is correct")
            print("    2. Check that SHP_ID_APP and SHP_ID_APP_SECRET are correct")
            print("    3. Ensure your Azure AD app has SharePoint API permissions")
            print("    4. Grant admin consent in Azure Portal")
        return False
        
    except Exception as e:
        print_error("Connection Failed")
        print(f"  Error: {e}")
        print("\n  Possible causes:")
        print("    1. Invalid credentials")
        print("    2. Network connectivity issues")
        print("    3. SharePoint site URL is incorrect")
        print("    4. Azure AD app lacks proper permissions")
        print("    5. Admin consent not granted")
        return False

def test_basic_operations():
    """Test basic SharePoint operations"""
    print_header("Testing Basic Operations")
    
    try:
        from mcp_sharepoint.auth import create_sharepoint_context
        
        ctx = create_sharepoint_context()
        doc_lib = os.getenv("SHP_DOC_LIBRARY", "Shared Documents")
        
        # Test listing folders
        print("Testing folder listing...")
        try:
            folder = ctx.web.get_folder_by_server_relative_path(doc_lib)
            folders = folder.folders.get().execute_query()
            print_success(f"Successfully listed folders (found {len(folders)} folders)")
        except Exception as e:
            print_error(f"Failed to list folders: {e}")
            return False
        
        # Test listing files
        print("Testing file listing...")
        try:
            files = folder.files.get().execute_query()
            print_success(f"Successfully listed files (found {len(files)} files)")
        except Exception as e:
            print_error(f"Failed to list files: {e}")
            return False
        
        return True
        
    except Exception as e:
        print_error(f"Operation test failed: {e}")
        return False

def print_summary(results):
    """Print test summary"""
    print_header("Test Summary")
    
    all_passed = all(results.values())
    
    for test, passed in results.items():
        if passed:
            print_success(test)
        else:
            print_error(test)
    
    print(f"\n{'='*60}")
    if all_passed:
        print("  🎉 All tests passed! You're ready to use SharePoint MCP Server")
    else:
        print("  ⚠ Some tests failed. Please fix the issues above.")
    print(f"{'='*60}\n")
    
    return all_passed

def print_next_steps():
    """Print next steps"""
    print("\nNext Steps:")
    print("  1. Review the configuration above")
    print("  2. If all tests passed, integrate with Claude Desktop")
    print("  3. See QUICKSTART.md for Claude Desktop integration")
    print("  4. Use Test_Connection tool in Claude to verify\n")

def main():
    """Main test function"""
    print("\n" + "="*60)
    print("  SharePoint MCP Server - Connection Test")
    print("="*60)
    
    results = {
        "Environment Variables": check_environment_variables(),
        "Python Packages": test_imports(),
    }
    
    # Only test connection if prerequisites are met
    if results["Environment Variables"] and results["Python Packages"]:
        results["SharePoint Connection"] = test_sharepoint_connection()
        
        # Only test operations if connection succeeded
        if results["SharePoint Connection"]:
            results["Basic Operations"] = test_basic_operations()
    else:
        print_warning("\nSkipping connection tests due to failed prerequisites")
    
    # Print summary
    all_passed = print_summary(results)
    
    # Print next steps
    if all_passed:
        print_next_steps()
    else:
        print("\nTroubleshooting:")
        print("  - See QUICKSTART.md for setup instructions")
        print("  - See AZURE_PORTAL_GUIDE.md for Azure AD configuration")
        print("  - See README.md for detailed documentation\n")
    
    # Exit with appropriate code
    sys.exit(0 if all_passed else 1)

if __name__ == "__main__":
    main()
