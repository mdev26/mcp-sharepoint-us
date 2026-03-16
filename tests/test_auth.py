"""
Quick mock test to verify code structure without real credentials
"""
from unittest.mock import Mock, patch, MagicMock
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_imports():
    """Test that all imports work"""
    print("✓ Testing imports...")
    from mcp_sharepoint import auth
    from mcp_sharepoint import __init__
    print("✓ All imports successful")

def test_authenticator_creation():
    """Test that SharePointAuthenticator can be instantiated"""
    print("\n✓ Testing authenticator creation...")
    from mcp_sharepoint.auth import SharePointAuthenticator

    authenticator = SharePointAuthenticator(
        site_url="https://test.sharepoint.us/sites/test",
        client_id="test-client-id",
        client_secret="test-secret",
        tenant_id="test-tenant-id",
        cloud="government"
    )

    print(f"  Cloud: {authenticator.cloud}")
    print(f"  Site URL: {authenticator.site_url}")
    print("✓ Authenticator created successfully")

def test_token_acquisition():
    """Test that get_access_token() returns a token string"""
    print("\n✓ Testing token acquisition...")
    from mcp_sharepoint.auth import SharePointAuthenticator

    with patch('mcp_sharepoint.auth.msal.ConfidentialClientApplication') as mock_msal:
        # Mock MSAL to return a token
        mock_app = Mock()
        mock_app.acquire_token_for_client.return_value = {
            "access_token": "test-token-123",
            "expires_in": 3600
        }
        mock_msal.return_value = mock_app

        authenticator = SharePointAuthenticator(
            site_url="https://test.sharepoint.us/sites/test",
            client_id="test-client-id",
            client_secret="test-secret",
            tenant_id="test-tenant-id",
            cloud="government"
        )

        # Test the get_access_token method
        token = authenticator.get_access_token()

        assert isinstance(token, str), "Token should be a string"
        assert token == "test-token-123", "Token should match the mocked value"
        print("✓ get_access_token() returns token string correctly")

def test_graph_api_scope():
    """Test that Graph API scope is set correctly"""
    print("\n✓ Testing Graph API scope...")

    test_cases = [
        ("government", "https://graph.microsoft.us/.default"),
        ("us", "https://graph.microsoft.us/.default"),
        ("commercial", "https://graph.microsoft.com/.default"),
    ]

    for cloud, expected_scope in test_cases:
        from mcp_sharepoint.auth import SharePointAuthenticator

        with patch('mcp_sharepoint.auth.msal.ConfidentialClientApplication') as mock_msal:
            mock_app = Mock()
            mock_msal.return_value = mock_app

            authenticator = SharePointAuthenticator(
                site_url=f"https://test.sharepoint.{'us' if cloud == 'government' else 'com'}/sites/test",
                client_id="test-client-id",
                client_secret="test-secret",
                tenant_id="test-tenant-id",
                cloud=cloud
            )

            assert authenticator._scopes == [expected_scope], f"Failed for cloud: {cloud}"
            print(f"  {cloud} → {expected_scope} ✓")

    print("✓ Graph API scope configuration working correctly")

def test_package_structure():
    """Test that package can be imported"""
    print("\n✓ Testing package structure...")
    import mcp_sharepoint
    print(f"  Package location: {mcp_sharepoint.__file__}")
    print("✓ Package structure valid")

if __name__ == "__main__":
    print("=" * 60)
    print("Running Mock Tests (No Credentials Required)")
    print("=" * 60)

    try:
        test_imports()
        test_authenticator_creation()
        test_token_acquisition()
        test_graph_api_scope()
        test_package_structure()

        print("\n" + "=" * 60)
        print("✓ ALL TESTS PASSED")
        print("=" * 60)
        print("\nCode structure is valid. Ready for deployment testing with real credentials.")

    except Exception as e:
        print(f"\n✗ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
