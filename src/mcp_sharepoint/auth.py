"""
Authentication module for SharePoint MCP Server
Supports modern Azure AD authentication methods
"""
import os
import logging
from typing import Optional
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import msal

logger = logging.getLogger(__name__)


class SharePointAuthenticator:
    """
    Handles authentication to SharePoint using modern Azure AD methods.
    Supports multiple authentication flows for compatibility with new tenants.
    """
    
    def __init__(
        self,
        site_url: str,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        cert_path: Optional[str] = None,
        cert_thumbprint: Optional[str] = None,
    ):
        """
        Initialize SharePoint authenticator.
        
        Args:
            site_url: SharePoint site URL
            client_id: Azure AD application client ID
            client_secret: Azure AD application client secret
            tenant_id: Azure AD tenant ID
            cert_path: Optional path to certificate file for cert-based auth
            cert_thumbprint: Optional certificate thumbprint
        """
        self.site_url = site_url.rstrip("/")
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.cert_path = cert_path
        self.cert_thumbprint = cert_thumbprint
        
    def get_context_with_msal(self) -> ClientContext:
        """
        Get ClientContext using MSAL for modern Azure AD authentication.
        This is the recommended method for new tenants.
        
        Returns:
            Authenticated ClientContext
        """
        def acquire_token():
            """Acquire token using MSAL"""
            authority_url = f'https://login.microsoftonline.com/{self.tenant_id}'
            
            app = msal.ConfidentialClientApplication(
                authority=authority_url,
                client_id=self.client_id,
                client_credential=self.client_secret
            )
            
            # SharePoint requires the site-specific scope
            scopes = [f"{self.site_url}/.default"]
            
            result = app.acquire_token_for_client(scopes=scopes)
            
            if "access_token" not in result:
                error_desc = result.get("error_description", "Unknown error")
                raise ValueError(f"Failed to acquire token: {error_desc}")
            
            return result
        
        ctx = ClientContext(self.site_url).with_access_token(acquire_token)
        logger.info("Successfully authenticated using MSAL (Modern Azure AD)")
        return ctx
    
    def get_context_with_certificate(self) -> ClientContext:
        """
        Get ClientContext using certificate-based authentication.
        This is an alternative modern authentication method.
        
        Returns:
            Authenticated ClientContext
        
        Raises:
            ValueError: If certificate credentials are not provided
        """
        if not self.cert_path or not self.cert_thumbprint:
            raise ValueError(
                "Certificate path and thumbprint are required for cert-based auth"
            )
        
        ctx = ClientContext(self.site_url).with_client_certificate(
            tenant=self.tenant_id,
            client_id=self.client_id,
            thumbprint=self.cert_thumbprint,
            cert_path=self.cert_path
        )
        
        logger.info("Successfully authenticated using certificate")
        return ctx
    
    def get_context_legacy(self) -> ClientContext:
        """
        Get ClientContext using legacy ACS authentication (deprecated).
        This method is included for backwards compatibility but may not work
        with new tenants where ACS app-only is disabled.
        
        Returns:
            Authenticated ClientContext
        """
        logger.warning(
            "Using legacy ACS authentication. This may fail on new tenants. "
            "Consider using MSAL or certificate-based auth instead."
        )
        
        credentials = ClientCredential(self.client_id, self.client_secret)
        ctx = ClientContext(self.site_url).with_credentials(credentials)
        
        return ctx
    
    def get_context(self, auth_method: str = "msal") -> ClientContext:
        """
        Get authenticated ClientContext using the specified method.
        
        Args:
            auth_method: Authentication method to use.
                        Options: "msal" (default), "certificate", "legacy"
        
        Returns:
            Authenticated ClientContext
        
        Raises:
            ValueError: If invalid auth method specified
        """
        auth_methods = {
            "msal": self.get_context_with_msal,
            "certificate": self.get_context_with_certificate,
            "legacy": self.get_context_legacy
        }
        
        if auth_method not in auth_methods:
            raise ValueError(
                f"Invalid auth method: {auth_method}. "
                f"Must be one of: {', '.join(auth_methods.keys())}"
            )
        
        try:
            return auth_methods[auth_method]()
        except Exception as e:
            logger.error(f"Authentication failed with method '{auth_method}': {e}")
            raise


def create_sharepoint_context() -> ClientContext:
    """
    Factory function to create SharePoint context from environment variables.
    Tries modern authentication methods first, falls back to legacy if needed.
    
    Environment variables required:
        - SHP_SITE_URL: SharePoint site URL
        - SHP_ID_APP: Azure AD application client ID
        - SHP_ID_APP_SECRET: Azure AD application client secret
        - SHP_TENANT_ID: Azure AD tenant ID
        
    Optional environment variables:
        - SHP_AUTH_METHOD: Authentication method (msal, certificate, legacy)
        - SHP_CERT_PATH: Path to certificate file
        - SHP_CERT_THUMBPRINT: Certificate thumbprint
    
    Returns:
        Authenticated ClientContext
    
    Raises:
        ValueError: If required environment variables are missing
    """
    # Get required environment variables
    site_url = os.getenv("SHP_SITE_URL")
    client_id = os.getenv("SHP_ID_APP")
    client_secret = os.getenv("SHP_ID_APP_SECRET")
    tenant_id = os.getenv("SHP_TENANT_ID")
    
    # Validate required variables
    missing_vars = []
    if not site_url:
        missing_vars.append("SHP_SITE_URL")
    if not client_id:
        missing_vars.append("SHP_ID_APP")
    if not client_secret:
        missing_vars.append("SHP_ID_APP_SECRET")
    if not tenant_id:
        missing_vars.append("SHP_TENANT_ID")
    
    if missing_vars:
        raise ValueError(
            f"Missing required environment variables: {', '.join(missing_vars)}"
        )
    
    # Get optional environment variables
    auth_method = os.getenv("SHP_AUTH_METHOD", "msal")
    cert_path = os.getenv("SHP_CERT_PATH")
    cert_thumbprint = os.getenv("SHP_CERT_THUMBPRINT")
    
    # Create authenticator
    authenticator = SharePointAuthenticator(
        site_url=site_url,
        client_id=client_id,
        client_secret=client_secret,
        tenant_id=tenant_id,
        cert_path=cert_path,
        cert_thumbprint=cert_thumbprint
    )
    
    # Try to authenticate
    try:
        ctx = authenticator.get_context(auth_method=auth_method)
        logger.info(f"Successfully created SharePoint context using {auth_method} auth")
        return ctx
    except Exception as e:
        logger.error(f"Failed to create SharePoint context: {e}")
        
        # If MSAL failed and we haven't tried legacy, suggest it
        if auth_method == "msal":
            logger.info(
                "MSAL authentication failed. If you're using an older tenant, "
                "you can try setting SHP_AUTH_METHOD=legacy, but note that "
                "legacy ACS authentication is deprecated and may not work on new tenants."
            )
        
        raise
