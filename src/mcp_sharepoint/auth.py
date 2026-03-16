"""
Authentication module for SharePoint MCP Server
Supports Azure US Government Cloud and Commercial Cloud
"""
import os
import logging
import time
import random
from typing import Optional
from datetime import datetime, timezone
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.token_response import TokenResponse
import msal

logger = logging.getLogger(__name__)

# Workaround for office365 library datetime comparison bug
# This will be fixed in the next library release
def _patch_datetime_bug():
    """
    Patches the office365 library's datetime comparison issue.
    The library compares timezone-aware datetime.now(timezone.utc) with
    timezone-naive datetime.max, causing a TypeError.
    """
    try:
        from office365.runtime.auth import authentication_context

        # Store the original __init__
        original_init = authentication_context.AuthenticationContext.__init__

        def patched_init(self, *args, **kwargs):
            # Call original init with all arguments
            original_init(self, *args, **kwargs)
            # Make token_expires timezone-aware to prevent comparison errors
            if hasattr(self, '_token_expires') and self._token_expires is not None:
                if self._token_expires.tzinfo is None:
                    # If it's datetime.max (naive), make it timezone-aware
                    if self._token_expires == datetime.max:
                        self._token_expires = datetime.max.replace(tzinfo=timezone.utc)

        authentication_context.AuthenticationContext.__init__ = patched_init
        logger.info("Applied datetime comparison patch for office365 library")

    except Exception as e:
        logger.warning(f"Could not apply datetime patch (may not be needed): {e}")

# Apply the patch when module is imported
_patch_datetime_bug()


class SharePointAuthenticator:
    """
    Handles authentication to SharePoint using modern Azure AD methods.
    Supports both Commercial and US Government clouds.
    """
    
    def __init__(
        self,
        site_url: str,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        cert_path: Optional[str] = None,
        cert_thumbprint: Optional[str] = None,
        cloud: str = "government"  # Default to US Government
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
            cloud: Cloud environment - "commercial" or "government" (default)
        """
        self.site_url = site_url.rstrip("/")
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.cert_path = cert_path
        self.cert_thumbprint = cert_thumbprint
        self.cloud = cloud.lower()

        # Initialize token cache
        self._access_token = None
        self._access_token_exp = 0

        # Set Graph API scope based on cloud environment
        if self.cloud in ("government", "us"):
            self._scopes = ["https://graph.microsoft.us/.default"]
        else:
            self._scopes = ["https://graph.microsoft.com/.default"]

    def get_context_with_msal(self) -> ClientContext:
        """
        Get ClientContext using MSAL for modern Azure AD authentication.
        Uses a cached MSAL app + simple in-memory token cache to avoid repeated
        OIDC discovery calls and reduce connection resets.
        """

        # Build and cache the MSAL app once per authenticator instance
        if not hasattr(self, "_msal_app"):
            if self.cloud in ("government", "us"):
                authority_url = f"https://login.microsoftonline.us/{self.tenant_id}"
                logger.info("Using Azure US Government Cloud endpoints")
            else:
                authority_url = f"https://login.microsoftonline.com/{self.tenant_id}"
                logger.info("Using Azure Commercial Cloud endpoints")

            # Optional: enable MSAL token cache (in-memory). Helps reduce calls.
            self._token_cache = getattr(self, "_token_cache", msal.SerializableTokenCache())

            # For government cloud, disable instance discovery to prevent
            # MSAL from trying to connect to commercial cloud endpoints
            msal_params = {
                "authority": authority_url,
                "client_id": self.client_id,
                "client_credential": self.client_secret,
                "token_cache": self._token_cache,
            }

            # Disable instance discovery for sovereign clouds to avoid
            # hitting login.microsoftonline.com endpoints
            if self.cloud in ("government", "us"):
                msal_params["validate_authority"] = False
                logger.info("Disabled authority validation for government cloud")

            self._msal_app = msal.ConfidentialClientApplication(**msal_params)
            self._authority_url = authority_url

        logger.info(f"Using Graph API scope: {self._scopes[0]}")

        def acquire_token():
            """
            Token callback used by office365 ClientContext.
            Retries transient network errors like ConnectionResetError(104).
            Returns a TokenResponse object with tokenType and accessToken attributes.
            """
            now = int(time.time())
            if self._access_token and now < (self._access_token_exp - 60):
                return TokenResponse(
                    access_token=self._access_token,
                    token_type="Bearer"
                )

            last_err = None
            for attempt in range(1, 6):  # 5 attempts
                try:
                    result = self._msal_app.acquire_token_for_client(scopes=self._scopes)

                    if "access_token" not in result:
                        error_desc = result.get("error_description", "Unknown error")
                        error = result.get("error", "Unknown")
                        raise ValueError(
                            f"Failed to acquire token: {error} - {error_desc}\n"
                            f"Authority: {self._authority_url}\n"
                            f"Scopes: {self._scopes}"
                        )

                    token = result["access_token"]

                    # MSAL returns expires_in (seconds) for client credential tokens
                    expires_in = int(result.get("expires_in", 3600))
                    self._access_token = token
                    self._access_token_exp = int(time.time()) + expires_in

                    logger.info(f"Successfully acquired token for {self.site_url}")
                    # Use from_json to automatically convert MSAL's snake_case to camelCase
                    return TokenResponse.from_json(result)

                except Exception as e:
                    last_err = e
                    # Exponential backoff with jitter
                    sleep_s = min(8.0, (2 ** (attempt - 1)) * 0.5) + random.random() * 0.25
                    logger.warning(
                        f"Token acquisition attempt {attempt}/5 failed: {e}. Retrying in {sleep_s:.2f}s"
                    )
                    time.sleep(sleep_s)

            # If we get here, all retries failed
            raise RuntimeError(f"Token acquisition failed after retries: {last_err}")

        ctx = ClientContext(self.site_url).with_access_token(acquire_token)
        logger.info("Successfully authenticated using MSAL (Modern Azure AD)")
        return ctx

    def get_access_token(self) -> str:
        """
        Get access token directly for use with Microsoft Graph API.
        Uses the same retry logic as get_context_with_msal() but returns just the token string.

        Returns:
            Access token as string

        Raises:
            RuntimeError: If token acquisition fails after retries
        """
        # Initialize MSAL app if not already done
        if not hasattr(self, "_msal_app"):
            if self.cloud in ("government", "us"):
                authority_url = f"https://login.microsoftonline.us/{self.tenant_id}"
            else:
                authority_url = f"https://login.microsoftonline.com/{self.tenant_id}"

            self._token_cache = msal.SerializableTokenCache()

            msal_params = {
                "authority": authority_url,
                "client_id": self.client_id,
                "client_credential": self.client_secret,
                "token_cache": self._token_cache,
            }

            if self.cloud in ("government", "us"):
                msal_params["validate_authority"] = False

            self._msal_app = msal.ConfidentialClientApplication(**msal_params)
            self._authority_url = authority_url

        now = int(time.time())
        if self._access_token and now < (self._access_token_exp - 60):
            return self._access_token

        last_err = None
        for attempt in range(1, 6):  # 5 attempts
            try:
                result = self._msal_app.acquire_token_for_client(scopes=self._scopes)

                if "access_token" not in result:
                    error_desc = result.get("error_description", "Unknown error")
                    error = result.get("error", "Unknown")
                    raise ValueError(
                        f"Failed to acquire token: {error} - {error_desc}\n"
                        f"Authority: {self._authority_url}\n"
                        f"Scopes: {self._scopes}"
                    )

                token = result["access_token"]

                # MSAL returns expires_in (seconds) for client credential tokens
                expires_in = int(result.get("expires_in", 3600))
                self._access_token = token
                self._access_token_exp = int(time.time()) + expires_in

                return token

            except Exception as e:
                last_err = e
                # Exponential backoff with jitter
                sleep_s = min(8.0, (2 ** (attempt - 1)) * 0.5) + random.random() * 0.25
                if attempt < 5:
                    logger.warning(f"Token acquisition failed, retrying in {sleep_s:.2f}s...")
                time.sleep(sleep_s)

        # If we get here, all retries failed
        raise RuntimeError(f"Token acquisition failed after retries: {last_err}")


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
    Automatically detects cloud environment and uses appropriate endpoints.
    
    Environment variables required:
        - SHP_SITE_URL: SharePoint site URL
        - SHP_ID_APP: Azure AD application client ID
        - SHP_ID_APP_SECRET: Azure AD application client secret
        - SHP_TENANT_ID: Azure AD tenant ID
        
    Optional environment variables:
        - SHP_CLOUD: Cloud environment ("commercial" or "government")
                    Auto-detected from site URL if not specified
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
    
    # Detect cloud environment from site URL if not explicitly set
    cloud = os.getenv("SHP_CLOUD", "").lower()
    if not cloud:
        # Auto-detect based on SharePoint URL
        if ".sharepoint.us" in site_url:
            cloud = "government"
            logger.info("Auto-detected Azure US Government Cloud from site URL")
        elif ".sharepoint.com" in site_url:
            cloud = "commercial"
            logger.info("Auto-detected Azure Commercial Cloud from site URL")
        else:
            # Default to government for safety
            cloud = "government"
            logger.warning(
                f"Could not auto-detect cloud from URL: {site_url}. "
                "Defaulting to US Government Cloud. Set SHP_CLOUD explicitly if needed."
            )
    
    # Get optional environment variables
    auth_method = os.getenv("SHP_AUTH_METHOD", "msal")
    cert_path = os.getenv("SHP_CERT_PATH")
    cert_thumbprint = os.getenv("SHP_CERT_THUMBPRINT")
    
    # Log configuration
    logger.info(f"Cloud environment: {cloud}")
    logger.info(f"Authentication method: {auth_method}")
    logger.info(f"Site URL: {site_url}")
    
    # Create authenticator
    authenticator = SharePointAuthenticator(
        site_url=site_url,
        client_id=client_id,
        client_secret=client_secret,
        tenant_id=tenant_id,
        cert_path=cert_path,
        cert_thumbprint=cert_thumbprint,
        cloud=cloud
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
