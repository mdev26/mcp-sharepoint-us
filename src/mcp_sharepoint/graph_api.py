"""
Microsoft Graph API implementation for SharePoint operations.
Primary API for all SharePoint operations in Azure Government Cloud.
"""
import os
import logging
import asyncio
import socket
import ssl
from typing import Optional, Dict, Any, List
from urllib.parse import urlparse, quote
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

logger = logging.getLogger(__name__)


class GraphAPIClient:
    """
    Microsoft Graph API client for SharePoint operations.
    Primary client for all SharePoint operations, especially in Azure Government Cloud
    where SharePoint REST API may not support app-only authentication.
    """

    def __init__(self, site_url: str, token_callback, document_library_name: str = "Documents"):
        """
        Initialize Graph API client.

        Args:
            site_url: SharePoint site URL (e.g., https://tenant.sharepoint.us/sites/SiteName)
            token_callback: Function that returns access token
            document_library_name: Name of the document library to use (e.g., "Shared Documents1")
        """
        self.site_url = site_url.rstrip("/")
        self.token_callback = token_callback
        self.document_library_name = document_library_name
        self._site_id = None
        self._drive_id = None  # Cache drive ID to avoid repeated API calls

        # Determine Graph API endpoint based on cloud
        if ".sharepoint.us" in site_url:
            self.graph_endpoint = "https://graph.microsoft.us/v1.0"
            logger.info("Using Microsoft Graph US Government endpoint")
        else:
            self.graph_endpoint = "https://graph.microsoft.com/v1.0"
            logger.info("Using Microsoft Graph Commercial endpoint")

        # Create a requests session with retry logic
        self._session = self._create_session()

    def _create_session(self) -> requests.Session:
        """
        Create a requests session with retry logic and connection pooling.
        """
        session = requests.Session()

        # Configure retry strategy for transient errors
        retry_strategy = Retry(
            total=3,  # Total number of retries
            backoff_factor=1,  # Wait 1, 2, 4 seconds between retries
            status_forcelist=[429, 500, 502, 503, 504],  # Retry on these HTTP status codes
            allowed_methods=["HEAD", "GET", "PUT", "DELETE", "OPTIONS", "TRACE", "POST"]
        )

        adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=10, pool_maxsize=10)
        session.mount("http://", adapter)
        session.mount("https://", adapter)

        return session

    def _diagnose_connectivity(self, url: str) -> None:
        """
        Perform detailed connectivity diagnostics for a URL.

        Args:
            url: The URL to diagnose
        """
        parsed = urlparse(url)
        hostname = parsed.hostname
        port = parsed.port or (443 if parsed.scheme == "https" else 80)

        logger.info(f"=== CONNECTIVITY DIAGNOSTICS for {hostname} ===")

        # 1. DNS Resolution
        try:
            logger.info(f"[DNS] Resolving {hostname}...")
            ip_addresses = socket.getaddrinfo(hostname, port, socket.AF_UNSPEC, socket.SOCK_STREAM)
            for family, socktype, proto, canonname, sockaddr in ip_addresses:
                family_name = "IPv4" if family == socket.AF_INET else "IPv6"
                logger.info(f"[DNS] ✓ Resolved to {sockaddr[0]} ({family_name})")
        except socket.gaierror as e:
            logger.error(f"[DNS] ✗ DNS resolution failed: {e}")
            return
        except Exception as e:
            logger.error(f"[DNS] ✗ Unexpected error during DNS resolution: {e}")
            return

        # 2. TCP Connection Test
        try:
            logger.info(f"[TCP] Testing TCP connection to {hostname}:{port}...")
            with socket.create_connection((hostname, port), timeout=10) as sock:
                logger.info(f"[TCP] ✓ TCP connection successful")
                peer_name = sock.getpeername()
                logger.info(f"[TCP] Connected to {peer_name[0]}:{peer_name[1]}")

                # 3. SSL/TLS Test (if HTTPS)
                if parsed.scheme == "https":
                    logger.info(f"[TLS] Testing TLS handshake to {hostname}...")
                    logger.info(f"[TLS] This will attempt to establish encrypted HTTPS connection")
                    context = ssl.create_default_context()
                    try:
                        with context.wrap_socket(sock, server_hostname=hostname) as ssock:
                            logger.info(f"[TLS] ✓ TLS handshake successful")
                            logger.info(f"[TLS] Protocol: {ssock.version()}")
                            cipher = ssock.cipher()
                            if cipher:
                                logger.info(f"[TLS] Cipher: {cipher[0]} (bits: {cipher[2]})")

                            # Get certificate info
                            cert = ssock.getpeercert()
                            if cert:
                                subject = dict(x[0] for x in cert['subject'])
                                logger.info(f"[TLS] Certificate subject: {subject.get('commonName', 'N/A')}")
                                logger.info(f"[TLS] Certificate issuer: {dict(x[0] for x in cert['issuer']).get('organizationName', 'N/A')}")
                    except ssl.SSLError as e:
                        logger.error(f"[TLS] ✗ TLS/SSL handshake failed: {e}")
                        logger.error(f"[TLS] This could indicate:")
                        logger.error(f"[TLS]   - Certificate validation failure")
                        logger.error(f"[TLS]   - TLS version mismatch")
                        logger.error(f"[TLS]   - Cipher suite incompatibility")
                        return
                    except ConnectionResetError as e:
                        logger.error(f"[TLS] ✗ Connection reset during TLS handshake")
                        logger.error(f"[TLS] TCP connection was established BUT connection dropped during TLS negotiation")
                        logger.error(f"[TLS] This indicates:")
                        logger.error(f"[TLS]   - Firewall is doing deep packet inspection (DPI)")
                        logger.error(f"[TLS]   - Firewall is blocking TLS connections to {hostname}")
                        logger.error(f"[TLS]   - SNI (Server Name Indication) filtering is active")
                        logger.error(f"[TLS]")
                        logger.error(f"[TLS] SOLUTION: Ask network team to whitelist {hostname} in firewall")
                        logger.error(f"[TLS] The firewall needs to allow TLS/HTTPS traffic to this endpoint")
                        return
        except socket.timeout:
            logger.error(f"[TCP] ✗ Connection timeout after 10 seconds")
            return
        except ConnectionRefusedError:
            logger.error(f"[TCP] ✗ Connection refused by server")
            return
        except ConnectionResetError:
            logger.error(f"[TCP] ✗ Connection reset by peer during TCP handshake")
            return
        except Exception as e:
            logger.error(f"[TCP] ✗ Connection failed: {type(e).__name__}: {e}")
            return

        # 4. HTTP Basic Connectivity Test
        try:
            logger.info(f"[HTTP] Testing basic HTTP GET to {parsed.scheme}://{hostname}/")
            test_url = f"{parsed.scheme}://{hostname}/"
            response = self._session.get(test_url, timeout=10)
            logger.info(f"[HTTP] ✓ Basic HTTP request successful (status: {response.status_code})")
        except requests.exceptions.RequestException as e:
            logger.error(f"[HTTP] ✗ Basic HTTP request failed: {type(e).__name__}: {e}")

        logger.info(f"=== END DIAGNOSTICS ===\n")

    def _get_headers(self) -> Dict[str, str]:
        """Get authorization headers with access token."""
        token_obj = self.token_callback()
        # Handle both TokenResponse objects and plain strings
        if hasattr(token_obj, 'accessToken'):
            token = token_obj.accessToken
        else:
            token = str(token_obj)

        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }

    def _handle_response(self, response: requests.Response) -> None:
        """
        Handle Graph API response and raise detailed errors if needed.

        Graph API returns errors in format:
        {
          "error": {
            "code": "itemNotFound",
            "message": "The resource could not be found."
          }
        }
        """
        if response.ok:
            return

        try:
            error_data = response.json()
            if "error" in error_data:
                error = error_data["error"]
                code = error.get("code", "Unknown")
                message = error.get("message", "Unknown error")
                raise requests.HTTPError(
                    f"Graph API error [{code}]: {message}",
                    response=response
                )
        except (ValueError, KeyError):
            # If we can't parse the error, fall back to standard handling
            pass

        self._handle_response(response)

    def _get_site_id(self) -> str:
        """
        Get the site ID from the site URL.
        Caches the result for reuse.
        """
        if self._site_id:
            logger.debug(f"Using cached site ID: {self._site_id}")
            return self._site_id

        parsed = urlparse(self.site_url)
        hostname = parsed.netloc
        path = parsed.path.strip("/")

        # For root site: https://tenant.sharepoint.us
        if not path or path == "sites":
            url = f"{self.graph_endpoint}/sites/{hostname}"
        # For subsite: https://tenant.sharepoint.us/sites/SiteName
        else:
            url = f"{self.graph_endpoint}/sites/{hostname}:/{path}"

        logger.info(f"Fetching site ID...")

        try:
            response = self._session.get(url, headers=self._get_headers(), timeout=30)
            self._handle_response(response)

            self._site_id = response.json()["id"]
            logger.info(f"✓ Retrieved site ID")
            return self._site_id

        except requests.exceptions.ConnectionError as e:
            logger.error(f"✗ ConnectionError getting site ID: {e}")

            # Run detailed diagnostics if DEBUG mode is enabled
            if os.getenv("DEBUG", "").lower() in ("true", "1", "yes"):
                logger.error("Running comprehensive diagnostics (DEBUG=true)...")
                self._diagnose_connectivity(url)
            else:
                logger.error("Hint: Set DEBUG=true environment variable for detailed diagnostics")

            raise

        except requests.exceptions.Timeout:
            logger.error(f"✗ Request timeout after 30 seconds", exc_info=True)
            raise

        except requests.exceptions.RequestException as e:
            logger.error(f"✗ Network error getting site ID: {type(e).__name__}: {e}", exc_info=True)
            raise

    def _get_drive_id(self) -> str:
        """
        Get the document library drive ID by name.
        Fetches all drives and finds the one matching self.document_library_name.
        Caches the result for reuse.
        """
        if self._drive_id:
            logger.debug(f"Using cached drive ID: {self._drive_id}")
            return self._drive_id

        site_id = self._get_site_id()
        url = f"{self.graph_endpoint}/sites/{site_id}/drives"

        logger.info(f"Looking for document library: '{self.document_library_name}'")

        try:
            response = self._session.get(url, headers=self._get_headers(), timeout=30)
            self._handle_response(response)

            drives = response.json().get("value", [])

            # Find the drive matching the configured document library name
            for drive in drives:
                if drive.get("name") == self.document_library_name:
                    self._drive_id = drive["id"]
                    logger.info(f"✓ Found document library '{self.document_library_name}'")
                    return self._drive_id

            # If no match found, raise an error with helpful message
            available_drives = [d.get("name", "Unknown") for d in drives]
            error_msg = (
                f"Could not find document library named '{self.document_library_name}'. "
                f"Available drives: {', '.join(available_drives)}"
            )
            logger.error(error_msg)
            raise ValueError(error_msg)

        except requests.exceptions.ConnectionError as e:
            logger.error(f"✗ ConnectionError getting drives: {e}", exc_info=True)
            raise

        except requests.exceptions.RequestException as e:
            logger.error(f"✗ Network error getting drives: {type(e).__name__}: {e}", exc_info=True)
            raise

    def list_folders(self, folder_path: str = "") -> List[Dict[str, Any]]:
        """
        List folders in the specified path.

        Args:
            folder_path: Relative path from document library root

        Returns:
            List of folder objects with name, id, webUrl
        """
        site_id = self._get_site_id()
        drive_id = self._get_drive_id()

        if folder_path:
            encoded_path = quote(folder_path)
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
        else:
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root/children"

        try:
            response = self._session.get(url, headers=self._get_headers(), timeout=30)
            self._handle_response(response)

            items = response.json().get("value", [])
            folders = [
                {
                    "name": item["name"],
                    "id": item["id"],
                    "webUrl": item.get("webUrl", ""),
                }
                for item in items
                if "folder" in item
            ]

            return folders
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error listing folders: {type(e).__name__}: {e}", exc_info=True)
            raise

    def list_documents(self, folder_path: str = "") -> List[Dict[str, Any]]:
        """
        List documents in the specified folder.

        Args:
            folder_path: Relative path from document library root

        Returns:
            List of file objects with name, id, size, webUrl
        """
        site_id = self._get_site_id()
        drive_id = self._get_drive_id()

        if folder_path:
            encoded_path = quote(folder_path)
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
        else:
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root/children"

        try:
            response = self._session.get(url, headers=self._get_headers(), timeout=30)
            self._handle_response(response)

            items = response.json().get("value", [])
            files = [
                {
                    "name": item["name"],
                    "id": item["id"],
                    "size": item.get("size", 0),
                    "webUrl": item.get("webUrl", ""),
                }
                for item in items
                if "file" in item
            ]

            return files
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error listing documents: {type(e).__name__}: {e}", exc_info=True)
            raise

    def get_file_content(self, file_path: str) -> bytes:
        """
        Get the content of a file.

        Args:
            file_path: Relative path to the file

        Returns:
            File content as bytes
        """
        site_id = self._get_site_id()
        drive_id = self._get_drive_id()

        encoded_path = quote(file_path)
        url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/content"

        try:
            response = self._session.get(url, headers=self._get_headers(), timeout=60)
            self._handle_response(response)
            return response.content
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error getting file content: {type(e).__name__}: {e}", exc_info=True)
            raise

    def upload_file(self, folder_path: str, file_name: str, content: bytes) -> Dict[str, Any]:
        """
        Upload a file to SharePoint.

        Args:
            folder_path: Destination folder path
            file_name: Name of the file
            content: File content as bytes

        Returns:
            File metadata
        """
        site_id = self._get_site_id()
        drive_id = self._get_drive_id()

        if folder_path:
            full_path = f"{folder_path}/{file_name}"
        else:
            full_path = file_name

        encoded_path = quote(full_path)
        url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/content"

        headers = self._get_headers()
        headers["Content-Type"] = "application/octet-stream"

        try:
            response = self._session.put(url, headers=headers, data=content, timeout=120)
            self._handle_response(response)
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error uploading file: {type(e).__name__}: {e}", exc_info=True)
            raise

    def delete_file(self, file_path: str) -> None:
        """
        Delete a file.

        Args:
            file_path: Relative path to the file
        """
        site_id = self._get_site_id()
        drive_id = self._get_drive_id()

        encoded_path = quote(file_path)
        url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"

        try:
            response = self._session.delete(url, headers=self._get_headers(), timeout=30)
            self._handle_response(response)
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error deleting file: {type(e).__name__}: {e}", exc_info=True)
            raise

    def create_folder(self, parent_path: str, folder_name: str) -> Dict[str, Any]:
        """
        Create a new folder.

        Args:
            parent_path: Path to parent folder
            folder_name: Name of the new folder

        Returns:
            Folder metadata
        """
        site_id = self._get_site_id()
        drive_id = self._get_drive_id()

        if parent_path:
            encoded_path = quote(parent_path)
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
        else:
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root/children"

        payload = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"
        }

        try:
            response = self._session.post(url, headers=self._get_headers(), json=payload, timeout=30)
            self._handle_response(response)
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error creating folder: {type(e).__name__}: {e}", exc_info=True)
            raise

    def delete_folder(self, folder_path: str) -> None:
        """
        Delete a folder.

        Args:
            folder_path: Relative path to the folder
        """
        site_id = self._get_site_id()
        drive_id = self._get_drive_id()

        encoded_path = quote(folder_path)
        url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"

        try:
            response = self._session.delete(url, headers=self._get_headers(), timeout=30)
            self._handle_response(response)
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error deleting folder: {type(e).__name__}: {e}", exc_info=True)
            raise
