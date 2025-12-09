"""SharePoint API client for Home Assistant integration."""

import logging
import random
from collections import deque
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

import aiohttp
import msal
from homeassistant.core import HomeAssistant
from homeassistant.helpers.aiohttp_client import async_get_clientsession
from homeassistant.util import dt as dt_util

from .const import (
    AUTHORITY_BASE,
    CONF_CLIENT_ID,
    CONF_CLIENT_SECRET,
    CONF_TENANT_ID,
    DEFAULT_FOLDER_HISTORY_SIZE,
    GRAPH_API_BASE,
    IMAGE_EXTENSIONS,
    SCOPE,
)

_LOGGER = logging.getLogger(__name__)


class SharePointPhotosApiClient:
    """Client for interacting with SharePoint via Microsoft Graph API."""

    def __init__(
        self,
        hass: HomeAssistant,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_url: str,
        library_name: str = "Documents",
        base_folder_path: str = "/Photos",
        recent_history_size: int = DEFAULT_FOLDER_HISTORY_SIZE,
    ) -> None:
        """Initialize the SharePoint client."""
        self.hass = hass
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_url = site_url
        self.library_name = library_name
        self.base_folder_path = base_folder_path
        self._recent_history_size = max(0, recent_history_size or 0)
        self._recent_folder_paths = (
            deque(maxlen=self._recent_history_size) if self._recent_history_size > 0 else None
        )
        
        self._session = async_get_clientsession(hass)
        self._access_token: Optional[str] = None
        self._token_expires: Optional[datetime] = None
        self._site_id: Optional[str] = None
        self._drive_id: Optional[str] = None
        
        # Cache for folder structure
        self._folder_cache: List[Dict[str, Any]] = []
        self._cache_expires: Optional[datetime] = None
        
        # Current folder state (to prevent random changes on every refresh)
        self._current_folder_path: Optional[str] = None
        self._current_folder_data: Optional[Dict[str, Any]] = None

    def _build_display_folder_name(self, folder_path: str) -> str:
        """Return folder name relative to the configured base path."""
        if not folder_path:
            return ""

        normalized_path = folder_path.strip("/")
        normalized_base = (self.base_folder_path or "").strip("/")

        if not normalized_path:
            return ""

        if normalized_base:
            path_parts = normalized_path.split("/")
            base_parts = normalized_base.split("/")

            if path_parts[: len(base_parts)] == base_parts:
                relative_parts = path_parts[len(base_parts) :]
                if relative_parts:
                    return "/".join(relative_parts)

        return normalized_path.split("/")[-1]

    def _filter_recent_folders(self, folders: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Filter out recently used folders when enough choices are available."""
        if not folders or not self._recent_folder_paths:
            return folders

        recent_paths = set(self._recent_folder_paths)
        available = [folder for folder in folders if folder.get("path") not in recent_paths]

        if available:
            filtered_count = len(folders) - len(available)
            if filtered_count:
                _LOGGER.debug(
                    "Excluded %d recently used folders (limit=%d)",
                    filtered_count,
                    self._recent_history_size,
                )
            return available

        _LOGGER.debug(
            "All folders are within the recent history window (%d); allowing reuse",
            self._recent_history_size,
        )
        return folders

    def _record_folder_history(self, folder_path: Optional[str]) -> None:
        """Remember the most recently selected folder to avoid immediate reuse."""
        if not folder_path or not self._recent_folder_paths:
            return

        try:
            self._recent_folder_paths.remove(folder_path)
        except ValueError:
            pass

        self._recent_folder_paths.append(folder_path)
        _LOGGER.debug(
            "Recent folder history updated (%d/%d): %s",
            len(self._recent_folder_paths),
            self._recent_history_size,
            folder_path,
        )

    async def authenticate(self) -> bool:
        """Authenticate with Microsoft Graph API."""
        try:
            # Try direct HTTP authentication first (to avoid MSAL blocking issues)
            if await self._authenticate_direct():
                return True
            
            # Fallback to MSAL if direct method fails
            return await self._authenticate_msal()
                
        except Exception as e:
            _LOGGER.error("Authentication error: %s", str(e))
            return False

    async def _authenticate_direct(self) -> bool:
        """Authenticate using direct HTTP requests to avoid MSAL blocking."""
        try:
            _LOGGER.info("Attempting direct HTTP authentication")
            
            url = f"{AUTHORITY_BASE}/{self.tenant_id}/oauth2/v2.0/token"
            
            data = {
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': 'https://graph.microsoft.com/.default',
                'grant_type': 'client_credentials'
            }
            
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            async with self._session.post(url, data=data, headers=headers) as response:
                if response.status == 200:
                    result = await response.json()
                    if "access_token" in result:
                        self._access_token = result["access_token"]
                        self._token_expires = dt_util.utcnow() + timedelta(seconds=result.get("expires_in", 3600) - 60)
                        _LOGGER.info("Successfully authenticated with direct HTTP method")
                        return True
                else:
                    error_text = await response.text()
                    _LOGGER.warning("Direct HTTP authentication failed: %s - %s", response.status, error_text)
                    return False
                    
        except Exception as e:
            _LOGGER.warning("Direct HTTP authentication error: %s", str(e))
            return False

    async def _authenticate_msal(self) -> bool:
        """Authenticate using MSAL library (fallback method)."""
        try:
            authority = f"{AUTHORITY_BASE}/{self.tenant_id}"
            _LOGGER.info("Attempting MSAL authentication with authority: %s", authority)
            _LOGGER.info("Client ID: %s", self.client_id[:8] + "..." if len(self.client_id) > 8 else "short")
            _LOGGER.info("Tenant ID: %s", self.tenant_id[:8] + "..." if len(self.tenant_id) > 8 else "short")
            
            # Run MSAL in executor to avoid blocking the event loop
            def _get_token():
                app = msal.ConfidentialClientApplication(
                    self.client_id,
                    authority=authority,
                    client_credential=self.client_secret,
                )
                return app.acquire_token_for_client(scopes=SCOPE)
            
            # Execute the blocking MSAL call in a thread pool
            result = await self.hass.async_add_executor_job(_get_token)
            _LOGGER.info("MSAL result keys: %s", list(result.keys()))
            
            if "access_token" in result:
                self._access_token = result["access_token"]
                # Token typically expires in 1 hour
                self._token_expires = dt_util.utcnow() + timedelta(seconds=result.get("expires_in", 3600) - 60)
                _LOGGER.info("Successfully authenticated with Microsoft Graph API using MSAL")
                return True
            else:
                _LOGGER.error("MSAL authentication failed - result: %s", {k: v for k, v in result.items() if k != "access_token"})
                if "error" in result:
                    _LOGGER.error("Error: %s", result.get("error"))
                if "error_description" in result:
                    _LOGGER.error("Error description: %s", result.get("error_description"))
                if "correlation_id" in result:
                    _LOGGER.error("Correlation ID: %s", result.get("correlation_id"))
                return False
                
        except Exception as e:
            _LOGGER.error("MSAL authentication error: %s", str(e))
            return False

    async def _ensure_authenticated(self) -> bool:
        """Ensure we have a valid access token."""
        if not self._access_token or (self._token_expires and dt_util.utcnow() >= self._token_expires):
            return await self.authenticate()
        return True

    async def _get_headers(self) -> Dict[str, str]:
        """Get headers for API requests."""
        if not await self._ensure_authenticated():
            raise Exception("Authentication failed")
        
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }

    async def _make_authenticated_request(self, url: str, max_retries: int = 1) -> tuple[int, dict]:
        """Make an authenticated HTTP request with automatic token refresh on 401."""
        headers = await self._get_headers()
        
        for attempt in range(max_retries + 1):
            try:
                async with self._session.get(url, headers=headers) as response:
                    if response.status == 401 and attempt < max_retries:
                        _LOGGER.warning("Got 401 error, refreshing token and retrying (attempt %d/%d)", attempt + 1, max_retries + 1)
                        # Clear current token to force refresh
                        self._access_token = None
                        self._token_expires = None
                        # Get new headers with fresh token
                        headers = await self._get_headers()
                        continue
                    
                    # Return status and data for successful requests or final attempt
                    if response.status == 200:
                        data = await response.json()
                        return response.status, data
                    else:
                        return response.status, {}
                        
            except Exception as e:
                if attempt < max_retries:
                    _LOGGER.warning("Request failed, retrying (attempt %d/%d): %s", attempt + 1, max_retries + 1, str(e))
                    continue
                raise
        
        # This should never be reached, but just in case
        raise Exception("Failed to make authenticated request after all retries")

    async def _get_site_id(self) -> Optional[str]:
        """Get the SharePoint site ID."""
        if self._site_id:
            return self._site_id

        try:
            headers = await self._get_headers()
            
            # Extract hostname and site path from URL
            if self.site_url.startswith("https://"):
                site_parts = self.site_url.replace("https://", "").split("/", 1)
                hostname = site_parts[0]
                site_path = site_parts[1] if len(site_parts) > 1 else ""
            else:
                raise ValueError("Site URL must start with https://")

            # Get site ID using hostname and path
            if site_path:
                url = f"{GRAPH_API_BASE}/sites/{hostname}:/{site_path}"
            else:
                url = f"{GRAPH_API_BASE}/sites/{hostname}"

            status, data = await self._make_authenticated_request(url)
            if status == 200:
                self._site_id = data.get("id")
                _LOGGER.debug("Found site ID: %s", self._site_id)
                return self._site_id
            else:
                _LOGGER.error("Failed to get site ID: %s", status)
                return None

        except Exception as e:
            _LOGGER.error("Error getting site ID: %s", str(e))
            return None

    async def _get_drive_id(self) -> Optional[str]:
        """Get the drive ID for the specified library."""
        if self._drive_id:
            return self._drive_id

        site_id = await self._get_site_id()
        if not site_id:
            return None

        try:
            headers = await self._get_headers()
            url = f"{GRAPH_API_BASE}/sites/{site_id}/drives"
            
            # URL decode the library name in case it was URL encoded
            import urllib.parse
            decoded_library_name = urllib.parse.unquote(self.library_name)
            _LOGGER.info("Looking for library: '%s' (decoded: '%s')", self.library_name, decoded_library_name)

            status, data = await self._make_authenticated_request(url)
            if status == 200:
                _LOGGER.info("Available drives/libraries:")
                
                # Log all available drives for debugging
                for drive in data.get("value", []):
                    drive_name = drive.get("name", "")
                    drive_id = drive.get("id", "")
                    _LOGGER.info("  - Name: '%s', ID: %s", drive_name, drive_id[:20] + "..." if len(drive_id) > 20 else drive_id)
                    
                    # Try multiple matching strategies
                    if (drive_name == self.library_name or 
                        drive_name == decoded_library_name or
                        drive_name.lower() == self.library_name.lower() or
                        drive_name.lower() == decoded_library_name.lower()):
                        self._drive_id = drive.get("id")
                        _LOGGER.info("Found matching drive: '%s' with ID: %s", drive_name, self._drive_id[:20] + "..." if len(self._drive_id) > 20 else self._drive_id)
                        return self._drive_id
                
                # If no exact match, try partial matching for common variations
                for drive in data.get("value", []):
                    drive_name = drive.get("name", "")
                    if ("document" in drive_name.lower() and "document" in self.library_name.lower()) or \
                       ("shared" in drive_name.lower() and ("shared" in self.library_name.lower() or "freigegebene" in self.library_name.lower())):
                        self._drive_id = drive.get("id")
                        _LOGGER.info("Found partial match drive: '%s' with ID: %s", drive_name, self._drive_id[:20] + "..." if len(self._drive_id) > 20 else self._drive_id)
                        return self._drive_id
                
                _LOGGER.error("Library '%s' (decoded: '%s') not found in available drives", self.library_name, decoded_library_name)
                return None
            else:
                _LOGGER.error("Failed to get drives: %s", status)
                return None

        except Exception as e:
            _LOGGER.error("Error getting drive ID: %s", str(e))
            return None

    async def get_photo_folders(self, force_refresh: bool = False) -> List[Dict[str, Any]]:
        """Get all folders containing photos from SharePoint."""
        _LOGGER.info("Starting photo folder scan (force_refresh=%s)", force_refresh)
        
        # Check cache first
        if not force_refresh and self._folder_cache and self._cache_expires and dt_util.utcnow() < self._cache_expires:
            _LOGGER.info("Using cached folder data (%d folders)", len(self._folder_cache))
            return self._folder_cache

        drive_id = await self._get_drive_id()
        if not drive_id:
            _LOGGER.error("No drive ID available for folder scanning")
            return []

        try:
            _LOGGER.info("Starting recursive folder scan from path: %s", self.base_folder_path)
            folders = []
            await self._scan_folders_recursive(drive_id, self.base_folder_path, folders)
            
            # Cache the results for 1 hour
            self._folder_cache = folders
            self._cache_expires = dt_util.utcnow() + timedelta(hours=1)
            
            _LOGGER.info("Folder scan complete: found %d photo folders", len(folders))
            return folders

        except Exception as e:
            _LOGGER.error("Error getting photo folders: %s", str(e))
            import traceback
            _LOGGER.error("Traceback: %s", traceback.format_exc())
            return []

    async def _scan_folders_recursive(self, drive_id: str, folder_path: str, folders: List[Dict[str, Any]]) -> None:
        """Recursively scan folders for photos."""
        _LOGGER.debug("Scanning folder: %s", folder_path)
        try:
            headers = await self._get_headers()
            url = f"{GRAPH_API_BASE}/drives/{drive_id}/root:{folder_path}:/children"

            status, data = await self._make_authenticated_request(url)
            if status == 200:
                # Check if current folder has photos
                has_photos = False
                subfolders = []
                
                items = data.get("value", [])
                _LOGGER.debug("Found %d items in folder %s", len(items), folder_path)
                
                for item in items:
                    if item.get("folder"):
                        # It's a subfolder
                        subfolders.append(item["name"])
                    elif item.get("file"):
                        # It's a file, check if it's an image
                        file_name = item.get("name", "").lower()
                        if any(file_name.endswith(ext) for ext in IMAGE_EXTENSIONS):
                            has_photos = True

                # If current folder has photos, add it to the list
                if has_photos:
                    folders.append({
                        "name": self._build_display_folder_name(folder_path),
                        "path": folder_path,
                        "full_path": folder_path,
                    })
                    _LOGGER.debug("Added photo folder: %s", folder_path)

                # Recursively scan subfolders
                _LOGGER.debug("Scanning %d subfolders in %s", len(subfolders), folder_path)
                for subfolder in subfolders:
                    subfolder_path = f"{folder_path}/{subfolder}"
                    await self._scan_folders_recursive(drive_id, subfolder_path, folders)

            elif status == 404:
                _LOGGER.warning("Folder not found: %s", folder_path)
            else:
                _LOGGER.error("Error scanning folder %s: %s", folder_path, status)

        except Exception as e:
            _LOGGER.error("Error scanning folder %s: %s", folder_path, str(e))
            import traceback
            _LOGGER.error("Traceback: %s", traceback.format_exc())

    async def get_folder_photos(self, folder_path: str) -> List[Dict[str, Any]]:
        """Get all photos from a specific folder."""
        drive_id = await self._get_drive_id()
        if not drive_id:
            return []

        try:
            headers = await self._get_headers()
            # Add expand parameter to get thumbnails
            url = f"{GRAPH_API_BASE}/drives/{drive_id}/root:{folder_path}:/children?$expand=thumbnails"

            photos = []
            status, data = await self._make_authenticated_request(url)
            if status == 200:
                for item in data.get("value", []):
                    if item.get("file"):
                        file_name = item.get("name", "").lower()
                        if any(file_name.endswith(ext) for ext in IMAGE_EXTENSIONS):
                            # Try to get thumbnail URL first (better for browser display)
                            thumbnail_url = None
                            thumbnails = item.get("thumbnails", [])
                            if thumbnails:
                                # Get the largest available thumbnail
                                for thumbnail_set in thumbnails:
                                    if "large" in thumbnail_set:
                                        thumbnail_url = thumbnail_set["large"].get("url")
                                        break
                                    elif "medium" in thumbnail_set:
                                        thumbnail_url = thumbnail_set["medium"].get("url")
                                        break
                                    elif "small" in thumbnail_set:
                                        thumbnail_url = thumbnail_set["small"].get("url")
                                        break
                            
                            # Fallback URLs
                            download_url = item.get("@microsoft.graph.downloadUrl")
                            web_url = item.get("webUrl")
                            
                            if download_url:
                                # Create proxy URL for better browser compatibility
                                # We'll use the photo index as the image ID
                                photo_index = len(photos)
                                
                                # Prioritize thumbnail over download URL for shorter URLs
                                display_url = thumbnail_url if thumbnail_url else download_url
                                
                                photos.append({
                                    "name": item["name"],
                                    "url": display_url,  # Use thumbnail primarily, download as fallback
                                    "proxy_url": f"/api/sharepoint_photos/image/{{entry_id}}/{photo_index}",  # Placeholder for entry_id
                                    "thumbnail_url": thumbnail_url,
                                    "download_url": download_url,
                                    "web_url": web_url,
                                    "size": item.get("size", 0),
                                    "modified": item.get("lastModifiedDateTime"),
                                    "index": photo_index,
                                })
                                _LOGGER.debug("Added photo: %s (using %s)", 
                                              item["name"], 
                                              "thumbnail" if thumbnail_url else "download URL")
                            else:
                                _LOGGER.warning("No download URL found for photo: %s", item["name"])

            _LOGGER.debug("Found %d photos in folder %s", len(photos), folder_path)
            return photos

        except Exception as e:
            _LOGGER.error("Error getting photos from folder %s: %s", folder_path, str(e))
            return []

    async def get_random_photo_folder(self) -> Optional[Dict[str, Any]]:
        """Get a random photo folder with its images - this always selects a NEW folder."""
        folders = await self.get_photo_folders()
        if not folders:
            return None

        # Select a random folder while avoiding recently used ones when possible
        candidate_folders = self._filter_recent_folders(folders)
        selected_folder = random.choice(candidate_folders)
        _LOGGER.info("Selected random folder: %s", selected_folder["path"])
        
        # Get photos from the selected folder
        photos = await self.get_folder_photos(selected_folder["path"])
        
        folder_data = {
            "folder_name": self._build_display_folder_name(selected_folder["path"]),
            "folder_path": selected_folder["path"],
            "photos": photos,
            "photo_count": len(photos),
            "last_updated": dt_util.utcnow().isoformat(),
        }
        
        # Update current folder state
        self._current_folder_path = selected_folder["path"]
        self._current_folder_data = folder_data
        self._record_folder_history(selected_folder["path"])
        
        return folder_data

    async def async_get_random_folder_photos(self, force_new_folder: bool = False) -> Optional[Dict[str, Any]]:
        """Get current folder photos, or select a new random folder if needed."""
        # If we have a current folder and we're not forcing a new one, refresh the current folder
        if self._current_folder_path and not force_new_folder:
            _LOGGER.debug("Refreshing current folder: %s", self._current_folder_path)
            try:
                photos = await self.get_folder_photos(self._current_folder_path)
                
                folder_data = {
                    "folder_name": self._build_display_folder_name(self._current_folder_path),
                    "folder_path": self._current_folder_path,
                    "photos": photos,
                    "photo_count": len(photos),
                    "last_updated": dt_util.utcnow().isoformat(),
                }
                
                # Update cached data
                self._current_folder_data = folder_data
                return folder_data
                
            except Exception as e:
                _LOGGER.warning("Failed to refresh current folder %s: %s", self._current_folder_path, str(e))
                # Fall back to selecting a new folder
        
        # Select a new random folder
        _LOGGER.info("Selecting new random folder (force_new_folder=%s)", force_new_folder)
        return await self.get_random_photo_folder()

    async def select_specific_folder(self, folder_path: str) -> Optional[Dict[str, Any]]:
        """Select a specific folder and get its photos."""
        try:
            photos = await self.get_folder_photos(folder_path)
            
            folder_data = {
                "folder_name": self._build_display_folder_name(folder_path),
                "folder_path": folder_path,
                "photos": photos,
                "photo_count": len(photos),
                "last_updated": dt_util.utcnow().isoformat(),
            }
            
            # Update current folder state
            self._current_folder_path = folder_path
            self._current_folder_data = folder_data
            self._record_folder_history(folder_path)
            _LOGGER.info("Selected specific folder: %s", folder_path)
            
            return folder_data
        except Exception as e:
            _LOGGER.error("Error selecting folder %s: %s", folder_path, str(e))
            return None

    async def fetch_image_content(self, download_url: str) -> tuple[bytes, str, int]:
        """Fetch image content from SharePoint with automatic token refresh."""
        try:
            # For SharePoint download URLs, we don't need to add our own auth headers
            # as they contain their own auth tokens. But if they're expired, we need 
            # to refresh the photo data to get new URLs.
            async with self._session.get(download_url) as response:
                if response.status == 401:
                    _LOGGER.warning("Download URL expired (401), this requires refreshing photo data")
                    # Clear our access token to force re-authentication on next API call
                    self._access_token = None
                    self._token_expires = None
                    # Return the error info so the caller can handle it appropriately
                    return b"", "", 401
                elif response.status == 200:
                    content = await response.read()
                    content_type = response.headers.get('content-type', 'image/jpeg')
                    _LOGGER.debug("Successfully fetched image: %d bytes", len(content))
                    return content, content_type, 200
                else:
                    _LOGGER.error("Failed to fetch image: HTTP %d", response.status)
                    return b"", "", response.status
                    
        except Exception as e:
            _LOGGER.error("Error fetching image content: %s", str(e))
            return b"", "", 500

    async def test_connection(self) -> bool:
        """Test the connection to SharePoint."""
        try:
            _LOGGER.info("Testing SharePoint connection...")
            
            # Step 1: Test authentication
            _LOGGER.info("Step 1: Testing authentication...")
            if not await self.authenticate():
                _LOGGER.error("Authentication failed")
                return False
            _LOGGER.info("Authentication successful")
            
            # Step 2: Test site access
            _LOGGER.info("Step 2: Testing site access for URL: %s", self.site_url)
            site_id = await self._get_site_id()
            if not site_id:
                _LOGGER.error("Failed to get site ID for: %s", self.site_url)
                return False
            _LOGGER.info("Site ID found: %s", site_id[:20] + "..." if len(site_id) > 20 else site_id)
            
            # Step 3: Test library access
            _LOGGER.info("Step 3: Testing library access for: %s", self.library_name)
            drive_id = await self._get_drive_id()
            if not drive_id:
                _LOGGER.error("Failed to get drive ID for library: %s", self.library_name)
                return False
            _LOGGER.info("Drive ID found: %s", drive_id[:20] + "..." if len(drive_id) > 20 else drive_id)
            
            _LOGGER.info("All connection tests passed!")
            return True
            
        except Exception as e:
            _LOGGER.error("Connection test failed with exception: %s", str(e))
            import traceback
            _LOGGER.error("Traceback: %s", traceback.format_exc())
            return False