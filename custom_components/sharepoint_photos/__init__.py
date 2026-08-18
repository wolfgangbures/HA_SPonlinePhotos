"""The SharePoint Photos integration."""
from __future__ import annotations

import asyncio
import hashlib
import logging
import random
import time
from datetime import timedelta
from typing import Any

from homeassistant.config_entries import ConfigEntry
from homeassistant.const import Platform
from homeassistant.core import HomeAssistant, callback
from homeassistant.helpers.update_coordinator import DataUpdateCoordinator, UpdateFailed
from homeassistant.components.http import HomeAssistantView
from homeassistant.helpers.event import async_track_time_interval

from .api import SharePointPhotosApiClient
from .const import (
    CONF_BASE_FOLDER_PATH,
    CONF_FOLDER_HISTORY_SIZE,
    CONF_MIN_PHOTO_COUNT,
    CONF_ROTATION_INTERVAL_SECONDS,
    CONF_LIBRARY_NAME,
    DEFAULT_BASE_FOLDER_PATH,
    DEFAULT_FOLDER_HISTORY_SIZE,
    DEFAULT_MIN_PHOTO_COUNT,
    DEFAULT_ROTATION_INTERVAL_SECONDS,
    DEFAULT_LIBRARY_NAME,
    DOMAIN,
)

PLATFORMS: list[Platform] = [Platform.SENSOR, Platform.IMAGE]

_LOGGER = logging.getLogger(__name__)

_DOMAIN_SERVICES_REGISTERED = "_services_registered"
_SETUP_LOCKS_KEY = "_setup_locks"
_SETUP_DONE_PREFIX = "_setup_done_"
_UPDATE_LISTENER_PREFIX = "_update_listener_"


def _iter_entry_ids(domain_data: dict[str, Any]) -> list[str]:
    """Return real config-entry IDs stored under hass.data[DOMAIN]."""
    return [key for key in domain_data.keys() if not key.startswith("_")]


def _resolve_target_coordinator(hass: HomeAssistant, requested_entry_id: str | None):
    """Resolve the coordinator for a service call."""
    domain_data = hass.data.get(DOMAIN, {})

    if requested_entry_id:
        return domain_data.get(requested_entry_id)

    entry_ids = _iter_entry_ids(domain_data)
    if not entry_ids:
        return None

    return domain_data.get(entry_ids[0])


class SharePointImageProxyView(HomeAssistantView):
    """Proxy view for SharePoint images to handle authentication."""
    
    url = "/api/sharepoint_photos/image/{entry_id}/{image_id}"
    name = "api:sharepoint_photos:image"
    requires_auth = False  # We'll handle auth internally

    def __init__(self, hass: HomeAssistant):
        """Initialize the proxy view."""
        self.hass = hass
        # Cache last successful image bytes per entry/image key so transient
        # SharePoint URL expiry does not surface as broken media in dashboards.
        self._last_success: dict[str, dict[str, Any]] = {}

    @staticmethod
    def _cache_key(entry_id: str, image_id: str) -> str:
        """Build the in-memory cache key for a proxied image."""
        return f"{entry_id}:{image_id}"

    @staticmethod
    def _normalize_content_type(content_type: str | None) -> str:
        """Return a browser-safe image content type."""
        if content_type and content_type.lower().startswith("image/"):
            return content_type
        return "image/jpeg"

    def _build_image_response(self, content: bytes, content_type: str, etag: str, include_body: bool = True):
        """Create a consistent HTTP response for image consumers."""
        from aiohttp import web

        headers = {
            "Cache-Control": "public, max-age=30, must-revalidate",
            "Content-Length": str(len(content)),
            "ETag": etag,
            "Access-Control-Allow-Origin": "*",
            "Content-Disposition": "inline",
        }

        if include_body:
            return web.Response(body=content, content_type=content_type, headers=headers)
        return web.Response(status=200, content_type=content_type, headers=headers)

    async def _proxy_image(self, entry_id: str, image_id: str, include_body: bool = True):
        """Fetch and proxy a SharePoint image with stale-cache fallback."""
        from aiohttp import web

        
        _LOGGER.debug("Proxy request received: entry_id=%s, image_id=%s", entry_id, image_id)
        cache_key = self._cache_key(entry_id, image_id)
        stale = self._last_success.get(cache_key)
        
        try:
            # Get the coordinator for this entry
            coordinator = self.hass.data.get(DOMAIN, {}).get(entry_id)
            if not coordinator:
                _LOGGER.error("Coordinator not found for entry_id: %s", entry_id)
                return web.Response(status=404, text="Integration not found")
            
            # Find the image in the current data
            data = coordinator.data
            if not data or not data.get("photos"):
                _LOGGER.error("No photos available in coordinator data")
                return web.Response(status=404, text="No photos available")
            
            # Find the photo by ID (using index as ID)
            try:
                photo_index = int(image_id)
                photos = data["photos"]
                if photo_index < 0:
                    _LOGGER.error("Negative photo index %d", photo_index)
                    return web.Response(status=400, text="Invalid photo ID")
                if photo_index >= len(photos):
                    original_index = photo_index
                    photo_index = photo_index % len(photos)
                    _LOGGER.debug(
                        "Photo index %d out of range for %d photos, remapped to %d",
                        original_index,
                        len(photos),
                        photo_index,
                    )
                
                photo = photos[photo_index]
                download_url = photo.get("download_url")
                if not download_url:
                    _LOGGER.error("No download URL available for photo at index %d", photo_index)
                    return web.Response(status=404, text="Photo URL not available")
                
                _LOGGER.debug("Fetching image from: %s", download_url[:100])
                
            except (ValueError, IndexError) as e:
                _LOGGER.error("Invalid photo ID '%s': %s", image_id, str(e))
                return web.Response(status=400, text="Invalid photo ID")
            
            # Fetch the image from SharePoint using the API client
            coordinator = self.hass.data[DOMAIN][entry_id]
            api_client = coordinator._api_client
            
            content, content_type, status_code = await api_client.fetch_image_content(download_url)
            
            if status_code in (401, 403):
                # Token expired, try to refresh the data and get new URLs
                _LOGGER.info("Image URL expired (status=%d), refreshing photo data...", status_code)
                await coordinator.async_request_refresh()
                
                # Get updated data
                updated_data = coordinator.data
                if updated_data and updated_data.get("photos"):
                    updated_photos = updated_data["photos"]
                    # Try to find a photo with the same name first
                    original_photo_name = photo.get("name", "")
                    updated_photo = None
                    
                    # First, try to find the same photo by name
                    for up in updated_photos:
                        if up.get("name") == original_photo_name:
                            updated_photo = up
                            _LOGGER.debug("Found same photo by name: %s", original_photo_name)
                            break
                    
                    # If not found by name, try the same index if it exists
                    if not updated_photo and photo_index < len(updated_photos):
                        updated_photo = updated_photos[photo_index]
                        _LOGGER.debug("Using photo at same index %d", photo_index)
                    
                    # If still not found, use the first photo
                    if not updated_photo and updated_photos:
                        updated_photo = updated_photos[0]
                        _LOGGER.debug("Using first photo as fallback")
                    
                    if updated_photo:
                        updated_download_url = updated_photo.get("download_url")
                        if updated_download_url and updated_download_url != download_url:
                            _LOGGER.debug("Retrying with refreshed URL")
                            content, content_type, status_code = await api_client.fetch_image_content(updated_download_url)
                        else:
                            _LOGGER.warning("Refreshed photo has same download URL, token refresh may have failed")
                    else:
                        _LOGGER.error("No photos available after refresh")
                else:
                    _LOGGER.error("No photo data available after refresh")
                
            if status_code == 200 and content:
                normalized_content_type = self._normalize_content_type(content_type)
                etag = hashlib.md5(content).hexdigest()  # nosec - weak hash is fine for cache validation
                self._last_success[cache_key] = {
                    "content": content,
                    "content_type": normalized_content_type,
                    "etag": etag,
                }
                _LOGGER.debug("Successfully proxied image: %d bytes, type: %s", len(content), normalized_content_type)
                return self._build_image_response(content, normalized_content_type, etag, include_body=include_body)

            if stale:
                _LOGGER.warning(
                    "Returning stale cached image after fetch failed: HTTP %d",
                    status_code,
                )
                response = self._build_image_response(
                    stale["content"],
                    stale["content_type"],
                    stale["etag"],
                    include_body=include_body,
                )
                response.headers["X-SharePoint-Proxy"] = "stale-cache"
                return response

            else:
                _LOGGER.error("Failed to fetch image from SharePoint: HTTP %d", status_code)
                return web.Response(status=status_code if status_code else 500, text="Failed to fetch image")
                        
        except Exception as e:
            _LOGGER.error("Error proxying SharePoint image: %s", str(e))
            if stale:
                response = self._build_image_response(
                    stale["content"],
                    stale["content_type"],
                    stale["etag"],
                    include_body=include_body,
                )
                response.headers["X-SharePoint-Proxy"] = "stale-cache-exception"
                return response
            return web.Response(status=500, text="Internal server error")

    async def get(self, request, entry_id: str, image_id: str):
        """Proxy SharePoint image requests."""
        return await self._proxy_image(entry_id, image_id, include_body=True)

    async def head(self, request, entry_id: str, image_id: str):
        """Handle HEAD requests for image metadata compatibility."""
        return await self._proxy_image(entry_id, image_id, include_body=False)


class SharePointCurrentImageView(HomeAssistantView):
    """Serve the currently cached integration image via a stable URL."""

    url = "/api/sharepoint_photos/current/{entry_id}"
    name = "api:sharepoint_photos:current_image"
    requires_auth = False

    def __init__(self, hass: HomeAssistant):
        """Initialize the current-image view."""
        self.hass = hass
        self._last_success: dict[str, dict[str, Any]] = {}

    @staticmethod
    def _normalize_content_type(content_type: str | None) -> str:
        """Return a browser-safe image content type."""
        if content_type and content_type.lower().startswith("image/"):
            return content_type
        return "image/jpeg"

    @staticmethod
    def _build_response(
        content: bytes,
        content_type: str,
        etag: str,
        include_body: bool = True,
        cache_control: str = "no-store",
    ):
        """Build a consistent image response."""
        from aiohttp import web

        headers = {
            "Cache-Control": cache_control,
            "Content-Length": str(len(content)),
            "Content-Disposition": "inline",
            "Access-Control-Allow-Origin": "*",
            "ETag": etag,
        }

        if include_body:
            return web.Response(body=content, content_type=content_type, headers=headers)
        return web.Response(status=200, content_type=content_type, headers=headers)

    async def _handle(self, entry_id: str, include_body: bool = True):
        """Return the current cached image."""
        from aiohttp import web

        stale = self._last_success.get(entry_id)

        try:
            coordinator = self.hass.data.get(DOMAIN, {}).get(entry_id)
            if not coordinator:
                if stale:
                    response = self._build_response(
                        stale["content"],
                        stale["content_type"],
                        stale["etag"],
                        include_body=include_body,
                        cache_control="public, max-age=30, must-revalidate",
                    )
                    response.headers["X-SharePoint-Current"] = "stale-cache-no-coordinator"
                    return response
                return web.Response(status=404, text="Integration not found")

            await coordinator.async_ensure_recent_image()
            content, content_type = await coordinator.async_get_or_load_current_image()
            if content:
                normalized_content_type = self._normalize_content_type(content_type)
                etag = hashlib.md5(content).hexdigest()  # nosec - weak hash is fine for cache validation
                self._last_success[entry_id] = {
                    "content": content,
                    "content_type": normalized_content_type,
                    "etag": etag,
                }
                return self._build_response(
                    content,
                    normalized_content_type,
                    etag,
                    include_body=include_body,
                )

            if stale:
                response = self._build_response(
                    stale["content"],
                    stale["content_type"],
                    stale["etag"],
                    include_body=include_body,
                    cache_control="public, max-age=30, must-revalidate",
                )
                response.headers["X-SharePoint-Current"] = "stale-cache-empty-current"
                return response

            return web.Response(status=503, text="Current image not available")
        except Exception as exc:
            _LOGGER.error("Error serving current image for entry %s: %s", entry_id, exc)
            if stale:
                response = self._build_response(
                    stale["content"],
                    stale["content_type"],
                    stale["etag"],
                    include_body=include_body,
                    cache_control="public, max-age=30, must-revalidate",
                )
                response.headers["X-SharePoint-Current"] = "stale-cache-exception"
                return response
            return web.Response(status=500, text="Internal server error")

    async def get(self, request, entry_id: str):
        """Return current image body."""
        return await self._handle(entry_id, include_body=True)

    async def head(self, request, entry_id: str):
        """Return current image metadata."""
        return await self._handle(entry_id, include_body=False)


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Set up this integration using UI."""
    if hass.data.get(DOMAIN) is None:
        hass.data.setdefault(DOMAIN, {})

    domain_data = hass.data[DOMAIN]

    setup_locks = domain_data.setdefault(_SETUP_LOCKS_KEY, {})
    entry_lock = setup_locks.get(entry.entry_id)
    if entry_lock is None:
        entry_lock = asyncio.Lock()
        setup_locks[entry.entry_id] = entry_lock

    async with entry_lock:
        setup_done_key = f"{_SETUP_DONE_PREFIX}{entry.entry_id}"
        if domain_data.get(setup_done_key):
            _LOGGER.debug("Setup already completed for entry %s, skipping duplicate setup", entry.entry_id)
            return True

        tenant_id = entry.data.get("tenant_id")
        client_id = entry.data.get("client_id")
        # Use client_secret from options if updated, otherwise fall back to initial config
        client_secret = entry.options.get("client_secret") or entry.data.get("client_secret")
        site_url = entry.data.get("site_url")
        library_name = entry.options.get(
            CONF_LIBRARY_NAME,
            entry.data.get(CONF_LIBRARY_NAME, DEFAULT_LIBRARY_NAME),
        )
        base_folder_path = entry.options.get(
            CONF_BASE_FOLDER_PATH,
            entry.data.get(CONF_BASE_FOLDER_PATH, DEFAULT_BASE_FOLDER_PATH),
        )
        recent_history_size = entry.options.get(
            CONF_FOLDER_HISTORY_SIZE,
            entry.data.get(CONF_FOLDER_HISTORY_SIZE, DEFAULT_FOLDER_HISTORY_SIZE),
        )
        min_photos_per_folder = entry.options.get(
            CONF_MIN_PHOTO_COUNT,
            entry.data.get(CONF_MIN_PHOTO_COUNT, DEFAULT_MIN_PHOTO_COUNT),
        )
        rotation_interval_seconds = entry.options.get(
            CONF_ROTATION_INTERVAL_SECONDS,
            entry.data.get(CONF_ROTATION_INTERVAL_SECONDS, DEFAULT_ROTATION_INTERVAL_SECONDS),
        )

        client = SharePointPhotosApiClient(
            hass=hass,
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            site_url=site_url,
            library_name=library_name,
            base_folder_path=base_folder_path,
            recent_history_size=recent_history_size,
            min_photos_per_folder=min_photos_per_folder,
        )

        coordinator = SharePointPhotosDataUpdateCoordinator(
            hass,
            client=client,
            entry_id=entry.entry_id,
            rotation_interval_seconds=rotation_interval_seconds,
        )

        # Set empty placeholder data so entities are available immediately,
        # then schedule the actual folder scan as a background task.
        coordinator.async_set_updated_data(coordinator.build_initial_state())
        previous_coordinator = domain_data.get(entry.entry_id)
        if previous_coordinator and previous_coordinator is not coordinator:
            try:
                previous_coordinator.stop_rotation_timer()
            except Exception:
                _LOGGER.debug("Failed to stop previous coordinator timer for entry %s", entry.entry_id)
        domain_data[entry.entry_id] = coordinator

        async def _deferred_first_refresh(_=None):
            """Run the first full folder scan after startup completes."""
            _LOGGER.info("Starting deferred first refresh for SharePoint Photos")
            await coordinator.async_request_refresh()

        hass.async_create_task(_deferred_first_refresh())

        # Register the image proxy view (only if not already registered)
        if not hasattr(hass.http, '_sharepoint_photos_proxy_registered'):
            hass.http.register_view(SharePointImageProxyView(hass))
            hass.http._sharepoint_photos_proxy_registered = True
            _LOGGER.debug("Registered SharePoint Photos image proxy view")

        if not hasattr(hass.http, '_sharepoint_photos_current_view_registered'):
            hass.http.register_view(SharePointCurrentImageView(hass))
            hass.http._sharepoint_photos_current_view_registered = True
            _LOGGER.debug("Registered SharePoint Photos current image view")

        await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)

        listener_key = f"{_UPDATE_LISTENER_PREFIX}{entry.entry_id}"
        old_listener_unsub = domain_data.get(listener_key)
        if old_listener_unsub:
            try:
                old_listener_unsub()
            except Exception:
                _LOGGER.debug("Failed to remove previous update listener for entry %s", entry.entry_id)

        listener_unsub = entry.add_update_listener(async_reload_entry)
        entry.async_on_unload(listener_unsub)
        domain_data[listener_key] = listener_unsub

        # Register domain services only once across reloads.
        if not domain_data.get(_DOMAIN_SERVICES_REGISTERED):

            async def handle_refresh_photos(call):
                """Handle the refresh photos service call - switches to a NEW random folder."""
                requested_entry_id = call.data.get("entry_id")
                target = _resolve_target_coordinator(hass, requested_entry_id)
                if not target:
                    _LOGGER.warning("No coordinator available for refresh_photos service")
                    return
                await target.async_refresh_new_folder()

            async def handle_select_folder(call):
                """Handle the select folder service call."""
                folder_path = call.data.get("folder_path")
                if not folder_path:
                    return

                requested_entry_id = call.data.get("entry_id")
                target = _resolve_target_coordinator(hass, requested_entry_id)
                if not target:
                    _LOGGER.warning("No coordinator available for select_folder service")
                    return

                await target.client.select_specific_folder(folder_path)
                await target.async_request_refresh()

            async def handle_refresh_token(call):
                """Handle the refresh token service call."""
                requested_entry_id = call.data.get("entry_id")
                target = _resolve_target_coordinator(hass, requested_entry_id)
                if not target:
                    _LOGGER.warning("No coordinator available for refresh_token service")
                    return

                # Clear the current token to force re-authentication
                target.client._access_token = None
                target.client._token_expires = None
                _LOGGER.info("Cleared authentication token, next API call will re-authenticate")
                # Refresh current folder data (don't change folders)
                await target.async_request_refresh()

            if not hass.services.has_service(DOMAIN, "refresh_photos"):
                hass.services.async_register(DOMAIN, "refresh_photos", handle_refresh_photos)
            if not hass.services.has_service(DOMAIN, "select_folder"):
                hass.services.async_register(DOMAIN, "select_folder", handle_select_folder)
            if not hass.services.has_service(DOMAIN, "refresh_token"):
                hass.services.async_register(DOMAIN, "refresh_token", handle_refresh_token)

            domain_data[_DOMAIN_SERVICES_REGISTERED] = True

        coordinator.start_rotation_timer()
        domain_data[setup_done_key] = True

        return True


class SharePointPhotosDataUpdateCoordinator(DataUpdateCoordinator):
    """Class to manage fetching data from the API."""

    def __init__(
        self,
        hass: HomeAssistant,
        client: SharePointPhotosApiClient,
        entry_id: str,
        rotation_interval_seconds: int,
    ) -> None:
        """Initialize."""
        self.client = client
        self._api_client = client  # Also store as _api_client for the proxy view
        self.entry_id = entry_id
        self.rotation_interval_seconds = max(5, int(rotation_interval_seconds or DEFAULT_ROTATION_INTERVAL_SECONDS))
        self._current_photo_index: int | None = None
        self._current_photo_name: str | None = None
        self._image_version: int = 0
        self._current_image_bytes: bytes | None = None
        self._current_image_type: str = "image/jpeg"
        self._current_image_loaded_monotonic: float | None = None
        self._last_on_demand_rotate_monotonic: float = 0.0
        self._on_demand_rotate_cooldown_seconds = max(5.0, min(30.0, self.rotation_interval_seconds / 2))
        self._last_empty_recovery_monotonic: float = 0.0
        self._empty_recovery_cooldown_seconds = 300.0
        self._image_lock = asyncio.Lock()
        self._rotation_unsub = None
        self._rotation_task: asyncio.Task | None = None
        super().__init__(
            hass=hass,
            logger=_LOGGER,
            name=DOMAIN,
            update_interval=None,  # Disable automatic updates - only manual refresh
        )

    def build_initial_state(self) -> dict[str, Any]:
        """Return initial empty state with stable current-image URL."""
        return {
            "photos": [],
            "photo_count": 0,
            "current_proxy_url": f"/api/sharepoint_photos/current/{self.entry_id}?v={self._image_version}",
            "rotation_interval_seconds": self.rotation_interval_seconds,
            "current_photo_index": None,
            "current_photo_name": None,
        }

    def _apply_proxy_urls(self, data: dict[str, Any]) -> None:
        """Inject entry ID into per-photo proxy URLs."""
        photos = data.get("photos", [])
        for photo in photos:
            if "proxy_url" in photo:
                photo["proxy_url"] = photo["proxy_url"].replace("{entry_id}", self.entry_id)

    def _build_state_payload(self, data: dict[str, Any]) -> dict[str, Any]:
        """Attach current-image metadata to coordinator payload."""
        payload = dict(data)
        payload["current_proxy_url"] = f"/api/sharepoint_photos/current/{self.entry_id}?v={self._image_version}"
        payload["rotation_interval_seconds"] = self.rotation_interval_seconds
        payload["photo_count"] = len(payload.get("photos", []))
        payload["current_photo_index"] = self._current_photo_index
        payload["current_photo_name"] = self._current_photo_name
        return payload

    async def _try_swap_current_photo(self, photos: list[dict[str, Any]], force: bool = False) -> bool:
        """Fetch a new random photo and atomically swap cache on success."""
        if not photos:
            return False

        indices = list(range(len(photos)))
        random.shuffle(indices)
        if not force and self._current_photo_index in indices and len(indices) > 1:
            indices = [idx for idx in indices if idx != self._current_photo_index]

        max_attempts = min(3, len(indices))
        for idx in indices[:max_attempts]:
            candidate = photos[idx]
            download_url = candidate.get("download_url")
            if not download_url:
                continue

            content, content_type, status_code = await self._api_client.fetch_image_content(download_url)
            if status_code == 200 and content:
                self._current_image_bytes = content
                if content_type:
                    self._current_image_type = content_type
                self._current_photo_index = idx
                self._current_photo_name = candidate.get("name")
                self._image_version += 1
                self._current_image_loaded_monotonic = time.monotonic()
                _LOGGER.debug(
                    "Swapped current photo (index=%s, name=%s, bytes=%d, version=%s)",
                    self._current_photo_index,
                    self._current_photo_name,
                    len(content),
                    self._image_version,
                )
                return True

            _LOGGER.debug(
                "Candidate photo fetch failed (index=%s, name=%s, status=%s)",
                idx,
                candidate.get("name"),
                status_code,
            )

        return False

    async def _async_recover_empty_photos(self) -> bool:
        """Re-select a folder after the cached photo list was emptied by an API error."""
        now = time.monotonic()
        if now - self._last_empty_recovery_monotonic < self._empty_recovery_cooldown_seconds:
            return False

        self._last_empty_recovery_monotonic = now
        _LOGGER.warning("No photos cached; attempting recovery by selecting a new folder")
        return await self.async_refresh_new_folder() is not None

    async def async_rotate_current_photo(self, force: bool = False) -> bool:
        """Rotate to a new random photo from the current folder."""
        if not (self.data or {}).get("photos"):
            return await self._async_recover_empty_photos()

        async with self._image_lock:
            data = self.data or {}
            photos = data.get("photos", [])
            swapped = await self._try_swap_current_photo(photos, force=force)
            if swapped:
                self.async_set_updated_data(self._build_state_payload(data))
                return True

        # If a rotation attempt failed but we still have photo metadata, URLs may be stale.
        # Refresh folder data to regenerate Graph download URLs, then retry once.
        if (self.data or {}).get("photos"):
            _LOGGER.warning("Rotation swap failed; refreshing folder data before one retry")
            try:
                await self.async_request_refresh()
            except Exception:
                _LOGGER.exception("Failed to refresh data during rotation recovery")

            async with self._image_lock:
                refreshed_data = self.data or {}
                refreshed_photos = refreshed_data.get("photos", [])
                swapped = await self._try_swap_current_photo(refreshed_photos, force=True)
                if swapped:
                    self.async_set_updated_data(self._build_state_payload(refreshed_data))
                    return True

        return False

    async def async_get_or_load_current_image(self) -> tuple[bytes | None, str | None]:
        """Return cached current image, loading one if cache is still empty."""
        if self._current_image_bytes:
            return self._current_image_bytes, self._current_image_type

        had_photos = bool((self.data or {}).get("photos"))

        async with self._image_lock:
            if self._current_image_bytes:
                return self._current_image_bytes, self._current_image_type

            data = self.data or {}
            photos = data.get("photos", [])
            swapped = await self._try_swap_current_photo(photos, force=True)
            if swapped:
                self.async_set_updated_data(self._build_state_payload(data))

        # If we had photos but still failed to load bytes, the stored download URLs
        # may have expired. Force a coordinator refresh to rebuild URLs and retry.
        if had_photos and not self._current_image_bytes:
            _LOGGER.warning(
                "Current image bytes unavailable despite photo metadata; refreshing coordinator data"
            )
            try:
                await self.async_request_refresh()
            except Exception:
                _LOGGER.exception("Failed to refresh data while recovering current image")

            async with self._image_lock:
                if not self._current_image_bytes:
                    refreshed_photos = (self.data or {}).get("photos", [])
                    await self._try_swap_current_photo(refreshed_photos, force=True)

        return self._current_image_bytes, self._current_image_type

    async def async_ensure_recent_image(self) -> None:
        """Ensure current image is not stale when consumers continuously request it."""
        now = time.monotonic()

        if self._current_image_loaded_monotonic is None:
            await self.async_get_or_load_current_image()
            return

        max_allowed_age = max(15.0, float(self.rotation_interval_seconds) * 1.5)
        age_seconds = now - self._current_image_loaded_monotonic
        if age_seconds <= max_allowed_age:
            return

        if now - self._last_on_demand_rotate_monotonic < self._on_demand_rotate_cooldown_seconds:
            return

        self._last_on_demand_rotate_monotonic = now
        _LOGGER.warning(
            "Current image age %.1fs exceeds threshold %.1fs; triggering on-demand rotation",
            age_seconds,
            max_allowed_age,
        )
        rotated = await self.async_rotate_current_photo(force=False)
        if not rotated:
            _LOGGER.warning("On-demand rotation did not swap to a new image")

    @callback
    def _async_rotation_tick(self, now=None) -> None:
        """Timer callback for periodic image swaps."""
        if self._rotation_task is not None and not self._rotation_task.done():
            _LOGGER.debug("Skipping rotation tick because previous rotation task is still running")
            return

        self._rotation_task = self.hass.async_create_task(self._async_run_rotation_tick())

    async def _async_run_rotation_tick(self) -> None:
        """Run one rotation tick with task-level guard."""
        try:
            await self.async_rotate_current_photo()
        except Exception:
            _LOGGER.exception("Unexpected error during scheduled rotation tick")

    def start_rotation_timer(self) -> None:
        """Start periodic rotation timer."""
        if self._rotation_unsub is not None:
            return

        self._rotation_unsub = async_track_time_interval(
            self.hass,
            self._async_rotation_tick,
            timedelta(seconds=self.rotation_interval_seconds),
        )
        _LOGGER.info("Started rotation timer (interval=%ss)", self.rotation_interval_seconds)

    def stop_rotation_timer(self) -> None:
        """Stop periodic rotation timer."""
        if self._rotation_unsub is not None:
            self._rotation_unsub()
            self._rotation_unsub = None

        if self._rotation_task is not None and not self._rotation_task.done():
            self._rotation_task.cancel()

    async def _async_update_data(self):
        """Update data via library."""
        try:
            _LOGGER.info("Starting data update for SharePoint Photos")
            data = await self.client.async_get_random_folder_photos()
            _LOGGER.info("Data update result: %s", "SUCCESS" if data else "NO DATA")
            
            if data and data.get("photos"):
                _LOGGER.info("Found %d photos in folder '%s'", len(data["photos"]), data.get("folder_name", "unknown"))
                self._apply_proxy_urls(data)
                _LOGGER.debug("Updated proxy URLs for all photos")
                await self._try_swap_current_photo(data["photos"], force=True)
            else:
                _LOGGER.warning("No photos found in data update")
                if (self.data or {}).get("photos"):
                    # Keep the previously working folder instead of publishing an empty payload.
                    raise UpdateFailed("Folder refresh returned no photos")

            if not data:
                data = self.build_initial_state()

            return self._build_state_payload(data)
        except UpdateFailed:
            raise
        except Exception as exception:
            _LOGGER.error("Error during data update: %s", str(exception))
            import traceback
            _LOGGER.error("Traceback: %s", traceback.format_exc())
            raise UpdateFailed() from exception

    async def async_refresh_new_folder(self):
        """Force refresh to a new random folder."""
        _LOGGER.info("Forcing refresh to new random folder")
        try:
            data = await self.client.async_get_random_folder_photos(force_new_folder=True)
            
            if data and data.get("photos"):
                self._apply_proxy_urls(data)
                await self._try_swap_current_photo(data["photos"], force=True)
                
                self.async_set_updated_data(self._build_state_payload(data))
                _LOGGER.info("Successfully switched to new folder: %s (%d photos)", 
                           data.get("folder_name", "unknown"), len(data["photos"]))
                return data
            else:
                _LOGGER.warning("No photos found when refreshing to new folder")
                return None
        except Exception as exception:
            _LOGGER.error("Error refreshing to new folder: %s", str(exception))
            return None


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Handle removal of an entry."""
    coordinator = hass.data.get(DOMAIN, {}).get(entry.entry_id)
    if coordinator:
        coordinator.stop_rotation_timer()

    try:
        unloaded = await hass.config_entries.async_unload_platforms(entry, PLATFORMS)
    except ValueError:
        # Platform never finished loading (e.g., setup failed); treat as unloaded.
        _LOGGER.debug("Platform %s was never loaded; skipping unload", PLATFORMS)
        unloaded = True

    if unloaded:
        domain_data = hass.data.get(DOMAIN, {})
        domain_data.pop(entry.entry_id, None)

        setup_done_key = f"{_SETUP_DONE_PREFIX}{entry.entry_id}"
        domain_data.pop(setup_done_key, None)

        listener_key = f"{_UPDATE_LISTENER_PREFIX}{entry.entry_id}"
        domain_data.pop(listener_key, None)

        setup_locks = domain_data.get(_SETUP_LOCKS_KEY, {})
        setup_locks.pop(entry.entry_id, None)

        if not _iter_entry_ids(domain_data):
            if hass.services.has_service(DOMAIN, "refresh_photos"):
                hass.services.async_remove(DOMAIN, "refresh_photos")
            if hass.services.has_service(DOMAIN, "select_folder"):
                hass.services.async_remove(DOMAIN, "select_folder")
            if hass.services.has_service(DOMAIN, "refresh_token"):
                hass.services.async_remove(DOMAIN, "refresh_token")
            domain_data[_DOMAIN_SERVICES_REGISTERED] = False

    return unloaded


async def async_reload_entry(hass: HomeAssistant, entry: ConfigEntry) -> None:
    """Persist updated options without reloading the integration immediately."""
    _LOGGER.info(
        "Config entry %s updated; saved options will be applied on the next Home Assistant restart",
        entry.entry_id,
    )
