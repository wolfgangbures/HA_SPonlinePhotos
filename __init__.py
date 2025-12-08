"""The SharePoint Photos integration."""
from __future__ import annotations

import asyncio
import logging
from datetime import timedelta
from typing import Any

from homeassistant.config_entries import ConfigEntry
from homeassistant.const import Platform
from homeassistant.core import HomeAssistant
from homeassistant.exceptions import ConfigEntryNotReady
from homeassistant.helpers.update_coordinator import DataUpdateCoordinator, UpdateFailed
from homeassistant.components.http import HomeAssistantView

from .api import SharePointPhotosApiClient
from .const import DOMAIN

PLATFORMS: list[Platform] = [Platform.SENSOR]

_LOGGER = logging.getLogger(__name__)


class SharePointImageProxyView(HomeAssistantView):
    """Proxy view for SharePoint images to handle authentication."""
    
    url = "/api/sharepoint_photos/image/{entry_id}/{image_id}"
    name = "api:sharepoint_photos:image"
    requires_auth = False  # We'll handle auth internally

    def __init__(self, hass: HomeAssistant):
        """Initialize the proxy view."""
        self.hass = hass

    async def get(self, request, entry_id: str, image_id: str):
        """Proxy SharePoint image requests."""
        from aiohttp import web
        import aiohttp
        
        _LOGGER.debug("Proxy request received: entry_id=%s, image_id=%s", entry_id, image_id)
        
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
                if photo_index < 0 or photo_index >= len(photos):
                    _LOGGER.error("Photo index %d out of range (0-%d)", photo_index, len(photos)-1)
                    return web.Response(status=404, text="Photo not found")
                
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
            
            if status_code == 401:
                # Token expired, try to refresh the data and get new URLs
                _LOGGER.info("Image URL expired, refreshing photo data...")
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
                _LOGGER.debug("Successfully proxied image: %d bytes, type: %s", len(content), content_type)
                
                return web.Response(
                    body=content,
                    content_type=content_type,
                    headers={
                        'Cache-Control': 'public, max-age=3600',  # Cache for 1 hour
                        'Content-Length': str(len(content)),
                        'Access-Control-Allow-Origin': '*',  # Allow CORS
                    }
                )
            else:
                _LOGGER.error("Failed to fetch image from SharePoint: HTTP %d", status_code)
                return web.Response(status=status_code if status_code else 500, text="Failed to fetch image")
                        
        except Exception as e:
            _LOGGER.error("Error proxying SharePoint image: %s", str(e))
            return web.Response(status=500, text="Internal server error")


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Set up this integration using UI."""
    if hass.data.get(DOMAIN) is None:
        hass.data.setdefault(DOMAIN, {})

    tenant_id = entry.data.get("tenant_id")
    client_id = entry.data.get("client_id")
    client_secret = entry.data.get("client_secret")
    site_url = entry.data.get("site_url")
    library_name = entry.data.get("library_name", "Documents")
    base_folder_path = entry.data.get("base_folder_path", "/Photos")

    client = SharePointPhotosApiClient(
        hass=hass,
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        site_url=site_url,
        library_name=library_name,
        base_folder_path=base_folder_path,
    )

    coordinator = SharePointPhotosDataUpdateCoordinator(hass, client=client, entry_id=entry.entry_id)
    await coordinator.async_config_entry_first_refresh()

    if not coordinator.last_update_success:
        raise ConfigEntryNotReady

    hass.data[DOMAIN][entry.entry_id] = coordinator

    # Register the image proxy view (only if not already registered)
    if not hasattr(hass.http, '_sharepoint_photos_proxy_registered'):
        hass.http.register_view(SharePointImageProxyView(hass))
        hass.http._sharepoint_photos_proxy_registered = True
        _LOGGER.debug("Registered SharePoint Photos image proxy view")

    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)
    entry.async_on_unload(entry.add_update_listener(async_reload_entry))

    # Register services
    async def handle_refresh_photos(call):
        """Handle the refresh photos service call - switches to a NEW random folder."""
        coordinator = hass.data[DOMAIN][entry.entry_id]
        await coordinator.async_refresh_new_folder()

    async def handle_select_folder(call):
        """Handle the select folder service call."""
        folder_path = call.data.get("folder_path")
        if folder_path:
            coordinator = hass.data[DOMAIN][entry.entry_id]
            await coordinator.client.select_specific_folder(folder_path)
            await coordinator.async_request_refresh()
    
    async def handle_refresh_token(call):
        """Handle the refresh token service call."""
        coordinator = hass.data[DOMAIN][entry.entry_id]
        # Clear the current token to force re-authentication
        coordinator.client._access_token = None
        coordinator.client._token_expires = None
        _LOGGER.info("Cleared authentication token, next API call will re-authenticate")
        # Refresh current folder data (don't change folders)
        await coordinator.async_request_refresh()

    hass.services.async_register(DOMAIN, "refresh_photos", handle_refresh_photos)
    hass.services.async_register(DOMAIN, "select_folder", handle_select_folder)
    hass.services.async_register(DOMAIN, "refresh_token", handle_refresh_token)

    return True


class SharePointPhotosDataUpdateCoordinator(DataUpdateCoordinator):
    """Class to manage fetching data from the API."""

    def __init__(
        self,
        hass: HomeAssistant,
        client: SharePointPhotosApiClient,
        entry_id: str,
    ) -> None:
        """Initialize."""
        self.client = client
        self._api_client = client  # Also store as _api_client for the proxy view
        self.entry_id = entry_id
        super().__init__(
            hass=hass,
            logger=_LOGGER,
            name=DOMAIN,
            update_interval=None,  # Disable automatic updates - only manual refresh
        )

    async def _async_update_data(self):
        """Update data via library."""
        try:
            _LOGGER.info("Starting data update for SharePoint Photos")
            data = await self.client.async_get_random_folder_photos()
            _LOGGER.info("Data update result: %s", "SUCCESS" if data else "NO DATA")
            
            if data and data.get("photos"):
                _LOGGER.info("Found %d photos in folder '%s'", len(data["photos"]), data.get("folder_name", "unknown"))
                # Update proxy URLs with the actual entry_id
                for photo in data["photos"]:
                    if "proxy_url" in photo:
                        photo["proxy_url"] = photo["proxy_url"].replace("{entry_id}", self.entry_id)
                _LOGGER.debug("Updated proxy URLs for all photos")
            else:
                _LOGGER.warning("No photos found in data update")
                
            return data
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
                # Update proxy URLs with the actual entry_id
                for photo in data["photos"]:
                    if "proxy_url" in photo:
                        photo["proxy_url"] = photo["proxy_url"].replace("{entry_id}", self.entry_id)
                
                # Update the coordinator's data directly
                self.async_set_updated_data(data)
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
    if unloaded := await hass.config_entries.async_unload_platforms(entry, PLATFORMS):
        hass.data[DOMAIN].pop(entry.entry_id)

    return unloaded


async def async_reload_entry(hass: HomeAssistant, entry: ConfigEntry) -> None:
    """Reload config entry."""
    await async_unload_entry(hass, entry)
    await async_setup_entry(hass, entry)
