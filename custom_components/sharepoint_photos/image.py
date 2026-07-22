"""Image platform for SharePoint Photos integration."""
from __future__ import annotations

import asyncio
import logging
import time
from datetime import datetime
from typing import Optional

from homeassistant.components.image import ImageEntity
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity
from homeassistant.util import dt as dt_util

from .const import DOMAIN

_LOGGER = logging.getLogger(__name__)


async def async_setup_entry(
    hass: HomeAssistant,
    config_entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up SharePoint Photos image entity from a config entry."""
    coordinator = hass.data[DOMAIN][config_entry.entry_id]
    async_add_entities([SharePointPhotosCurrentImage(coordinator, config_entry)])


class SharePointPhotosCurrentImage(CoordinatorEntity, ImageEntity):
    """Image entity for the current SharePoint photo."""

    _attr_name = "SharePoint Photos Current Picture"

    def __init__(self, coordinator, config_entry: ConfigEntry) -> None:
        CoordinatorEntity.__init__(self, coordinator)
        ImageEntity.__init__(self, coordinator.hass)
        if not hasattr(self, "access_tokens"):
            self.access_tokens = []
        if not self.access_tokens:
            self.async_update_token()
        self._last_content: bytes | None = None
        self._last_content_type: str | None = None
        self._last_photo_name: str | None = None
        self._fetch_lock = asyncio.Lock()
        self._config_entry = config_entry
        site_name = config_entry.data.get("site_url", "").replace("https://", "").replace("/", "_")
        self._attr_unique_id = f"{DOMAIN}_{site_name}_current_image"
        self._attr_content_type = "image/jpeg"

    def _get_current_photo(self):
        data = self.coordinator.data or {}
        photos = data.get("photos", [])
        if not photos:
            return None

        cycle_time = 10
        current_cycle = int(time.time() / cycle_time)
        photo_index = current_cycle % len(photos)
        return photos[photo_index]

    @property
    def image_last_updated(self) -> datetime | None:
        """Return the last update time for the image."""
        cycle_time = 10
        current_cycle = int(time.time() / cycle_time)
        return dt_util.utc_from_timestamp(current_cycle * cycle_time)

    async def async_image(self) -> Optional[bytes]:
        """Return bytes of image."""
        photo = self._get_current_photo()
        if not photo:
            # No photo in coordinator data yet – return stale cache if available.
            return self._last_content

        download_url = photo.get("download_url")
        if not download_url:
            return self._last_content

        photo_name = photo.get("name")

        # If this is the exact same photo we already have in cache, skip the
        # network round-trip and return immediately.  HA's image proxy calls
        # async_image() on every token request, so this avoids re-downloading
        # the same multi-MB file within the same 10-second rotation slot.
        if photo_name and photo_name == self._last_photo_name and self._last_content:
            return self._last_content

        # If a fetch is already in flight (e.g., a concurrent proxy request),
        # serve the stale cache instead of spawning a second download.
        if self._fetch_lock.locked():
            _LOGGER.debug("Fetch already in progress, returning stale cache")
            return self._last_content

        async with self._fetch_lock:
            # Re-read after acquiring the lock; state may have changed while waiting.
            photo = self._get_current_photo()
            if not photo:
                return self._last_content
            download_url = photo.get("download_url")
            if not download_url:
                return self._last_content
            photo_name = photo.get("name")

            # Second-check: another coroutine may have fetched this photo already.
            if photo_name and photo_name == self._last_photo_name and self._last_content:
                return self._last_content

            api_client = self.coordinator._api_client
            content, content_type, status_code = await api_client.fetch_image_content(download_url)

            if status_code in (401, 403):
                _LOGGER.info("Image URL expired (status=%s), refreshing coordinator data", status_code)
                await self.coordinator.async_request_refresh()
                photo = self._get_current_photo()
                if not photo:
                    return self._last_content
                download_url = photo.get("download_url")
                if not download_url:
                    return self._last_content
                photo_name = photo.get("name")
                content, content_type, status_code = await api_client.fetch_image_content(download_url)

            if status_code == 200 and content:
                if content_type:
                    self._attr_content_type = content_type
                self._last_content = content
                self._last_content_type = self._attr_content_type
                self._last_photo_name = photo_name
                return content

            # Fetch failed – serve stale cache so WallPanel never sees a blank.
            if self._last_content:
                if self._last_content_type:
                    self._attr_content_type = self._last_content_type
                _LOGGER.debug(
                    "Returning cached image after fetch failed (status=%s)",
                    status_code,
                )
                return self._last_content

            _LOGGER.debug("Failed to fetch image content (status=%s)", status_code)
            return None

    @property
    def available(self) -> bool:
        """Return if entity is available."""
        return (
            (self.coordinator.last_update_success and self.coordinator.data is not None)
            or self._last_content is not None
        )
