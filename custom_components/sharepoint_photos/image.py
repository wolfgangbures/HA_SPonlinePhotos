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
        try:
            data = self.coordinator.data or {}
            photos = data.get("photos", [])
            if not photos:
                return None

            cycle_time = 10
            current_cycle = int(time.time() / cycle_time)
            photo_index = current_cycle % len(photos)
            return photos[photo_index]
        except Exception as e:
            _LOGGER.warning("Error in _get_current_photo: %s", e)
            return None

    def _get_current_photo_index(self) -> int | None:
        """Return rotating photo index used by the image entity."""
        try:
            data = self.coordinator.data or {}
            photos = data.get("photos", [])
            if not photos:
                return None

            cycle_time = 10
            current_cycle = int(time.time() / cycle_time)
            return current_cycle % len(photos)
        except Exception as e:
            _LOGGER.debug("Failed to compute current photo index: %s", e)
            return None

    @property
    def image_last_updated(self) -> datetime | None:
        """Return the last update time for the image."""
        cycle_time = 10
        current_cycle = int(time.time() / cycle_time)
        return dt_util.utc_from_timestamp(current_cycle * cycle_time)

    async def async_image(self) -> Optional[bytes]:
        """Return bytes of image."""
        try:
            return await self._async_image_impl()
        except Exception as e:
            _LOGGER.error("Unexpected error in async_image: %s", e)
            # Always return stale cache on any exception to prevent WallPanel from seeing failures
            return self._last_content

    async def _async_image_impl(self) -> Optional[bytes]:
        """Internal implementation of image fetch."""
        started_at = time.monotonic()
        photo = self._get_current_photo()
        photo_index = self._get_current_photo_index()
        if not photo:
            # No photo in coordinator data yet – return stale cache if available.
            _LOGGER.debug(
                "No current photo in coordinator data (cache_available=%s)",
                self._last_content is not None,
            )
            return self._last_content

        download_url = photo.get("download_url")
        if not download_url:
            _LOGGER.debug(
                "Current photo has no download URL (index=%s, name=%s, cache_available=%s)",
                photo_index,
                photo.get("name"),
                self._last_content is not None,
            )
            return self._last_content

        photo_name = photo.get("name")

        # If this is the exact same photo we already have in cache, skip the
        # network round-trip and return immediately.  HA's image proxy calls
        # async_image() on every token request, so this avoids re-downloading
        # the same multi-MB file within the same 10-second rotation slot.
        if photo_name and photo_name == self._last_photo_name and self._last_content:
            _LOGGER.debug(
                "Serving cached image for same photo (index=%s, name=%s, size=%d bytes)",
                photo_index,
                photo_name,
                len(self._last_content),
            )
            return self._last_content

        # If a fetch is already in flight (e.g., a concurrent proxy request),
        # wait briefly for it when no stale cache exists to avoid returning None.
        if self._fetch_lock.locked():
            if self._last_content:
                _LOGGER.debug(
                    "Fetch already in progress; returning stale cache (index=%s, name=%s, size=%d bytes)",
                    photo_index,
                    photo_name,
                    len(self._last_content),
                )
                return self._last_content

            _LOGGER.debug(
                "Fetch already in progress with no stale cache; waiting for in-flight fetch (index=%s, name=%s)",
                photo_index,
                photo_name,
            )
            async with self._fetch_lock:
                if self._last_content:
                    _LOGGER.debug(
                        "In-flight fetch produced cache; serving it (index=%s, name=%s, size=%d bytes)",
                        photo_index,
                        photo_name,
                        len(self._last_content),
                    )
                    return self._last_content

            _LOGGER.debug(
                "In-flight fetch finished without cached content; proceeding with direct fetch (index=%s, name=%s)",
                photo_index,
                photo_name,
            )

        async with self._fetch_lock:
            # Re-read after acquiring the lock; state may have changed while waiting.
            photo = self._get_current_photo()
            photo_index = self._get_current_photo_index()
            if not photo:
                _LOGGER.debug(
                    "Photo disappeared while waiting for fetch lock (cache_available=%s)",
                    self._last_content is not None,
                )
                return self._last_content
            download_url = photo.get("download_url")
            if not download_url:
                _LOGGER.debug(
                    "Photo missing download URL after lock acquire (index=%s, name=%s)",
                    photo_index,
                    photo.get("name"),
                )
                return self._last_content
            photo_name = photo.get("name")

            # Second-check: another coroutine may have fetched this photo already.
            if photo_name and photo_name == self._last_photo_name and self._last_content:
                _LOGGER.debug(
                    "Photo already fetched by another coroutine (index=%s, name=%s, size=%d bytes)",
                    photo_index,
                    photo_name,
                    len(self._last_content),
                )
                return self._last_content

            api_client = self.coordinator._api_client
            _LOGGER.debug(
                "Fetching current image from SharePoint (index=%s, name=%s)",
                photo_index,
                photo_name,
            )
            content, content_type, status_code = await api_client.fetch_image_content(download_url)

            if status_code in (401, 403):
                _LOGGER.info("Image URL expired (status=%s), refreshing coordinator data", status_code)
                await self.coordinator.async_request_refresh()
                photo = self._get_current_photo()
                photo_index = self._get_current_photo_index()
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
                elapsed_ms = int((time.monotonic() - started_at) * 1000)
                _LOGGER.debug(
                    "Image fetch success (index=%s, name=%s, status=%s, bytes=%d, type=%s, duration_ms=%d)",
                    photo_index,
                    photo_name,
                    status_code,
                    len(content),
                    self._attr_content_type,
                    elapsed_ms,
                )
                return content

            # Fetch failed – serve stale cache so WallPanel never sees a blank.
            if self._last_content:
                if self._last_content_type:
                    self._attr_content_type = self._last_content_type
                elapsed_ms = int((time.monotonic() - started_at) * 1000)
                _LOGGER.debug(
                    "Returning cached image after fetch failed (index=%s, name=%s, status=%s, cache_bytes=%d, duration_ms=%d)",
                    photo_index,
                    photo_name,
                    status_code,
                    len(self._last_content),
                    elapsed_ms,
                )
                return self._last_content

            elapsed_ms = int((time.monotonic() - started_at) * 1000)
            _LOGGER.warning(
                "No image available after fetch (index=%s, name=%s, status=%s, duration_ms=%d)",
                photo_index,
                photo_name,
                status_code,
                elapsed_ms,
            )
            return None

    @property
    def available(self) -> bool:
        """Return if entity is available."""
        return (
            (self.coordinator.last_update_success and self.coordinator.data is not None)
            or self._last_content is not None
        )
