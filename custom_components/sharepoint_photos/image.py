"""Image platform for SharePoint Photos integration."""
from __future__ import annotations

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
        self._config_entry = config_entry
        site_name = config_entry.data.get("site_url", "").replace("https://", "").replace("/", "_")
        self._attr_unique_id = f"{DOMAIN}_{site_name}_current_image"
        self._attr_content_type = "image/jpeg"

    @property
    def image_last_updated(self) -> datetime | None:
        """Return the last update time for the image."""
        interval = (self.coordinator.data or {}).get(
            "rotation_interval_seconds",
            getattr(self.coordinator, "rotation_interval_seconds", 10),
        )
        cycle_time = max(1, int(interval))
        current_cycle = int(time.time() / cycle_time)
        return dt_util.utc_from_timestamp(current_cycle * cycle_time)

    async def async_image(self) -> Optional[bytes]:
        """Return bytes of image."""
        try:
            content, content_type = await self.coordinator.async_get_or_load_current_image()
            if content_type:
                self._attr_content_type = content_type
            return content
        except Exception as e:
            _LOGGER.error("Unexpected error in async_image: %s", e)
            return None

    @property
    def available(self) -> bool:
        """Return if entity is available."""
        data = self.coordinator.data or {}
        return self.coordinator.last_update_success or bool(data.get("current_photo_name"))
