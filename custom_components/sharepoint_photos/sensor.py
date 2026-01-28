"""Sensor platform for SharePoint Photos integration."""

import asyncio
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional

from homeassistant.components.sensor import SensorEntity, SensorEntityDescription
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant, callback
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity
from homeassistant.helpers.event import async_track_time_interval
from homeassistant.util import dt as dt_util
from datetime import timedelta

from .const import (
    DOMAIN,
    SENSOR_CURRENT_FOLDER,
    SENSOR_FOLDER_PATH,
    SENSOR_LAST_UPDATED,
    SENSOR_PHOTO_COUNT,
    SENSOR_CURRENT_PICTURE,
)

_LOGGER = logging.getLogger(__name__)


def _select_photo_url(photo: Dict[str, Any]) -> Optional[str]:
    """Return the most reliable URL for a photo (proxy first)."""
    for key in ("proxy_url", "url", "thumbnail_url", "download_url", "web_url"):
        url = photo.get(key)
        if url:
            return url
    return None


SENSOR_DESCRIPTIONS = [
    SensorEntityDescription(
        key=SENSOR_CURRENT_FOLDER,
        name="Current Photo Folder",
        icon="mdi:folder-image",
    ),
    SensorEntityDescription(
        key=SENSOR_PHOTO_COUNT,
        name="Photo Count",
        icon="mdi:image-multiple",
        native_unit_of_measurement="photos",
    ),
    SensorEntityDescription(
        key=SENSOR_FOLDER_PATH,
        name="Folder Path",
        icon="mdi:folder-open",
    ),
    SensorEntityDescription(
        key=SENSOR_LAST_UPDATED,
        name="Last Updated",
        icon="mdi:clock-outline",
        device_class="timestamp",
    ),
    SensorEntityDescription(
        key=SENSOR_CURRENT_PICTURE,
        name="Current Picture URL",
        icon="mdi:image",
    ),
]


async def async_setup_entry(
    hass: HomeAssistant,
    config_entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up SharePoint Photos sensors from a config entry."""
    coordinator = hass.data[DOMAIN][config_entry.entry_id]
    
    entities = []
    for description in SENSOR_DESCRIPTIONS:
        if description.key == SENSOR_CURRENT_PICTURE:
            # Special sensor with 10-second updates
            entities.append(SharePointPhotosRotatingSensor(coordinator, description, config_entry))
        else:
            entities.append(SharePointPhotosSensor(coordinator, description, config_entry))
    
    async_add_entities(entities)


class SharePointPhotosSensor(CoordinatorEntity, SensorEntity):
    """Representation of a SharePoint Photos sensor."""

    def __init__(
        self,
        coordinator,
        description: SensorEntityDescription,
        config_entry: ConfigEntry,
    ) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self.entity_description = description
        self._config_entry = config_entry
        
        # Generate unique entity ID
        site_name = config_entry.data.get("site_url", "").replace("https://", "").replace("/", "_")
        self._attr_unique_id = f"{DOMAIN}_{site_name}_{description.key}"
        self._attr_name = f"SharePoint Photos {description.name}"

    @property
    def native_value(self) -> Any:
        """Return the state of the sensor."""
        if not self.coordinator.data:
            return None

        data = self.coordinator.data

        if self.entity_description.key == SENSOR_CURRENT_FOLDER:
            return data.get("folder_name")
        elif self.entity_description.key == SENSOR_PHOTO_COUNT:
            return data.get("photo_count", 0)
        elif self.entity_description.key == SENSOR_FOLDER_PATH:
            return data.get("folder_path")
        elif self.entity_description.key == SENSOR_LAST_UPDATED:
            # Convert ISO string to datetime object for timestamp sensor
            timestamp_str = data.get("last_updated")
            if timestamp_str:
                try:
                    # Parse ISO format string to datetime
                    return dt_util.parse_datetime(timestamp_str)
                except (ValueError, TypeError):
                    _LOGGER.warning("Invalid timestamp format: %s", timestamp_str)
                    return None
            return None
        elif self.entity_description.key == SENSOR_CURRENT_PICTURE:
            # Return the currently rotating picture URL
            photos = data.get("photos", [])
            if not photos:
                return None
            
            photo_urls = [_select_photo_url(photo) for photo in photos]
            photo_urls = [url for url in photo_urls if url]
            if not photo_urls:
                return None
            
            # Calculate rotating picture URL (changes every 10 seconds)
            import time
            cycle_time = 10  # seconds
            current_cycle = int(time.time() / cycle_time)
            picture_index = current_cycle % len(photo_urls)
            return photo_urls[picture_index]
        
        return None

    @property
    def extra_state_attributes(self) -> Dict[str, Any]:
        """Return additional state attributes."""
        if not self.coordinator.data:
            return {}

        data = self.coordinator.data
        
        if self.entity_description.key == SENSOR_CURRENT_FOLDER:
            # For the main folder sensor, include all photo URLs
            photos = data.get("photos", [])
            photo_urls = [_select_photo_url(photo) for photo in photos]
            photo_urls = [url for url in photo_urls if url]
            
            # Calculate rotating picture URL (changes every 10 seconds)
            current_picture_url = None
            current_photo = None
            preview_urls: List[str] = []

            if photo_urls:
                # Use current time to create a rotating index that changes every 10 seconds
                import time

                cycle_time = 10  # seconds
                current_cycle = int(time.time() / cycle_time)
                picture_index = current_cycle % len(photo_urls)
                current_picture_url = photo_urls[picture_index]
                if photos and picture_index < len(photos):
                    current_photo = photos[picture_index]

                preview_urls = photo_urls[:5]  # keep attributes compact for recorder

            attributes: Dict[str, Any] = {
                "folder_path": data.get("folder_path"),
                "photo_count": len(photos),
                "rotation_cycle_seconds": 10,
                "current_picture_url": current_picture_url,
                "current_picture_label": current_photo.get("name") if current_photo else None,
            }

            if preview_urls:
                attributes["preview_urls"] = preview_urls

            if current_photo:
                attributes["current_photo_id"] = current_photo.get("id")
                attributes["current_photo_name"] = current_photo.get("name")
                if current_photo.get("web_url"):
                    attributes["current_photo_web_url"] = current_photo["web_url"]

            recent_folders = data.get("recent_folders")
            if recent_folders:
                attributes["recent_folders"] = recent_folders

            return attributes
        
        return {}

    @property
    def entity_picture(self) -> Optional[str]:
        """Return the entity picture URL."""
        if not self.coordinator.data:
            return None
            
        data = self.coordinator.data
        
        if self.entity_description.key == SENSOR_CURRENT_FOLDER:
            # For the main folder sensor, return the rotating picture
            photos = data.get("photos", [])
            if not photos:
                return None
            
            photo_urls = [_select_photo_url(photo) for photo in photos]
            photo_urls = [url for url in photo_urls if url]
            if not photo_urls:
                return None
            
            # Calculate rotating picture URL (same logic as attributes)
            import time
            cycle_time = 10  # seconds
            current_cycle = int(time.time() / cycle_time)
            picture_index = current_cycle % len(photo_urls)
            return photo_urls[picture_index]
        
        return None

    @property
    def available(self) -> bool:
        """Return if entity is available."""
        return self.coordinator.last_update_success and self.coordinator.data is not None


class SharePointPhotosRotatingSensor(CoordinatorEntity, SensorEntity):
    """Special sensor for rotating picture URL that updates every 10 seconds."""

    def __init__(
        self,
        coordinator,
        description: SensorEntityDescription,
        config_entry: ConfigEntry,
    ) -> None:
        """Initialize the rotating sensor."""
        super().__init__(coordinator)
        self.entity_description = description
        self._config_entry = config_entry
        self._update_timer = None
        
        # Generate unique entity ID
        site_name = config_entry.data.get("site_url", "").replace("https://", "").replace("/", "_")
        self._attr_unique_id = f"{DOMAIN}_{site_name}_{description.key}"
        self._attr_name = f"SharePoint Photos {description.name}"

    async def async_added_to_hass(self) -> None:
        """Run when entity is added to hass."""
        await super().async_added_to_hass()
        
        # Set up timer to update every 10 seconds
        self._update_timer = async_track_time_interval(
            self.hass,
            self._async_update_state,
            timedelta(seconds=10)
        )

    async def async_will_remove_from_hass(self) -> None:
        """Run when entity is removed from hass."""
        if self._update_timer:
            self._update_timer()
        await super().async_will_remove_from_hass()

    @callback
    def _async_update_state(self, now=None) -> None:
        """Update the state of the sensor."""
        self.async_write_ha_state()

    @property
    def native_value(self) -> Any:
        """Return the state of the sensor."""
        if not self.coordinator.data:
            return None

        data = self.coordinator.data
        photos = data.get("photos", [])
        if not photos:
            return None
        
        photo_urls = [_select_photo_url(photo) for photo in photos]
        photo_urls = [url for url in photo_urls if url]
        if not photo_urls:
            return None
        
        # Return proxy/thumbnail URL for the current photo so dashboards can load it directly
        import time
        cycle_time = 10  # seconds
        current_cycle = int(time.time() / cycle_time)
        picture_index = current_cycle % len(photo_urls)
        
        return photo_urls[picture_index]

    @property
    def extra_state_attributes(self) -> Dict[str, Any]:
        """Return additional state attributes."""
        if not self.coordinator.data:
            return {}

        data = self.coordinator.data
        photos = data.get("photos", [])
        photo_urls = [_select_photo_url(photo) for photo in photos]
        photo_urls = [url for url in photo_urls if url]
        
        # Calculate current index for display
        import time
        cycle_time = 10  # seconds
        current_cycle = int(time.time() / cycle_time)
        current_index = current_cycle % len(photo_urls) if photo_urls else 0
        
        current_photo = photos[current_index] if photos and current_index < len(photos) else None
        
        attributes = {
            "total_photos": len(photo_urls),
            "current_index": current_index + 1,  # 1-based for display
            "current_photo_url": photo_urls[current_index] if photo_urls else None,
            "current_photo_name": current_photo.get("name") if current_photo else None,
            "current_photo_label": f"Photo {current_index + 1}: {current_photo.get('name')}" if current_photo else None,
            "cycle_time_seconds": 10,
            "folder_name": data.get("folder_name"),
        }
        
        # Add thumbnail and alternative URLs if available
        if current_photo:
            if current_photo.get("proxy_url"):
                attributes["current_photo_proxy_url"] = current_photo["proxy_url"]
            if current_photo.get("thumbnail_url"):
                attributes["current_photo_thumbnail"] = current_photo["thumbnail_url"]
            if current_photo.get("download_url"):
                attributes["current_photo_download_url"] = current_photo["download_url"]
            if current_photo.get("web_url"):
                attributes["current_photo_web_url"] = current_photo["web_url"]
        
        return attributes

    @property
    def entity_picture(self) -> Optional[str]:
        """Return the entity picture URL."""
        if not self.coordinator.data:
            return None

        data = self.coordinator.data
        photos = data.get("photos", [])
        if not photos:
            return None
        
        # Calculate rotating picture index (changes every 10 seconds)
        import time
        cycle_time = 10  # seconds
        current_cycle = int(time.time() / cycle_time)
        picture_index = current_cycle % len(photos)
        
        current_photo = photos[picture_index]
        preferred_url = _select_photo_url(current_photo)
        if preferred_url:
            _LOGGER.debug("Entity picture using URL: %s", preferred_url[:100])
        return preferred_url

    @property
    def available(self) -> bool:
        """Return if entity is available."""
        return self.coordinator.last_update_success and self.coordinator.data is not None