"""Config flow for SharePoint Photos integration."""

import logging
from typing import Any, Dict, Optional

import voluptuous as vol
from homeassistant import config_entries
from homeassistant.const import CONF_NAME
from homeassistant.core import HomeAssistant, callback
from homeassistant.data_entry_flow import FlowResult

from .api import SharePointPhotosApiClient
from .const import (
    CONF_BASE_FOLDER_PATH,
    CONF_CLIENT_ID,
    CONF_CLIENT_SECRET,
    CONF_FOLDER_HISTORY_SIZE,
    CONF_LIBRARY_NAME,
    CONF_SITE_URL,
    CONF_TENANT_ID,
    DEFAULT_BASE_FOLDER_PATH,
    DEFAULT_FOLDER_HISTORY_SIZE,
    DEFAULT_LIBRARY_NAME,
    DOMAIN,
    ERROR_AUTH_FAILED,
    ERROR_NETWORK,
    ERROR_SITE_NOT_FOUND,
)

_LOGGER = logging.getLogger(__name__)


class SharePointPhotosConfigFlow(config_entries.ConfigFlow, domain=DOMAIN):
    """Handle a config flow for SharePoint Photos."""

    VERSION = 1

    def __init__(self) -> None:
        """Initialize the config flow."""
        self._user_input: Dict[str, Any] = {}

    async def async_step_user(
        self, user_input: Optional[Dict[str, Any]] = None
    ) -> FlowResult:
        """Handle the initial step."""
        errors = {}

        if user_input is not None:
            self._user_input.update(user_input)
            
            # Test the connection
            client = SharePointPhotosApiClient(
                hass=self.hass,
                tenant_id=user_input[CONF_TENANT_ID],
                client_id=user_input[CONF_CLIENT_ID],
                client_secret=user_input[CONF_CLIENT_SECRET],
                site_url=user_input[CONF_SITE_URL],
                library_name=user_input.get(CONF_LIBRARY_NAME, DEFAULT_LIBRARY_NAME),
                base_folder_path=user_input.get(CONF_BASE_FOLDER_PATH, DEFAULT_BASE_FOLDER_PATH),
                recent_history_size=user_input.get(CONF_FOLDER_HISTORY_SIZE, DEFAULT_FOLDER_HISTORY_SIZE),
            )

            try:
                if await client.test_connection():
                    # Create a unique ID for this config entry
                    site_id = user_input[CONF_SITE_URL].replace("https://", "").replace("/", "_")
                    await self.async_set_unique_id(f"{DOMAIN}_{site_id}")
                    self._abort_if_unique_id_configured()

                    return self.async_create_entry(
                        title=f"SharePoint Photos ({user_input[CONF_SITE_URL]})",
                        data=self._user_input,
                    )
                else:
                    errors["base"] = ERROR_AUTH_FAILED

            except Exception as e:
                _LOGGER.error("Unexpected error during setup: %s", str(e))
                if "site" in str(e).lower():
                    errors["base"] = ERROR_SITE_NOT_FOUND
                else:
                    errors["base"] = ERROR_NETWORK

        return self.async_show_form(
            step_id="user",
            data_schema=vol.Schema({
                vol.Required(CONF_TENANT_ID): str,
                vol.Required(CONF_CLIENT_ID): str,
                vol.Required(CONF_CLIENT_SECRET): str,
                vol.Required(CONF_SITE_URL): str,
                vol.Optional(CONF_LIBRARY_NAME, default=DEFAULT_LIBRARY_NAME): str,
                vol.Optional(CONF_BASE_FOLDER_PATH, default=DEFAULT_BASE_FOLDER_PATH): str,
                vol.Optional(
                    CONF_FOLDER_HISTORY_SIZE,
                    default=DEFAULT_FOLDER_HISTORY_SIZE
                ): vol.All(vol.Coerce(int), vol.Range(min=0, max=200)),
            }),
            errors=errors,
        )

    @staticmethod
    @callback
    def async_get_options_flow(
        config_entry: config_entries.ConfigEntry,
    ) -> "SharePointPhotosOptionsFlow":
        """Get the options flow for this handler."""
        return SharePointPhotosOptionsFlow(config_entry)


class SharePointPhotosOptionsFlow(config_entries.OptionsFlow):
    """Handle SharePoint Photos options."""

    def __init__(self, config_entry: config_entries.ConfigEntry) -> None:
        """Initialize SharePoint Photos options flow."""
        self.config_entry = config_entry

    async def async_step_init(
        self, user_input: Optional[Dict[str, Any]] = None
    ) -> FlowResult:
        """Manage the options."""
        if user_input is not None:
            return self.async_create_entry(title="", data=user_input)

        return self.async_show_form(
            step_id="init",
            data_schema=vol.Schema({
                vol.Optional(
                    CONF_LIBRARY_NAME,
                    default=self.config_entry.options.get(
                        CONF_LIBRARY_NAME, 
                        self.config_entry.data.get(CONF_LIBRARY_NAME, DEFAULT_LIBRARY_NAME)
                    ),
                ): str,
                vol.Optional(
                    CONF_BASE_FOLDER_PATH,
                    default=self.config_entry.options.get(
                        CONF_BASE_FOLDER_PATH,
                        self.config_entry.data.get(CONF_BASE_FOLDER_PATH, DEFAULT_BASE_FOLDER_PATH)
                    ),
                ): str,
                vol.Optional(
                    CONF_FOLDER_HISTORY_SIZE,
                    default=self.config_entry.options.get(
                        CONF_FOLDER_HISTORY_SIZE,
                        self.config_entry.data.get(CONF_FOLDER_HISTORY_SIZE, DEFAULT_FOLDER_HISTORY_SIZE)
                    ),
                ): vol.All(vol.Coerce(int), vol.Range(min=0, max=200)),
            }),
        )