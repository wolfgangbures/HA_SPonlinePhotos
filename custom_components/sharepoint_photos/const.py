"""Constants for the SharePoint Photos integration."""

DOMAIN = "sharepoint_photos"
NAME = "SharePoint Photos"

# Configuration keys
CONF_TENANT_ID = "tenant_id"
CONF_CLIENT_ID = "client_id"
CONF_CLIENT_SECRET = "client_secret"
CONF_SITE_URL = "site_url"
CONF_LIBRARY_NAME = "library_name"
CONF_BASE_FOLDER_PATH = "base_folder_path"
CONF_REFRESH_INTERVAL = "refresh_interval"
CONF_FOLDER_HISTORY_SIZE = "folder_history_size"
CONF_MIN_PHOTO_COUNT = "min_photo_count"

# Default values
DEFAULT_LIBRARY_NAME = "Freigegebene Dokumente"  # German SharePoint default
DEFAULT_BASE_FOLDER_PATH = "/General/Fotos"
DEFAULT_REFRESH_INTERVAL = 6  # hours
DEFAULT_FOLDER_HISTORY_SIZE = 30
DEFAULT_MIN_PHOTO_COUNT = 5

# Microsoft Graph API
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
AUTHORITY_BASE = "https://login.microsoftonline.com"
SCOPE = ["https://graph.microsoft.com/.default"]

# Sensor types
SENSOR_CURRENT_FOLDER = "current_folder"
SENSOR_PHOTO_COUNT = "photo_count"
SENSOR_LAST_UPDATED = "last_updated"
SENSOR_FOLDER_PATH = "folder_path"
SENSOR_CURRENT_PICTURE = "current_picture"

# Services
SERVICE_REFRESH_PHOTOS = "refresh_photos"
SERVICE_SELECT_FOLDER = "select_folder"

# Data update coordinator
UPDATE_INTERVAL_SECONDS = 3600  # 1 hour for checking folder structure changes

# Supported image extensions
IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff"]

# Error messages
ERROR_AUTH_FAILED = "authentication_failed"
ERROR_SITE_NOT_FOUND = "site_not_found"
ERROR_LIBRARY_NOT_FOUND = "library_not_found"
ERROR_FOLDER_NOT_FOUND = "folder_not_found"
ERROR_NO_PHOTOS = "no_photos_found"
ERROR_NETWORK = "network_error"