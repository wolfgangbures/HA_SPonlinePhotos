# Copilot Instructions

## Overview
- This repository contains a Home Assistant custom integration that surfaces SharePoint folders/photos as sensors via Microsoft Graph (see [custom_components/sharepoint_photos](custom_components/sharepoint_photos)).
- The flow is: config flow collects Azure AD credentials → `SharePointPhotosApiClient` handles auth + folder scanning → `SharePointPhotosDataUpdateCoordinator` caches a single folder payload → sensor entities render metadata/rotation while `/api/sharepoint_photos/image/{entry_id}/{index}` proxies image bytes.
- The integration intentionally shows one folder at a time; coordinator updates happen only on manual refresh/service calls, so avoid adding periodic polling unless explicitly requested.

## Key Components
- [custom_components/sharepoint_photos/api.py](custom_components/sharepoint_photos/api.py) centralizes Microsoft Graph calls, folder caching, random selection, recent-folder exclusion, and proxy-friendly photo payloads; reuse its helpers (`get_photo_folders`, `async_get_random_folder_photos`, `select_specific_folder`) instead of duplicating Graph requests.
- [custom_components/sharepoint_photos/__init__.py](custom_components/sharepoint_photos/__init__.py) wires the DataUpdateCoordinator, registers the HTTP proxy view, and exposes HA services (`refresh_photos`, `select_folder`, `refresh_token`). Keep new behavior inside the coordinator/client so sensors stay thin.
- [custom_components/sharepoint_photos/sensor.py](custom_components/sharepoint_photos/sensor.py) defines five sensors; rotation is calculated client-side every 10s using timestamps, so any new attributes must remain lightweight to avoid HA recorder bloat.
- Config UI definitions live in [custom_components/sharepoint_photos/config_flow.py](custom_components/sharepoint_photos/config_flow.py) and [custom_components/sharepoint_photos/strings.json](custom_components/sharepoint_photos/strings.json); extend both when introducing new options.
- Dependencies (`aiohttp`, `msal`) are declared in [custom_components/sharepoint_photos/manifest.json](custom_components/sharepoint_photos/manifest.json); add Graph helpers here if you introduce new libraries.

## Developer Workflows
- Local testing happens inside a Home Assistant instance: copy the `custom_components/sharepoint_photos` folder into your HA config, restart, and use Settings → Devices & Services to add the integration (details in [README.md](README.md)).
- Use Developer Tools → Services to call `sharepoint_photos.refresh_photos` for random folder changes or `sharepoint_photos.select_folder` with a `folder_path` payload to target specific folders; sensors update immediately if the coordinator succeeds.
- The HTTP proxy relies on coordinator data; when troubleshooting broken images, hit `/api/sharepoint_photos/image/{entry_id}/{index}` in a browser while watching Home Assistant logs for `SharePointImageProxyView` output.
- Logging is already verbose (`_LOGGER.info/debug` across the client); prefer reusing existing log contexts rather than adding new standalone prints to keep HA logs cohesive.

## Conventions & Patterns
- Auth: `SharePointPhotosApiClient.authenticate()` first attempts a direct OAuth POST, then falls back to MSAL; keep new Graph calls async and reuse `_make_authenticated_request` so token refresh/401 retries stay centralized.
- Folder selection: the client tracks `_recent_folder_paths` and `_current_folder_path`; when adding features (e.g., filters, weighting) ensure `_record_folder_history` still reflects actual folder choices so UI history stays accurate.
- Payload shape: `SharePointPhotosApiClient._build_folder_payload()` defines the JSON the coordinator and sensors expect (`folder_name`, `folder_path`, `photos`, `recent_folders`). Extending sensors means updating this payload plus sensor attributes together.
- Image data: each photo dict carries `proxy_url` placeholders replaced in `SharePointPhotosDataUpdateCoordinator._async_update_data`; if you add new binary endpoints, remember to inject `entry_id` before exposing them.
- Services: All HA services are registered in `async_setup_entry`; new services should delegate to coordinator/client methods and include matching descriptions in [custom_components/sharepoint_photos/services.yaml](custom_components/sharepoint_photos/services.yaml).
- Translations: Keep English strings in `strings.json`/`translations/en.json`; add matching keys when you surface new form fields or errors so HA’s UI stays localized.

## Extending Safely
- When adding new config fields, update `const.py` (keys + defaults), `config_flow.py`, translations, and ensure options flow picks up overrides; otherwise options won’t persist.
- For new sensors or platforms, extend `PLATFORMS` in `__init__.py`, provide entity descriptions, and ensure coordinator payloads include the required data before registering the entity.
- Respect the single-folder design: any automation that changes folders should go through the provided services so folder history, proxy URLs, and coordinator cache stay in sync.
- Graph queries should stay paginated-light; reuse `_scan_folders_recursive` and `get_folder_photos` patterns (expand thumbnails, filter by `IMAGE_EXTENSIONS`) to minimize API calls and avoid blocking HA’s event loop.


## GitHub
- Always split branches by feature/fix for PRs; avoid working directly on `main`.
- Write clear, descriptive commit messages; reference related issues/PRs when applicable.
- Merge PRs after the changes immediately.
- Tag releases in GitHub matching the version in `manifest.json`.
- increment the version in `manifest.json` for every PR that changes functionality.