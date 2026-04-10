# Browser Runtime

## Page boot order

Most pages follow the same initialization sequence:

1. `app-config.js`
2. `sharedrive-gate.js` when the page requires a ShareDrive session
3. `layout.js`
4. `data-loader.js`
5. `core.js`
6. page-specific script

## Shared runtime globals

`data-loader.js` initializes these browser globals:

- `window.DASHBOARD_CONFIG`
- `window.DASHBOARD_PROXIES`
- `window.DEALS`
- `window.TASKS`
- `window.SHAREDRIVE_TASKS`

`core.js` then exposes a higher-level API as:

- `window.AppCore`

## Why both `data-loader.js` and `core.js` exist

### `data-loader.js`

Does lightweight bootstrap work:

- loads JSON files
- applies desktop overrides when available
- exposes initial in-memory values

### `core.js`

Adds application logic:

- normalized route helpers
- local and desktop persistence
- sync orchestration
- dashboard config merge logic
- Graph session helpers

## Route and navigation handling

`app-config.js` defines the route map once.

Pages and page scripts should prefer:

- `window.PlutusAppConfig.buildPageHref(...)`
- `window.AppCore.getPageUrl(...)`

This avoids scattering hardcoded page filenames throughout the app.

## Data sources at runtime

The app can read from several sources depending on environment:

1. Bundled JSON in `app/data/` or `build/web/data/`
2. Desktop-editable JSON files exposed by the Tauri bridge
3. Local browser storage
4. Sharedrive-hosted JSON files when sync is enabled

## When Desktop Bridge Changes Behavior

If `window.PlutusDesktop` is available:

- page code can read or write editable JSON through the desktop bridge
- page code can query Graph session state
- page code can list, upload, and download ShareDrive files

If it is not available, pages fall back to browser-only behavior where possible.
