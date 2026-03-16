# System Overview

## Layers

The project has five important layers:

1. Browser application
   - Static HTML, CSS, and JavaScript in `app/`
2. Generated web bundle
   - Copied output in `build/web/`
3. Electron desktop shell
   - Local persistence and Microsoft Graph bridge in `electron/`
4. Capacitor native wrappers
   - iOS and Android containers in `ios/` and `android/`
5. Tooling and configuration
   - Build utilities in `tools/` and shared paths in `config/`

## High-level flow

```text
app/ source
  -> tools/prepare-web-assets.cjs
  -> build/web/
  -> Electron loads build/web/
  -> Capacitor sync copies build/web/ into native projects
```

## Runtime responsibilities

### `app/scripts/shared/app-config.js`

Defines shared application metadata:

- page IDs and HTML filenames
- navigation labels
- storage key names
- data file names

### `app/scripts/shared/layout.js`

Builds repeated layout elements from shared config:

- left navigation sidebar
- route-based links declared with `data-route-id`

### `app/scripts/shared/data-loader.js`

Bootstraps the browser runtime by loading:

- base configuration
- desktop overrides when available
- seed data for deals and tasks when sync is not active

### `app/scripts/shared/core.js`

Owns the application runtime layer:

- local storage and desktop storage helpers
- dashboard config merging
- deals/tasks load and save
- Sharedrive sync behavior
- browser auth helpers

## Desktop-specific flow

Electron uses:

- `electron/preload.cjs` to expose a safe browser bridge
- `electron/main.cjs` to implement storage, file access, Graph auth, and ShareDrive operations

The browser app checks `window.PlutusDesktop` to know when Electron capabilities are available.
