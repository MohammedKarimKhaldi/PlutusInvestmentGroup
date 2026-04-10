# System Overview

## Layers

The project has four important layers:

1. Browser application
   - Static HTML, CSS, and JavaScript in `app/`
2. Generated web bundle
   - Copied output in `build/web/`
3. Tauri desktop shell
   - Local persistence and Microsoft Graph bridge in `src-tauri/`
4. Tooling and configuration
   - Build utilities in `tools/` and shared paths in `config/`

## High-level flow

```text
app/ source
  -> tools/prepare-web-assets.cjs
  -> build/web/
  -> Tauri loads build/web/
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

### `app/scripts/shared/core.js`

Owns the application runtime layer:

- local storage and desktop storage helpers
- dashboard config merging
- deals/tasks load and save
- Sharedrive sync behavior
- browser auth helpers

## Desktop-specific flow

Tauri uses:

- `app/scripts/shared/app-config.js` to expose a `window.PlutusDesktop` compatibility bridge in the browser layer
- `src-tauri/src/main.rs` to implement storage, file access, Graph auth, and ShareDrive operations

The browser app checks `window.PlutusDesktop` to know when desktop capabilities are available.
