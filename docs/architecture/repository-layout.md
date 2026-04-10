# Repository Layout

## Top-level directories

```text
app/
build/
config/
src-tauri/
tools/
```

## `app/`

Canonical product source.

```text
app/
├── data/
├── index.html
├── pages/
├── scripts/
└── styles/
```

### `app/data/`

Tracked JSON configuration and seed-like inputs used by the browser app and desktop runtime.

Important files:

- `config.json`
- `sharedrive-tasks.json`

### `app/pages/`

HTML entrypoints for user-facing screens such as:

- `investor-dashboard.html`
- `deals-overview.html`
- `deal-details.html`
- `tasks-management.html`
- `owner-tasks.html`
- `sharedrive-folders.html`
- `deal-ownership.html`
- `accounting.html`

### `app/scripts/`

Browser JavaScript split into:

- `shared/`
- `pages/`

### `app/styles/`

Shared and page-specific CSS.

## `build/web/`

Generated bundle copied from `app/`.

This folder exists so the desktop shell can consume a clean output tree without treating source folders as platform assets.

## `config/`

Shared metadata for tooling and packaging.

The key file is `config/project-paths.cjs`, which defines:

- root paths
- build output paths
- standard web subdirectory names
- standard JSON filenames

## `src-tauri/`

Desktop shell and desktop-only services.

- `src/main.rs`: Tauri commands, Graph API integration, and local persistence
- `tauri.conf.json`: desktop window and bundle configuration
- `build.rs`: Tauri build script

## `tools/`

Node-side repository tools.

- `prepare-web-assets.cjs`
- `generate-config.js`
- `generate-brand-icons.py`
- `debug-sharedrive-json.cjs`
