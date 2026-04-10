# Desktop Bridge

## Files

- `src-tauri/src/main.rs`
- `src-tauri/tauri.conf.json`
- `app/scripts/shared/app-config.js`

## Browser Bridge

`app-config.js` exposes a `window.PlutusDesktop` bridge with methods for:

- reading and writing array stores
- reading and writing editable JSON files
- listing ShareDrive folders and children
- downloading and uploading ShareDrive files
- running Microsoft Graph device-code auth
- reading the current Graph session

In Tauri, those calls are backed by `__TAURI__.tauri.invoke(...)` when the desktop runtime is available.

## Backend responsibilities

`src-tauri/src/main.rs` is responsible for:

- locating bundled app data from `build/web/` or fallback source paths
- loading ShareDrive configuration
- maintaining runtime storage under the app data directory
- exposing Tauri commands for browser requests
- handling Microsoft Graph access token lifecycle

## Storage model

The desktop runtime uses two storage concepts:

### Runtime store

Stored under the desktop app data directory by default.

Used for:

- array stores
- Graph session persistence
- editable JSON copies

### Bundled data fallback

If editable JSON is missing, the app falls back to bundled files from:

- `build/web/data/`
- `app/data/`

## Graph authentication

The Tauri desktop shell supports:

- device-code auth
- persisted refresh/access token reuse
- client credentials token fallback if environment variables are set

This lets browser code stay focused on product behavior while the desktop runtime manages the Microsoft-specific auth flow.
