# Electron Bridge

## Files

- `electron/main.cjs`
- `electron/preload.cjs`

## Preload API

`electron/preload.cjs` exposes a `window.PlutusDesktop` bridge with methods for:

- reading and writing array stores
- reading and writing editable JSON files
- listing ShareDrive folders and children
- downloading and uploading ShareDrive files
- running Microsoft Graph device-code auth
- reading the current Graph session

This keeps renderer code isolated from direct Node access.

## Main process responsibilities

`electron/main.cjs` is responsible for:

- locating bundled app data from `build/web/` or fallback source paths
- loading ShareDrive configuration
- maintaining runtime storage under the app data directory
- exposing IPC handlers for renderer requests
- creating the Electron browser window
- handling Microsoft Graph access token lifecycle

## Storage model

Electron uses two storage concepts:

### Runtime store

Stored under the Electron app data directory by default.

Used for:

- array stores
- Graph session persistence
- editable JSON copies

### Bundled data fallback

If editable JSON is missing, the app falls back to bundled files from:

- `build/web/data/`
- `app/data/`

## Graph authentication

The desktop shell supports:

- device-code auth
- persisted refresh/access token reuse
- client credentials token fallback if environment variables are set

This lets browser code stay focused on product behavior while Electron manages the Microsoft-specific auth flow.
