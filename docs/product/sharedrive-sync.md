# Sharedrive Sync

## What sync covers

The app can sync three JSON domains through ShareDrive / Microsoft Graph:

- tasks
- deals
- dashboard config

## Entry point

The main user-facing entry point is the Sharedrive folders page:

- sign in
- browse folders
- inspect diagnostics
- upload files

## Runtime behavior

`app/scripts/shared/core.js` contains the browser-side orchestration for:

- reading sync configuration
- refreshing tasks from ShareDrive
- refreshing deals from ShareDrive
- refreshing dashboard config from ShareDrive
- uploading changed task data

## Desktop support

On Tauri, the browser layer calls `window.PlutusDesktop` methods wired up in `app/scripts/shared/app-config.js`.

The Tauri backend then performs:

- Graph token resolution
- download and upload requests
- persisted session reuse

## Browser support

When desktop APIs are unavailable, browser code still supports limited device-code based auth flows using values from `sharedrive-tasks.json`.

## Important debugging paths

If sync fails, check:

1. `app/data/sharedrive-tasks.json`
2. `app/data/config.json`
3. the Sharedrive folders diagnostics page
4. `tools/debug-sharedrive-json.cjs`
5. Tauri logs from `src-tauri/src/main.rs`

## Common failure modes

- Share URL points to a folder when the code expects a file
- SharePoint returns HTML instead of JSON or workbook content
- access token is missing or expired
- native platform is running stale copied web assets because sync was not rerun
