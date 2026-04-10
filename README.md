# Plutus Investment Group Dashboard

Static HTML/CSS/JavaScript dashboard with a Tauri desktop shell.

## Project Layout

```text
.
├── app/                  # Canonical web app source
│   ├── data/             # Seed data + share drive config
│   ├── index.html        # Redirects to the default page
│   ├── pages/            # HTML screens
│   ├── scripts/          # Browser runtime code
│   └── styles/           # Shared and page CSS
├── build/web/            # Generated web bundle for Tauri
├── config/               # Shared project path/config metadata
├── src-tauri/            # Tauri desktop wrapper
└── tools/                # Node-side project tooling
```

## Source Of Truth

- Core app code lives in `app/`.
- Generated deployment assets live in `build/` and `dist/`.
- The desktop shell should consume generated assets, not duplicate source edits.

## Core Files

- `app/index.html`: root redirect
- `app/pages/*.html`: app pages
- `app/scripts/shared/app-config.js`: shared route, page, storage, and data-file manifest
- `app/scripts/shared/layout.js`: shared sidebar/nav renderer
- `app/scripts/shared/core.js`: runtime data and sync layer
- `app/data/config.json`: dashboard configuration
- `app/data/sharedrive-tasks.json`: ShareDrive sync configuration

## Build Flow

- `npm run web:prepare`
  - Copies `app/` into `build/web/`
- `npm run desktop`
  - Prepares `build/web/` and runs the Tauri desktop app
- `npm run dist`
  - Prepares `build/web/` and builds the Tauri desktop bundle

## Tooling

- `tools/prepare-web-assets.cjs`: generates the deployable web bundle
- `tools/generate-config.js`: regenerates `app/data/config.json` from `.env`
- `tools/generate-brand-icons.py`: regenerates Tauri app icons
- `tools/debug-sharedrive-json.cjs`: inspects synced ShareDrive JSON files

## Sharedrive Notes

- Update `app/data/sharedrive-tasks.json` to configure task, deals, or config sync.
- When shared sync is enabled, app data is read from the shared file or the desktop runtime store instead of local-only browser storage.
- The Tauri desktop app exposes the `window.PlutusDesktop` bridge used for Microsoft Graph and ShareDrive features.

## Run

- Install dependencies: `npm install`
- Launch desktop app: `npm run desktop`
- Build the web bundle only: `npm run web:prepare`
- Build the desktop package: `npm run dist`

## Documentation

- Read the Docs / MkDocs source lives in `docs/`
- Local config lives in `mkdocs.yml`
- Read the Docs build config lives in `.readthedocs.yaml`
- Local docs commands:
  - `python3 -m pip install -r docs/requirements.txt`
  - `npm run docs:serve`
  - `npm run docs:build`

## Important Convention

Edit `app/` for product changes.
Treat `build/` and `dist/` as generated deployment output.
