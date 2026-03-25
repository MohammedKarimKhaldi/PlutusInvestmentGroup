# Plutus Investment Group Dashboard

Static HTML/CSS/JavaScript dashboard with Electron desktop packaging and Capacitor mobile deployment.

## Project Layout

```text
.
├── app/                  # Canonical web app source
│   ├── data/             # Seed data + share drive config
│   ├── index.html        # Redirects to the default page
│   ├── pages/            # HTML screens
│   ├── scripts/          # Browser runtime code
│   └── styles/           # Shared and page CSS
├── build/web/            # Generated web bundle for Electron + Capacitor
├── config/               # Shared project path/config metadata
├── electron/             # Desktop wrapper
├── android/              # Capacitor Android project
├── ios/                  # Capacitor iOS project
└── tools/                # Node-side project tooling
```

## Source Of Truth

- Core app code lives in `app/`.
- Generated deployment assets live in `build/` and native platform output folders.
- Desktop, iOS, and Android should consume generated assets, not duplicate source edits.

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
  - Prepares `build/web/` and opens Electron
- `npm run ios:sync`
  - Prepares `build/web/` and syncs Capacitor iOS
- `npm run android:sync`
  - Prepares `build/web/` and syncs Capacitor Android

## Tooling

- `tools/prepare-web-assets.cjs`: generates the deployable web bundle
- `tools/generate-config.js`: regenerates `app/data/config.json` from `.env`
- `tools/debug-sharedrive-json.cjs`: inspects synced ShareDrive JSON files

## Sharedrive Notes

- Update `app/data/sharedrive-tasks.json` to configure task, deals, or config sync.
- When shared sync is enabled, app data is read from the shared file or the desktop runtime store instead of local-only browser storage.
- Electron also supports Microsoft Graph device-code auth through the preload bridge.

## Run

- Install dependencies: `npm install`
- Launch desktop app: `npm run desktop`
- Build the web bundle only: `npm run web:prepare`
- Build desktop package: `npm run dist:mac`
- Build Windows installer locally: `npm run dist:win`

## Windows Releases

- Windows builds now use an `NSIS` installer so installed copies can auto-update.
- Auto-update checks run only in packaged Windows builds published from GitHub Releases.
- GitHub Actions workflow: `.github/workflows/windows-release.yml`
- CI validation runs for Windows on pull requests to `main`, pushes to `main`, and manual workflow runs.
- CD release publishing runs when you push a version tag such as `v1.0.2`.
- Tagged Windows releases sign the installer when signing secrets are available, and otherwise fall back to an unsigned release.
- Optional GitHub secrets for signed Windows releases:
  - `WIN_SIGN_CERT_BASE64`
  - `WIN_SIGN_CERT_PASSWORD`
- `WIN_SIGN_CERT_BASE64` should contain the base64 text of an exported `.pfx` or `.p12` code-signing certificate.
- On macOS you can copy the certificate as a single-line base64 string with:
  - `base64 -i your-cert.pfx | tr -d '\n' | pbcopy`
- The release workflow decodes that secret into a temporary certificate file on the runner and passes it to Electron Builder using `WIN_CSC_LINK` and `WIN_CSC_KEY_PASSWORD`.
- If the secrets are missing, the workflow still publishes the release, but Windows will show `Unknown publisher` and SmartScreen warnings are more likely.
- A standard code-signing certificate removes the `Unknown publisher` label, but SmartScreen can still warn until the app builds reputation.
- An EV certificate builds trust faster, but the common USB-token EV format usually cannot be exported for GitHub Actions.
- Packaged desktop startup logs are written to Electron's `userData` folder as `desktop-runtime.log`.
- To force DevTools open in a packaged build for debugging, launch with `PLUTUS_OPEN_DEVTOOLS=1`.
- Electron Builder build resources now live in `electron-builder-resources/`, which keeps the generated `build/web/` app bundle available for packaging.
- To publish a new Windows version:
  - Update `package.json` `version`
  - Commit and push to `main`
  - Create and push a matching tag such as `v1.0.1`
  - GitHub Actions runs `.github/workflows/windows-release.yml`
  - The release job publishes a GitHub Release and uploads the installer, blockmap, and `latest.yml` updater metadata
- Installed Windows apps will detect the new release and prompt the user to restart after the update downloads.

## Mac App Store

- Mac App Store builds use Electron Builder's `mas` target with sandbox entitlements in `electron/entitlements.mas.plist`.
- Build a Mac App Store package locally with `npm run dist:mas:arm64` or `npm run dist:mas:x64`.
- Use `npm run dist:mas:dev:arm64` or `npm run dist:mas:dev:x64` for development-signed MAS test builds.
- MAS builds keep writable app data inside the app sandbox even if `team-store-path.json` sets a custom `storeDir`.
- GitHub Actions workflow: `.github/workflows/mac-app-store-release.yml`
- Required GitHub secrets for the MAS workflow:
  - `APPLE_MAS_APP_CERT_BASE64`
  - `APPLE_MAS_APP_CERT_PASSWORD`
  - `APPLE_MAS_INSTALLER_CERT_BASE64`
  - `APPLE_MAS_INSTALLER_CERT_PASSWORD`
  - `APPLE_MAS_PROVISIONING_PROFILE_BASE64`
- Apple-side setup is still required:
  - Create the macOS app record in App Store Connect
  - Enable App Sandbox for the app ID and provisioning profile
  - Sign with Mac App Store-compatible certificates/profiles
  - Upload the resulting MAS artifact with Transporter or Xcode Organizer

## Xcode Cloud

- Xcode Cloud must build `ios/App/App.xcworkspace`, not `ios/App/App.xcodeproj`.
- Use scheme `App` for both workflows.
- Create separate archive workflows for:
  - iOS
  - macOS using Mac Catalyst
- The repo's Xcode Cloud scripts run `npm run ios:sync` and rely on CocoaPods integration.
- If Xcode Cloud is pointed at `App.xcodeproj`, builds fail with missing Capacitor modules during archive.
- If a workflow still logs `-project /Volumes/workspace/repository/ios/App/App.xcodeproj`, edit or recreate it so the container is `ios/App/App.xcworkspace`.
- If a build fails during `npm ci` or `npm install` with `ENOTFOUND registry.npmjs.org`, that is an npm registry/DNS connectivity problem in the Xcode Cloud environment, not an Xcode compile error.
- In that case, retry the build first. If it persists, set `XCLOUD_NPM_REGISTRY` in the workflow Environment settings to a reachable registry URL and rerun.

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
Treat `build/`, `dist/`, and native copied web assets as generated deployment output.
