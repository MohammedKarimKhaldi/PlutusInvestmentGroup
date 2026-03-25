# Build and Deployment

## Source to output flow

```text
app/
  -> npm run web:prepare
  -> build/web/
  -> Electron package or Capacitor sync
```

## Web bundle generation

Command:

```bash
npm run web:prepare
```

Implementation:

- `tools/prepare-web-assets.cjs`

Shared paths:

- `config/project-paths.cjs`

## Desktop packaging

Common commands:

```bash
npm run desktop
npm run dist:mac
npm run dist:win
npm run dist:linux
```

The package build includes:

- `build/web/**/*`
- `config/**/*`
- `electron/**/*`

## GitHub Actions

- `.github/workflows/windows-release.yml`
  - Pull requests to `main` and pushes to `main` build the Windows installer for CI validation.
  - Version tag pushes such as `v1.0.2` publish the Windows installer and updater metadata to GitHub Releases.
  - If Windows signing secrets are configured, the tag release is signed; otherwise it is published unsigned.

## Capacitor sync

Commands:

```bash
npm run ios:sync
npm run android:sync
```

These commands prepare the web bundle first, then sync it into native projects.

## Important operational rule

If you change code in `app/`, rerun the relevant sync or build command before assuming the platform app reflects those changes.

## Generated folders

Treat these as generated output:

- `build/`
- copied web assets inside native platform folders
- `dist/`

## Config regeneration

If you use an environment-driven config workflow:

```bash
node tools/generate-config.js
```

This writes `app/data/config.json` from `.env`.
