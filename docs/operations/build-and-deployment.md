# Build and Deployment

## Source to output flow

```text
app/
  -> npm run web:prepare
  -> build/web/
  -> Tauri desktop build
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
npm run desktop:build
npm run dist
```

Tauri bundle icons live in `src-tauri/icons/`.

## GitHub release publishing

Once this folder is connected to a GitHub repository, use the workflow at `.github/workflows/publish.yml`.

What it does:

- builds macOS Apple Silicon bundles
- builds macOS Intel bundles
- builds Windows bundles
- uploads all generated installers to a GitHub Release

Recommended release flow:

```bash
git push origin main
git tag v1.0.19
git push origin v1.0.19
```

Notes:

- the workflow uses the app version from `src-tauri/tauri.conf.json` / `src-tauri/Cargo.toml`
- unsigned macOS and Windows builds can still be downloaded, but users may see Gatekeeper or SmartScreen warnings until code signing is added
- this local folder is not currently a git repository, so pushing must happen after it is linked to a remote repository

## Important operational rule

If you change code in `app/`, rerun the relevant build command before assuming the desktop app reflects those changes.

## Generated folders

Treat these as generated output:

- `build/`
- `dist/`

## Config regeneration

If you use an environment-driven config workflow:

```bash
node tools/generate-config.js
```

This writes `app/data/config.json` from `.env`.
