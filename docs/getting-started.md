# Getting Started

## Prerequisites

- Node.js and npm
- Python 3 if you want to build this documentation site locally
- Xcode for iOS work
- Android Studio for Android work

## Install project dependencies

```bash
npm install
```

## Main development commands

### Prepare the generated web bundle

```bash
npm run web:prepare
```

This copies `app/` into `build/web/`.

### Run the Electron desktop app

```bash
npm run desktop
```

This automatically prepares `build/web/` first.

### Sync to iOS

```bash
npm run ios:sync
```

### Sync to Android

```bash
npm run android:sync
```

## Where to edit code

- Edit browser app source in `app/`
- Edit desktop integration in `electron/`
- Edit Node tooling in `tools/`
- Edit shared project path metadata in `config/project-paths.cjs`

Do not treat `build/web/` as source code.

## Common first checks

When the app behaves unexpectedly, the fastest checks are usually:

1. Run `npm run web:prepare` to make sure generated output matches source.
2. Check the relevant JSON in `app/data/`.
3. Check whether Electron-only behavior depends on `window.PlutusDesktop`.
4. Check whether Sharedrive sync is enabled and overriding local seed data.
