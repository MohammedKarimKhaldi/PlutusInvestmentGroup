# Plutus Investment Group Dashboard

This documentation describes the structure, runtime behavior, and maintenance workflow for the Plutus Investment Group dashboard codebase.

The project is a static browser application wrapped by Tauri for desktop use.

## What this docs site covers

- How the repository is organized after the source/deployment cleanup
- How the browser app boots and where shared logic lives
- How the Tauri desktop bridge exposes local storage and Microsoft Graph features to the app
- How configuration and Sharedrive sync work
- How to build, package, and document the project

## Current architecture at a glance

```text
app/           Canonical browser application source
build/web/     Generated web bundle consumed by Tauri
src-tauri/     Desktop shell and Graph / filesystem bridge
tools/         Node-side project tooling
config/        Shared path and filename metadata used by tooling/runtime
```

## Core principles

1. `app/` is the source of truth for product code.
2. `build/web/` is generated output, not hand-edited source.
3. Shared routes, filenames, and storage keys should live in one place whenever possible.
4. Generated output should be treated as disposable build output.

## Read this first

- If you are onboarding to the repo, start with [Getting Started](getting-started.md).
- If you are changing architecture or file layout, read [Repository Layout](architecture/repository-layout.md).
- If you are debugging runtime behavior, read [Browser Runtime](architecture/browser-runtime.md) and [Desktop Bridge](architecture/desktop-bridge.md).
- If you are changing sync or configuration, read [Configuration](product/configuration.md) and [Sharedrive Sync](product/sharedrive-sync.md).
