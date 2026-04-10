# Configuration

## Main JSON files

The main tracked configuration files live in `app/data/`.

## `config.json`

This file controls browser-side dashboard behavior.

Typical responsibilities:

- application-level settings
- proxy fallback list
- default dashboard selection
- shared deals sync settings
- optional dashboard definitions and overrides

### Important sections

#### `settings`

May contain values such as:

- `defaultDashboard`
- `allowLocalUpload`
- `title`
- `userDirectory`
- `sharedDeals`

#### `proxies`

List of proxy base URLs used by the investor dashboard when direct workbook loading is blocked.

## `sharedrive-tasks.json`

This file controls Sharedrive-hosted JSON sync.

Top-level sections:

- `tasks`
- `deals`
- `config`

Each section can define:

- `enabled`
- `shareUrl`
- `downloadUrl`
- `fileName`
- `parentItemId`
- `pollIntervalMs`

The `tasks` section also carries browser auth metadata:

- `azureClientId`
- `azureTenantId`
- `graphScopes`

## Centralized filename and key metadata

Shared names should come from code, not be repeated ad hoc.

Current sources of truth:

- `config/project-paths.cjs`
  - file and directory names used by Node-side tooling
- `app/scripts/shared/app-config.js`
  - page filenames, storage keys, and browser data filenames

## Local persistence

Browser/desktop persistence keys are centralized in `app-config.js`.

Key examples:

- `deals_data_v1`
- `owner_tasks_v1`
- `sharedrive_connected_v1`
- `plutus_graph_session_v1`
- `plutus_dashboard_config_v1`
