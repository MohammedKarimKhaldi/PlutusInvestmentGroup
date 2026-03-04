# Plutus Investment Group Dashboard

Internal web dashboard for managing fundraising deals, investor tracking, and execution tasks.

The project is intentionally lightweight:
- Vanilla HTML/CSS/JavaScript
- No framework build pipeline
- LocalStorage-backed editing for deals and tasks

## Overview

The app is split into focused pages:
- `public/index.html`: investor dashboard synced from configured Excel sources
- `public/main.html`: deals overview and pipeline table
- `public/deal.html`: deal detail page, deal editing, task management, dashboard-sync task generation
- `public/tasks.html`: management view across tasks with grouping/filtering/sorting
- `public/person.html`: owner-level task view with grouped/list layouts and inline editing

## Project Structure

```text
.
├── data/
│   ├── config.js             # Dashboard/proxy configuration (generated or manually maintained)
│   ├── deals.js              # Seed deal dataset
│   └── tasks.js              # Seed task dataset
├── public/
│   ├── index.html
│   ├── main.html
│   ├── deal.html
│   ├── tasks.html
│   └── person.html
├── scripts/
│   ├── generate-config.js    # Optional config generation from .env
│   ├── shared/
│   │   └── core.js           # Shared app utilities (storage, normalization, lookup helpers)
│   └── pages/
│       ├── index-dashboard.js
│       ├── main.js
│       ├── deal.js
│       ├── tasks.js
│       └── person.js
├── styles/
│   ├── common.css
│   ├── pages.css
│   ├── index-dashboard.css
│   ├── deal.css
│   ├── tasks.css
│   └── person.css
├── .env.example
├── package.json
└── README.md
```

## Data and Persistence

- Static defaults come from `data/deals.js` and `data/tasks.js`.
- Runtime edits are persisted in browser LocalStorage:
  - deals: `deals_data_v1`
  - tasks: `owner_tasks_v1`
- Shared LocalStorage and lookup logic is centralized in `scripts/shared/core.js`.

## Setup

### 1. Install dependencies (optional, only needed for `serve`)

```bash
npm install
```

### 2. Configure dashboards (optional)

If you need custom Excel/proxy endpoints:

```bash
cp .env.example .env
npm run build
```

This regenerates `data/config.js` using `scripts/generate-config.js`.

You can also edit `data/config.js` directly for quick local changes.

### 3. Run locally

Option A:

```bash
npm run start
```

Option B:

```bash
npm run serve
```

Then open:
- `http://localhost:8000/public/index.html`
- `http://localhost:8000/public/main.html`

## Development Guidelines

- Keep page-specific behavior in `scripts/pages/*`.
- Put reusable logic in `scripts/shared/core.js` (no duplication of storage helpers).
- Keep design tokens/layout primitives in `styles/common.css`.
- Use page stylesheet files for page-only visuals.
- Preserve data object compatibility (`id`, `dealId`, `status`, `owner`, etc.) when extending features.

## Current Status Model

Task statuses supported in UI:
- `in progress`
- `waiting`
- `done`

Deal pipeline stages currently used in overview/detail:
- `prospect`
- `onboarding`
- `contacting investors`

## Troubleshooting

- If dashboards do not sync: verify `data/config.js` URLs/proxies and browser console errors.
- If edits are not reflected: clear LocalStorage keys `deals_data_v1` / `owner_tasks_v1`.
- If local file loading is blocked: run via local server (`npm run start`) instead of opening `file://` directly.

## License

Internal use only.
