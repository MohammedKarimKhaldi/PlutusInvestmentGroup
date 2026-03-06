# Plutus Investment Group Dashboard

Static HTML/CSS/JavaScript dashboard for:
- investor tracking
- deals overview and deal management
- task management by owner and deal

## Main Entrypoint

- `index.html` (redirects to `public/investor-dashboard.html`)

## Current Structure

```text
.
├── index.html
├── README.md
├── data/
│   ├── config.js
│   ├── deals.js
│   └── tasks.js
├── public/
│   ├── investor-dashboard.html
│   ├── deals-overview.html
│   ├── deal-details.html
│   ├── tasks-management.html
│   └── owner-tasks.html
├── scripts/
│   ├── generate-config.js
│   ├── shared/
│   │   └── core.js
│   └── pages/
│       ├── investor-dashboard.js
│       ├── deals-overview.js
│       ├── deal-details.js
│       ├── tasks-management.js
│       └── owner-tasks.js
└── styles/
    ├── base.css
    ├── components.css
    ├── investor-dashboard.css
    ├── deal-details.css
    ├── tasks-management.css
    └── owner-tasks.css
```

## Page Map

- `public/investor-dashboard.html`: investor dashboard
- `public/deals-overview.html`: pipeline table for all deals
- `public/deal-details.html`: single deal details, edits, and related tasks
- `public/tasks-management.html`: cross-owner management view for tasks
- `public/owner-tasks.html`: owner-specific task view

## Data and Persistence

- Seed data:
  - `data/deals.js`
  - `data/tasks.js`
- Dashboard configuration:
  - `data/config.js`
- Browser LocalStorage keys:
  - deals: `deals_data_v1`
  - tasks: `owner_tasks_v1`
- Shared utilities:
  - `scripts/shared/core.js`

## Run

Open `index.html` in your browser.  

## Desktop (macOS via Electron)

- Install dependencies: `npm install`
- Run app: `npm run desktop`
- Build macOS package: `npm run dist:mac`

## iOS (via Capacitor + Xcode)

- Prepare web assets: `npm run web:prepare`
- Add iOS project (first time only): `npm run ios:add`
- Sync latest web assets to iOS: `npm run ios:sync`
- Open in Xcode: `npm run ios:open`
- One-command sync + open: `npm run ios:run`

## Notes

- Edits to deals/tasks are stored in browser LocalStorage.
- `scripts/generate-config.js` is optional and only used when regenerating `data/config.js` from a local `.env`.
