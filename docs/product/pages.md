# Pages

## Navigation pages

The shared page map is defined in `app/scripts/shared/app-config.js`.

Primary navigation includes:

- Investor dashboard
- Deals overview
- Deal ownership
- Accounting
- Tasks by owner
- Sharedrive folders

## Page responsibilities

### `investor-dashboard.html`

- Displays dashboard data loaded from Excel/SharePoint sources
- Supports dashboard switching and dashboard creation
- Uses configured proxy fallbacks for workbook loading

### `deals-overview.html`

- Shows the deal pipeline
- Lets users create new deals
- Links deals to dashboards and detail pages

### `deal-details.html`

- Shows one deal in detail
- Displays stage progress
- Shows related tasks and deal metadata

### `tasks-management.html`

- Cross-owner view of tasks
- Links to deal and owner-specific contexts

### `owner-tasks.html`

- Focused task view for a specific owner

### `sharedrive-folders.html`

- Entry page for ShareDrive connectivity
- Device-code sign-in
- Folder browsing
- Upload and diagnostics tools

### `deal-ownership.html`

- Loads and visualizes staffing / ownership style workbook data

### `accounting.html`

- Accounting-related view for deal retainers and payment tracking

## Page scripts

Each page has a matching script in `app/scripts/pages/`.

Shared behavior should stay in `app/scripts/shared/`, with page scripts limited to page-specific rendering and interactions.
