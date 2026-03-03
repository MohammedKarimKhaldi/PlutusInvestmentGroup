# PlutusInvestmentGroup Dashboard

A professional investment deal management and task tracking dashboard built with vanilla JavaScript, HTML, and CSS.  
**Completely standalone – open the HTML files in a browser; no build or server required.**


## Project Structure

```
PlutusInvestmentGroup/
├── public/                    # Static HTML pages (primary entrypoint)
│   ├── main.html             # Deals overview table
│   ├── deal.html             # Individual deal details
│   ├── tasks.html            # Tasks grouped by owner
│   └── index.html            # Main investor dashboard (large file)
├── legacy/                    # Original HTML files (moved for reference)
│   ├── main.html
│   ├── deal.html
│   ├── tasks.html
│   └── person.html
├── data/                      # Data files and configuration
│   ├── deals.js              # Deal data store
│   ├── tasks.js              # Task data store
│   └── config.js             # Generated configuration (from .env)
├── styles/                    # CSS stylesheets
│   ├── common.css            # Shared styles and design system
│   └── pages.css             # Page-specific component styles
├── scripts/                   # Build and utility scripts
│   ├── generate-config.js    # Script to generate config.js from .env
│   └── generate-dashboard-config.js  # Legacy script (deprecated)
├── .env.example              # Environment template
├── .gitignore                # Git ignore rules
├── package.json              # Project metadata and dependencies
└── README.md                 # This file
```

## Getting Started

### 1. Setup Environment

Copy the example environment file and configure it:

```bash
cp .env.example .env
```

Edit `.env` and add your actual Excel URLs and proxy settings:

```env
BIOLUX_EXCEL_URL=https://...
WOD_EXCEL_URL=https://...
IQ500_EXCEL_URL=https://...
PROXY_1_BASE=https://api.codetabs.com/v1/proxy?quest=
PROXY_2_BASE=https://api.allorigins.win/get?url=
```

### 2. Configuration

The repository already includes a `data/config.js` file with example settings. You can modify the values directly in
`data/config.js` without running any build step – the app will load it automatically.

If you still want to regenerate the file from an `.env` template, the original build script is available in 
`scripts/generate-config.js`, but using it is **entirely optional**.

```bash
# optional: update config from .env
node scripts/generate-config.js
```



### 3. Serve the Application (or open directly)

Use any static server to serve the files:

```bash
# Using Python 3
python -m http.server 8000

# Using Node.js http-server
npx http-server

# Using VS Code Live Server extension
# Right-click and "Open with Live Server"
```

Then open `http://localhost:8000/public/main.html` (or `index.html`) in your browser.

Alternatively, simply open any of the HTML files directly using your file manager or `file://` URI; they are fully standalone.
## Features

- **Deal Pipeline** (`main.html`) - Overview of all active deals with stage tracking
- **Deal Details** (`deal.html`) - Individual deal information and related tasks
- **Task Management** (`tasks.html`) - Tasks grouped by owner with filtering and add functionality
- **Investor Dashboard** (`index.html`) - Fundraising analytics and investor tracking
- **Person View** (`person.html`) - Tasks filtered by specific team member

## Architecture

### Styling System

The project uses a centralized design system with CSS variables:

- **common.css** - Global styles, typography, layout, and component base styles
- **pages.css** - Page-specific components and variations

This eliminates CSS duplication across all HTML pages.

### Data Management

Data is stored in simple JavaScript files and localStorage:

- `data/deals.js` - Array of deal objects
- `data/tasks.js` - Array of task objects
- `localStorage` - Client-side persistence for added/modified tasks

### Configuration

The `data/config.js` file is generated from `.env` and contains:

- Dashboard definitions (Excel URLs, sheet mappings)
- CORS proxy configurations
- Application settings

**Important:** Do not edit `data/config.js` manually. It's regenerated from `.env` using the build script.

## Development

### Adding Data

Edit the data files directly:

```javascript
// data/deals.js
const DEALS = [
  {
    id: "your-deal-id",
    name: "Deal Name",
    company: "Company Name",
    // ... other properties
  }
];
```

### Styling

Add global styles to `styles/common.css` and page-specific styles to `styles/pages.css`.

### Building

Regenerate the configuration when `.env` changes:

```bash
npm run build
```

Or directly:

```bash
node scripts/generate-config.js
```

## Browser Support

Works in all modern browsers (Chrome, Firefox, Safari, Edge) that support:

- ES6 JavaScript
- CSS Grid and Flexbox
- LocalStorage API
- Fetch API (for index.html dashboard features)

## Files Not Included

The following files from the original structure are intentionally excluded from this reorganization:

- **Original root HTML files** - Replaced with optimized versions in `public/`
- **dashboard.config.js** - legacy file (no longer used)

## Best Practices

1. ✅ Keep data files (`data/`) separate from presentation (`public/`)
2. ✅ Centralize styles in `styles/` to avoid duplication
3. ✅ Generate configuration from `.env` rather than committing sensitive URLs
4. ✅ Use meaningful IDs for deals and tasks
5. ✅ Document custom properties in data objects

## Troubleshooting

**Config file not found?**
- Run `node scripts/generate-config.js`
- Check `.env` exists with required keys

**Styles not applying?**
- Verify stylesheet paths in HTML `<head>`
- Check browser console for 404 errors

**Data not persisting?**
- Browser localStorage is disabled or full
- Check browser console for errors

## License

Internal use only.