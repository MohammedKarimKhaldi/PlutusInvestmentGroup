const fs = require('fs');
const path = require('path');

const cwd = process.cwd();
const envPath = path.join(cwd, '.env');
const outputPath = path.join(cwd, 'data', 'config.js');

function parseEnv(content) {
  const result = {};
  const lines = content.split(/\r?\n/);
  for (const raw of lines) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const eq = line.indexOf('=');
    if (eq === -1) continue;
    const key = line.slice(0, eq).trim();
    let value = line.slice(eq + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    result[key] = value;
  }
  return result;
}

if (!fs.existsSync(envPath)) {
  console.error('Missing .env file. Create it from .env.example first.');
  process.exit(1);
}

const env = parseEnv(fs.readFileSync(envPath, 'utf8'));

const required = [
  'BIOLUX_EXCEL_URL',
  'WOD_EXCEL_URL',
  'IQ500_EXCEL_URL',
  'PROXY_1_BASE',
  'PROXY_2_BASE'
];

const missing = required.filter((key) => !env[key]);
if (missing.length) {
  console.error(`Missing required keys in .env: ${missing.join(', ')}`);
  process.exit(1);
}

const configObj = {
  dashboards: [
    {
      id: 'biolux',
      name: 'Biolux Analytics',
      description: 'Investor Network Live Dashboard',
      excelUrl: env.BIOLUX_EXCEL_URL,
      sheets: {
        funds: 'Biolux Investors (Funds)',
        familyOffices: 'Biolux Investors (F.Os)',
        figures: 'figure'
      }
    },
    {
      id: 'wod',
      name: 'William Oak Diagnostics Analytics',
      description: 'Investor Network Live Dashboard',
      excelUrl: env.WOD_EXCEL_URL,
      sheets: {
        funds: 'WOD Investors (Funds)',
        familyOffices: 'WOD Investors (F.Os)',
        figures: 'figure'
      }
    },
    {
      id: 'IQ500',
      name: 'IQ500 Analytics',
      description: 'This is an example dashboard configuration.',
      excelUrl: env.IQ500_EXCEL_URL,
      sheets: {
        funds: 'IQ500 Investors (Funds)',
        familyOffices: 'IQ500 (F.Os)',
        figures: 'FiguresSheet'
      }
    }
  ],
  settings: {
    defaultDashboard: 'biolux',
    allowLocalUpload: true,
    title: 'Investor Dashboard'
  }
};

const output = [
  '// Generated from .env by scripts/generate-config.js',
  `window.DASHBOARD_CONFIG = ${JSON.stringify(configObj, null, 4)};`,
  '',
  'window.DASHBOARD_PROXIES = [',
  `    (url) => ${JSON.stringify(env.PROXY_1_BASE)} + encodeURIComponent(url),`,
  `    (url) => ${JSON.stringify(env.PROXY_2_BASE)} + encodeURIComponent(url)`,
  '];',
  ''
].join('\n');

fs.writeFileSync(outputPath, output);
console.log(`✓ Generated ${outputPath}`);
