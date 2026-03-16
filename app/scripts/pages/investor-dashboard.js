                // --- CONFIGURATION ---
                // Sensitive links are loaded from config.json
                const DEFAULT_DASHBOARDS_CONFIG = window.DASHBOARD_CONFIG || {
                    dashboards: [],
                    settings: {
                        defaultDashboard: "",
                        allowLocalUpload: true,
                        title: "Investor Dashboard"
                    }
                };

                const DEFAULT_SHEET_NAMES = { funds: 'funds', familyOffices: 'f.o.', figures: 'figure' };

                function getMergedDashboardsConfig() {
                    if (window.AppCore && typeof window.AppCore.getDashboardConfig === "function") {
                        return window.AppCore.getDashboardConfig();
                    }
                    return DEFAULT_DASHBOARDS_CONFIG;
                }

                let dashboardsConfig = getMergedDashboardsConfig();
                let currentDashboard = null;

                // Proxy fallback chain loaded from external config if available
                const PROXIES = window.DASHBOARD_PROXIES || [];

                // --- STATE ---
                let rawData = { vc: [], fo: [] };
                let activeType = 'vc';
                let charts = {};
                let activeFilters = { all: true, calls: false, meetings: false, forward: false };
                let typeColors = {};
                let chartDrillDownStack = [];
                const COLOR_PALETTE = ['#818cf8', '#34d399', '#fbbf24', '#a78bfa', '#f472b6', '#22d3ee', '#fb7185'];

                function buildPageUrl(pageId, params) {
                    if (window.AppCore && typeof window.AppCore.getPageUrl === "function") {
                        return window.AppCore.getPageUrl(pageId, params);
                    }
                    if (window.PlutusAppConfig && typeof window.PlutusAppConfig.buildPageHref === "function") {
                        return window.PlutusAppConfig.buildPageHref(pageId, params);
                    }
                    const query = new URLSearchParams(params || {});
                    const queryString = query.toString();
                    return queryString ? `${pageId}.html?${queryString}` : `${pageId}.html`;
                }

                function looksLikeSharePointLink(url) {
                    if (!url) return false;
                    const normalized = String(url).toLowerCase();
                    return (
                        normalized.includes("sharepoint.com/:") ||
                        normalized.includes("sharepoint.com/_layouts/15/doc.aspx") ||
                        normalized.includes("1drv.ms/")
                    );
                }

                async function resolveShareLinkDownloadUrl(excelUrl) {
                    if (!looksLikeSharePointLink(excelUrl)) return null;
                    if (window.AppCore && typeof window.AppCore.resolveShareDriveDownloadUrl === "function") {
                        try {
                            return await window.AppCore.resolveShareDriveDownloadUrl(excelUrl);
                        } catch (error) {
                            console.warn("[Dashboard] SharePoint URL resolution unavailable, falling back to proxy sync", error);
                            return null;
                        }
                    }
                    return null;
                }

                async function fetchWorkbookFromUrl(fetchUrl, expectsJsonWrapper) {
                    let buffer;
                    if (expectsJsonWrapper) {
                        const response = await fetch(fetchUrl, { cache: 'no-store' });
                        if (!response.ok) throw new Error("Download failed");
                        const json = await response.json();
                        if (!json.contents) throw new Error("No contents in JSON");

                        const b64 = json.contents.split(',')[1] || json.contents;
                        const binaryString = atob(b64);
                        buffer = new ArrayBuffer(binaryString.length);
                        const view = new Uint8Array(buffer);
                        for (let i = 0; i < binaryString.length; i++) view[i] = binaryString.charCodeAt(i);
                    } else {
                        buffer = window.AppCore && typeof window.AppCore.downloadBinary === "function"
                            ? await window.AppCore.downloadBinary(fetchUrl, { cache: 'no-store' })
                            : await fetch(fetchUrl, { cache: 'no-store' }).then((response) => {
                                if (!response.ok) throw new Error("Download failed");
                                return response.arrayBuffer();
                            });
                    }

                    return XLSX.read(buffer, { type: 'array' });
                }

                // --- INITIALIZATION ---
                window.addEventListener('load', () => {
                    setupUiBindings();
                    initializeDashboard();
                });

                function setupUiBindings() {
                    const openLocalBtn = document.getElementById('btn-open-local-file');
                    const fileInput = document.getElementById('file-input');
                    const retryBtn = document.getElementById('btn-retry-sync');
                    const dashboardSwitcher = document.getElementById('dashboard-switcher');
                    const searchInput = document.getElementById('search');
                    const dashboardForm = document.getElementById('dashboard-form');

                    if (openLocalBtn && fileInput) {
                        openLocalBtn.addEventListener('click', () => fileInput.click());
                    }
                    if (fileInput) {
                        fileInput.addEventListener('change', (event) => loadLocalFile(event));
                    }
                    if (retryBtn) {
                        retryBtn.addEventListener('click', retrySync);
                    }
                    if (dashboardSwitcher) {
                        dashboardSwitcher.addEventListener('change', onDashboardChange);
                    }
                    if (searchInput) {
                        searchInput.addEventListener('input', filterData);
                    }
                    if (dashboardForm) {
                        dashboardForm.addEventListener('submit', onAddDashboard);
                    }
                    const exportBtn = document.getElementById('btn-export-html');
                    if (exportBtn) {
                        exportBtn.addEventListener('click', exportStandaloneHtml);
                    }

                    document.querySelectorAll('[data-tab-type]').forEach((button) => {
                        button.addEventListener('click', () => {
                            const type = button.getAttribute('data-tab-type');
                            if (type) switchTab(type, button);
                        });
                    });

                    document.querySelectorAll('[data-filter]').forEach((button) => {
                        button.addEventListener('click', () => {
                            const filter = button.getAttribute('data-filter');
                            if (filter) toggleFilter(filter);
                        });
                    });
                }

                async function initializeDashboard() {
                    try {
                        if (window.AppCore && typeof window.AppCore.refreshDashboardConfigFromShareDrive === "function") {
                            await window.AppCore.refreshDashboardConfigFromShareDrive();
                        }
                        dashboardsConfig = getMergedDashboardsConfig();
                        const params = new URLSearchParams(window.location.search);
                        const requestedId = params.get('dashboard');

                        const dashboards = (dashboardsConfig && dashboardsConfig.dashboards) || [];
                        const settings = (dashboardsConfig && dashboardsConfig.settings) || {};

                        console.log('[Dashboard] Available dashboards:', dashboards.map(d => d.id));
                        console.log('[Dashboard] Requested dashboard ID from URL:', requestedId);

                        currentDashboard =
                            dashboards.find(d => d.id === requestedId) ||
                            dashboards.find(d => d.id === settings.defaultDashboard) ||
                            dashboards[0];

                        console.log('[Dashboard] Selected dashboard:', currentDashboard && currentDashboard.id, currentDashboard);

                        if (!currentDashboard) {
                            throw new Error('No dashboard configuration found');
                        }

                        // Populate dropdown
                        const switcher = document.getElementById('dashboard-switcher');
                        if (switcher) {
                            switcher.innerHTML = '';
                            dashboards.forEach(d => {
                                const opt = document.createElement('option');
                                opt.value = d.id;
                                opt.textContent = d.name || d.id;
                                if (d.id === currentDashboard.id) opt.selected = true;
                                switcher.appendChild(opt);
                            });
                        }

                        // Update page chrome from config
                        const pageTitle = settings.title || currentDashboard.name || 'Dashboard';
                        document.title = pageTitle;

                        const titleEl = document.getElementById('dashboard-title');
                        const subtitleEl = document.getElementById('dashboard-subtitle');
                        if (titleEl) titleEl.textContent = currentDashboard.name || pageTitle;
                        if (subtitleEl) subtitleEl.textContent = currentDashboard.description || 'Live Dashboard';

                        // Start sync for selected dashboard
                        console.log('[Dashboard] Starting initial sync for', currentDashboard.id, 'with URL:', currentDashboard.excelUrl);
                        startSync(0, currentDashboard.excelUrl);
                    } catch (err) {
                        console.error(err);
                        showFailure();
                    }
                }

                function onDashboardChange(event) {
                    const selectedId = event.target.value;
                    dashboardsConfig = getMergedDashboardsConfig();
                    if (!dashboardsConfig || !dashboardsConfig.dashboards) return;

                    const next = dashboardsConfig.dashboards.find(d => d.id === selectedId);
                    if (!next || (currentDashboard && next.id === currentDashboard.id)) return;

                    currentDashboard = next;

                    console.log('[Dashboard] Switched dashboard via dropdown to:', currentDashboard.id, currentDashboard);

                    // Update header text immediately for responsiveness
                    const settings = (dashboardsConfig && dashboardsConfig.settings) || {};
                    const pageTitle = settings.title || currentDashboard.name || 'Dashboard';
                    document.title = pageTitle;

                    const titleEl = document.getElementById('dashboard-title');
                    const subtitleEl = document.getElementById('dashboard-subtitle');
                    if (titleEl) titleEl.textContent = currentDashboard.name || pageTitle;
                    if (subtitleEl) subtitleEl.textContent = currentDashboard.description || 'Live Dashboard';

                    // Show loader while switching datasets
                    document.getElementById('dashboard').style.display = 'none';
                    document.getElementById('loader').style.display = 'flex';
                    const loaderMsg = document.getElementById('loader-msg');
                    if (loaderMsg) loaderMsg.textContent = 'Switching dashboard data...';

                    document.getElementById('sync-type').innerText = "Cloud Sync Active";

                    // Restart sync with new Excel URL
                    if (currentDashboard.excelUrl) {
                        console.log('[Dashboard] Restarting sync for', currentDashboard.id, 'with URL:', currentDashboard.excelUrl);
                        startSync(0, currentDashboard.excelUrl);
                    }
                }

                function normalizeDashboardId(value) {
                    const raw = String(value || "").trim().toLowerCase();
                    if (!raw) return "";
                    return raw
                        .replace(/[^a-z0-9]+/g, "-")
                        .replace(/^-+|-+$/g, "");
                }

                function setDashboardFormStatus(message, isError) {
                    const status = document.getElementById('dashboard-form-status');
                    if (!status) return;
                    status.textContent = message || "";
                    status.style.color = isError ? "#fecdd3" : "var(--text-dim)";
                }

                function upsertDashboardEntry(entry, makeDefault) {
                    if (window.AppCore && typeof window.AppCore.upsertDashboardConfigEntry === "function") {
                        return window.AppCore.upsertDashboardConfigEntry(entry, { makeDefault });
                    }

                    try {
                        const key = (
                            (window.AppCore && window.AppCore.STORAGE_KEYS && window.AppCore.STORAGE_KEYS.dashboardConfig) ||
                            (window.PlutusAppConfig && window.PlutusAppConfig.storageKeys && window.PlutusAppConfig.storageKeys.dashboardConfig) ||
                            "plutus_dashboard_config_v1"
                        );
                        const raw = localStorage.getItem(key);
                        const existing = raw ? JSON.parse(raw) : {};
                        const dashboards = Array.isArray(existing.dashboards) ? existing.dashboards.slice() : [];
                        const idx = dashboards.findIndex(d => normalizeDashboardId(d.id) === normalizeDashboardId(entry.id));
                        if (idx >= 0) {
                            dashboards[idx] = Object.assign({}, dashboards[idx], entry);
                        } else {
                            dashboards.push(entry);
                        }
                        const settings = Object.assign({}, existing.settings || {});
                        if (makeDefault) settings.defaultDashboard = entry.id;
                        localStorage.setItem(key, JSON.stringify({ dashboards, settings }));
                    } catch (err) {
                        console.error('[Dashboard] Failed to persist dashboard config', err);
                    }
                    return getMergedDashboardsConfig();
                }

                async function onAddDashboard(event) {
                    event.preventDefault();
                    const idInput = document.getElementById('dashboard-id');
                    const nameInput = document.getElementById('dashboard-name');
                    const descInput = document.getElementById('dashboard-description');
                    const urlInput = document.getElementById('dashboard-excel-url');
                    const fundsInput = document.getElementById('dashboard-sheet-funds');
                    const foInput = document.getElementById('dashboard-sheet-fo');
                    const figInput = document.getElementById('dashboard-sheet-figures');
                    const makeDefaultInput = document.getElementById('dashboard-make-default');

                    const name = String(nameInput && nameInput.value ? nameInput.value : "").trim();
                    const excelUrl = String(urlInput && urlInput.value ? urlInput.value : "").trim();
                    if (!name || !excelUrl) {
                        setDashboardFormStatus("Name and Excel link are required.", true);
                        return;
                    }

                    let dashboardId = normalizeDashboardId(idInput && idInput.value ? idInput.value : "");
                    if (!dashboardId) {
                        dashboardId = normalizeDashboardId(name);
                    }
                    if (!dashboardId) {
                        setDashboardFormStatus("Dashboard ID is required.", true);
                        return;
                    }
                    if (idInput) idInput.value = dashboardId;

                    const entry = {
                        id: dashboardId,
                        name,
                        description: String(descInput && descInput.value ? descInput.value : "").trim(),
                        excelUrl,
                        sheets: {
                            funds: String(fundsInput && fundsInput.value ? fundsInput.value : "").trim() || DEFAULT_SHEET_NAMES.funds,
                            familyOffices: String(foInput && foInput.value ? foInput.value : "").trim() || DEFAULT_SHEET_NAMES.familyOffices,
                            figures: String(figInput && figInput.value ? figInput.value : "").trim() || DEFAULT_SHEET_NAMES.figures,
                        },
                    };

                    const existing = (dashboardsConfig && dashboardsConfig.dashboards || []).find(
                        (d) => normalizeDashboardId(d.id) === dashboardId,
                    );
                    setDashboardFormStatus(existing ? "Saving dashboard..." : "Creating dashboard...", false);

                    try {
                        const updatedConfig = await upsertDashboardEntry(
                            entry,
                            Boolean(makeDefaultInput && makeDefaultInput.checked),
                        );
                        dashboardsConfig = updatedConfig;

                        setDashboardFormStatus(
                            existing ? "Dashboard updated. Reloading..." : "Dashboard created. Reloading...",
                            false,
                        );

                        window.location.href = buildPageUrl("investor-dashboard", { dashboard: dashboardId });
                    } catch (err) {
                        setDashboardFormStatus(
                            err instanceof Error ? err.message : "Failed to save dashboard.",
                            true,
                        );
                    }
                }

                async function startSync(proxyIndex, excelUrl, shareAttempted = false) {
                    if (proxyIndex >= PROXIES.length) {
                        console.error('[Dashboard] All proxies failed for URL:', excelUrl);
                        showFailure();
                        return;
                    }

                    // Cache-bust to avoid stale SharePoint responses
                    const tsParam = `ts=${Date.now()}`;
                    const liveUrl = excelUrl ? `${excelUrl}${excelUrl.includes('?') ? '&' : '?'}${tsParam}` : excelUrl;

                    const proxyFunc = PROXIES[proxyIndex];
                    const loadMsg = document.getElementById('loader-msg');
                    loadMsg.innerText = proxyIndex === 0 ? "Connecting to SharePoint..." : "Sync failed, trying backup tunnel...";

                    if (!shareAttempted) {
                        try {
                            const resolvedUrl = await resolveShareLinkDownloadUrl(liveUrl);
                            if (resolvedUrl) {
                                const wb = await fetchWorkbookFromUrl(resolvedUrl, false);
                                processWorkbook(wb, currentDashboard && currentDashboard.sheets);
                                document.getElementById('sync-type').innerText = "Live Cloud Sync Active";
                                return;
                            }
                        } catch (err) {
                            console.warn("[Dashboard] SharePoint direct sync failed, trying proxy sync", err);
                        }
                    }

                    try {
                        const fetchUrl = proxyFunc(liveUrl);
                        console.log('[Dashboard] Using proxy', proxyIndex, '->', fetchUrl);
                        const wb = await fetchWorkbookFromUrl(fetchUrl, fetchUrl.includes('allorigins'));
                        processWorkbook(wb, currentDashboard && currentDashboard.sheets);
                        document.getElementById('sync-type').innerText = "Live Cloud Sync Active";

                    } catch (err) {
                        console.warn(`[Dashboard] Proxy ${proxyIndex} failed for URL: ${excelUrl}`, err);
                        startSync(proxyIndex + 1, excelUrl, true); // Try next proxy
                    }
                }

                function showFailure() {
                    document.getElementById('loading-spinner').style.display = 'none';
                    document.getElementById('error-panel').style.display = 'block';
                }

                function retrySync() {
                    document.getElementById('error-panel').style.display = 'none';
                    document.getElementById('loading-spinner').style.display = 'block';
                    if (currentDashboard && currentDashboard.excelUrl) {
                        startSync(0, currentDashboard.excelUrl);
                    } else {
                        showFailure();
                    }
                }

                function loadLocalFile(e) {
                    const file = e.target.files[0];
                    if (!file) return;

                    document.getElementById('error-panel').style.display = 'none';
                    document.getElementById('loading-spinner').style.display = 'block';
                    document.getElementById('loader-msg').innerText = "Processing local Excel tracker...";

                    const reader = new FileReader();
                    reader.onload = function (ex) {
                        const wb = XLSX.read(ex.target.result, { type: 'array' });
                        processWorkbook(wb, currentDashboard && currentDashboard.sheets);
                        document.getElementById('sync-type').innerText = "Local Folder Mode";
                    };
                    reader.readAsArrayBuffer(file);
                }

                function processWorkbook(wb, sheetsConfig) {
                    console.log('[Dashboard] Processing workbook for dashboard:', currentDashboard && currentDashboard.id);
                    console.log('[Dashboard] Workbook sheet names:', wb.SheetNames);

                    const getSheet = (name) => {
                        if (!name) return null;
                        const matchName = wb.SheetNames.find(n => n.toLowerCase().includes(String(name).toLowerCase()));
                        return matchName ? wb.Sheets[matchName] : null;
                    };

                    // Try to auto-detect the header row by looking for the "Investor" column
                    const detectHeaderRow = (sheet, maxScanRows = 20) => {
                        if (!sheet) return 3; // fallback to row 4 (0-based index 3)
                        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
                        const upperBound = Math.min(range.e.r, maxScanRows - 1);

                        for (let r = range.s.r; r <= upperBound; r++) {
                            let rowValues = [];
                            for (let c = range.s.c; c <= range.e.c; c++) {
                                const cellAddress = XLSX.utils.encode_cell({ r, c });
                                const cell = sheet[cellAddress];
                                if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
                                    rowValues.push(String(cell.v).trim().toLowerCase());
                                }
                            }
                            if (rowValues.includes('investor')) {
                                return r; // use this row as header (0-based)
                            }
                        }
                        return 3; // default if not found
                    };

                    const sheetNames = sheetsConfig || { funds: 'funds', familyOffices: 'f.o.', figures: 'figure' };

                    const fSheet = getSheet(sheetNames.funds);
                    const oSheet = getSheet(sheetNames.familyOffices);
                    const gSheet = getSheet(sheetNames.figures);

                    console.log('[Dashboard] Resolved sheet names config:', sheetNames);
                    console.log('[Dashboard] Found sheets -> funds:', !!fSheet, 'fo:', !!oSheet, 'figures:', !!gSheet);

                    const headerRowIndex = detectHeaderRow(fSheet);
                    console.log('[Dashboard] Detected header row index (0-based):', headerRowIndex);

                    // Parse using detected header row
                    rawData.vc = XLSX.utils.sheet_to_json(fSheet, { range: headerRowIndex, defval: "" }).filter(r => r["Investor"]);
                    rawData.fo = XLSX.utils.sheet_to_json(oSheet, { range: headerRowIndex, defval: "" }).filter(r => r["Investor"]);
                    const figures = gSheet ? XLSX.utils.sheet_to_json(gSheet, { defval: "" }) : [];

                    const combined = [...rawData.vc, ...rawData.fo];

                    console.log('[Dashboard] Parsed row counts -> vc:', rawData.vc.length, 'fo:', rawData.fo.length, 'combined:', combined.length);

                    // 1. Dynamic Type Color Mapping
                    const uniqueTypes = [...new Set(combined.map(r => r["Type"]).filter(t => t))];
                    typeColors = {};
                    uniqueTypes.forEach((type, idx) => {
                        typeColors[type] = COLOR_PALETTE[idx % COLOR_PALETTE.length];
                    });

                    // 2. Robust KPIs (Exhaustive Source-Aware Count)
                    // Rule: all rows on the Funds sheet are VC. FO/HNWI are split only within the Family Offices sheet.
                    let vcCount = rawData.vc.length;
                    let hnwiCount = 0;
                    let foCount = 0;

                    const hasTypeColumn = rawData.fo.some(r => Object.prototype.hasOwnProperty.call(r, "Type"));
                    console.log('[Dashboard] Has "Type" column on FO sheet:', hasTypeColumn);

                    const classifyHNWI = (row) => {
                        const t = String(row["Type"] || "").toLowerCase();
                        const d = String(row["Description"] || "").toLowerCase();
                        return t.includes('angel') || d.includes('angel') || t.includes('individual') || d.includes('individual') || t.includes('hnwi') || d.includes('hnwi');
                    };

                    const classifyFO = (row) => {
                        const t = String(row["Type"] || "").toLowerCase();
                        const d = String(row["Description"] || "").toLowerCase();
                        return t.includes('family') || d.includes('family') || t.includes('mfo') || t.includes('sfo') || t === 'fo' || d === 'fo';
                    };

                    if (!hasTypeColumn) {
                        // If there's no "Type" column on the FO sheet at all, treat all FO rows as Family Offices
                        foCount = rawData.fo.length;
                        hnwiCount = 0;
                    } else {
                        rawData.fo.forEach(r => {
                            if (classifyHNWI(r)) hnwiCount++;
                            else foCount++; // Default for F.O sheet
                        });
                    }

                    // 3. Stage Funnel KPIs (Yes/Waiting logic)
                    function stageValue(row, columns) {
                        for (const col of columns) {
                            const val = row && row[col];
                            if (val !== undefined && val !== null && String(val).trim() !== "") {
                                return String(val).toLowerCase().trim();
                            }
                        }
                        return "";
                    }

                    const countStage = (columns, includeWaiting = false) => combined.filter(r => {
                        const s = stageValue(r, columns);
                        if (s === 'no') return false;
                        return s === 'yes' || (includeWaiting && s === 'waiting');
                    }).length;
                    
                    const countExact = (columns, expectedValue) => combined.filter(r => {
                        const s = stageValue(r, columns);
                        if (s === 'no') return false;
                        return s === expectedValue;
                    }).length;

                    // Per requirements: contact attempts come from the contact-related column only.
                    const getContactStatus = (row) => {
                        const keys = ['Contact', 'Contact ', 'Contact/Call', 'Contact / Call', 'Call/Meeting', 'Call'];
                        for (const k of keys) {
                            if (row && Object.prototype.hasOwnProperty.call(row, k)) {
                                const val = String(row[k]).trim().toLowerCase();
                                if (val) return val;
                            }
                        }
                        return '';
                    };

                    const ongoingCount = combined.filter(r => getContactStatus(r) === 'waiting').length;
                    const contactedCount = combined.filter(r => getContactStatus(r) === 'yes').length;
                    const totalContactCount = combined.filter(r => {
                        const s = getContactStatus(r);
                        return s === 'yes' || s === 'waiting';
                    }).length;
                    const repliedCount = countExact(['Reply', 'Replied', 'Moving Forward'], 'yes');
                    const meetingsCount = countExact(['Meeting with Company', 'Meeting'], 'yes');

                    const poolSize = combined.length;

                    const setText = (id, value) => {
                        const el = document.getElementById(id);
                        if (el) el.innerText = value;
                    };

                    setText('k-pool', poolSize);
                    setText('k-contacted', totalContactCount);
                    setText('k-replied', repliedCount);
                    setText('k-meetings', meetingsCount);

                    // 4. Visuals Refresh
                    chartDrillDownStack = [];
                    renderDonut(
                        [vcCount, hnwiCount, foCount],
                        ['VC Funds', 'HNWI / Angels', 'Family Offices'],
                        ['#4f46e5', '#10b981', '#f59e0b']
                    );
                    renderRangeBars();
                    switchTab('vc');

                    // 5. Show Dashboard
                    document.getElementById('loader').style.display = 'none';
                    document.getElementById('dashboard').style.display = 'block';
                }

                function renderDonut(values, labels, colors) {
                    const ctx = document.getElementById('donutChart').getContext('2d');
                    if (charts.donut) charts.donut.destroy();

                    const ringValueLabelPlugin = {
                        id: 'ringValueLabels',
                        afterDatasetsDraw(chart) {
                            const { ctx } = chart;
                            ctx.save();
                            ctx.fillStyle = '#e2e8f0';
                            ctx.font = '600 12px Outfit';
                            ctx.textAlign = 'center';
                            ctx.textBaseline = 'middle';

                            chart.data.datasets.forEach((_dataset, datasetIndex) => {
                                const meta = chart.getDatasetMeta(datasetIndex);
                                if (!meta || !meta.data || !meta.data[0]) return;
                                
                                meta.data.forEach((arc, i) => {
                                    const val = chart.data.datasets[datasetIndex].data[i];
                                    if (!val) return;
                                    const midAngle = (arc.startAngle + arc.endAngle) / 2;
                                    const radius = (arc.innerRadius + arc.outerRadius) / 2;
                                    const x = arc.x + Math.cos(midAngle) * radius;
                                    const y = arc.y + Math.sin(midAngle) * radius;
                                    ctx.fillText(String(val), x, y);
                                });
                            });

                            ctx.restore();
                        }
                    };

                    charts.donut = new Chart(ctx, {
                        type: 'doughnut',
                        data: {
                            labels: labels,
                            datasets: [{
                                data: values,
                                backgroundColor: colors,
                                borderWidth: 0, hoverOffset: 12
                            }]
                        },
                        options: {
                            cutout: '75%', responsive: true, maintainAspectRatio: false,
                            plugins: { 
                                legend: { position: 'bottom', labels: { color: '#94a3b8', font: { family: 'Outfit', size: 12 }, padding: 25 } },
                                tooltip: {
                                    callbacks: {
                                        label(context) {
                                            const label = context.label || '';
                                            const value = context.parsed || 0;
                                            return `${label}: ${value}`;
                                        }
                                    }
                                } 
                            },
                            onClick(event, elements) {
                                if (elements && elements.length > 0) {
                                    const index = elements[0].index;
                                    const label = labels[index];
                                    handleChartClick(label);
                                }
                            }
                        },
                        plugins: [ringValueLabelPlugin]
                    });

                    updateChartControls();
                }

                function handleChartClick(label) {
                    if (chartDrillDownStack.length > 0) return; // Only drill down from level 1 for now

                    const combined = [...rawData.vc, ...rawData.fo];
                    let filteredData = [];

                    if (label === 'VC Funds') {
                        filteredData = rawData.vc;
                    } else if (label === 'HNWI / Angels') {
                        filteredData = rawData.fo.filter(r => {
                            const t = String(r["Type"] || "").toLowerCase();
                            const d = String(r["Description"] || "").toLowerCase();
                            return t.includes('angel') || d.includes('angel') || t.includes('individual') || d.includes('individual') || t.includes('hnwi') || d.includes('hnwi');
                        });
                    } else if (label === 'Family Offices') {
                        filteredData = rawData.fo.filter(r => {
                            const t = String(r["Type"] || "").toLowerCase();
                            const d = String(r["Description"] || "").toLowerCase();
                            return t.includes('family') || d.includes('family') || t.includes('mfo') || t.includes('sfo') || t === 'fo' || d === 'fo';
                        });
                    }

                    if (filteredData.length === 0) return;

                    // Drill down to Stages
                    const stageCounts = {};
                    filteredData.forEach(item => {
                        const stage = getStatusText(item);
                        stageCounts[stage] = (stageCounts[stage] || 0) + 1;
                    });

                    // Sort by value (descending)
                    const sortedStages = Object.entries(stageCounts)
                        .sort((a, b) => b[1] - a[1])
                        .filter(entry => entry[1] > 0);

                    const drillLabels = sortedStages.map(s => s[0]);
                    const drillValues = sortedStages.map(s => s[1]);
                    const drillColors = drillLabels.map((_, i) => COLOR_PALETTE[i % COLOR_PALETTE.length]);

                    const currentValues = charts.donut.data.datasets[0].data;
                    const currentLabels = charts.donut.data.labels;
                    const currentColors = charts.donut.data.datasets[0].backgroundColor;

                    chartDrillDownStack.push({ labels: currentLabels, values: currentValues, colors: currentColors, title: 'Investor Composition' });
                    
                    renderDonut(drillValues, drillLabels, drillColors);
                    
                    const cardTitle = document.querySelector('.card-panel .card-title');
                    if (cardTitle) cardTitle.textContent = `Breakdown: ${label}`;
                }

                function getStatusText(item) {
                    const forward = String(item['Moving Forward'] || item['Replied'] || '').toLowerCase();
                    const meeting = String(item['Meeting with Company'] || item['Meeting'] || '').toLowerCase();
                    const call = String(item['Call/Meeting'] || item['Call'] || item['Contact'] || '').toLowerCase();

                    if (forward === 'yes') return 'Replied / Moving Forward';
                    if (forward === 'waiting' || meeting === 'waiting' || call === 'waiting') return 'Waiting / Ongoing';
                    if (forward === 'no') return 'Passed';
                    if (meeting === 'yes') return 'Contacted / Meeting Done';
                    if (call === 'yes') return 'Contact Started';
                    return 'Target';
                }

                function updateChartControls() {
                    const container = document.getElementById('donut-level-controls');
                    if (!container) return;

                    if (chartDrillDownStack.length > 0) {
                        container.style.display = 'flex';
                        container.innerHTML = `
                            <button class="btn btn-ghost" style="padding: 4px 12px; font-size: 0.75rem;" onclick="goBackInChart()">
                                ← Back to Overview
                            </button>
                        `;
                    } else {
                        container.style.display = 'none';
                        container.innerHTML = '';
                    }
                }

                // Make goBackInChart available globally for the inline onclick handler
                window.goBackInChart = function() {
                    const previous = chartDrillDownStack.pop();
                    if (previous) {
                        renderDonut(previous.values, previous.labels, previous.colors);
                        const cardTitle = document.querySelector('.card-panel .card-title');
                        if (cardTitle) cardTitle.textContent = previous.title;
                    }
                };

                function renderRangeBars() {
                    const combined = [...rawData.vc, ...rawData.fo];
                    const BUCKETS = [
                        { label: '< $100K', min: 0, max: 100000 },
                        { label: '$100K - $500K', min: 100000, max: 500000 },
                        { label: '$500K - $1M', min: 500000, max: 1000000 },
                        { label: '$1M - $5M', min: 1000000, max: 5000000 },
                        { label: '$5M - $10M', min: 5000000, max: 10000000 },
                        { label: '> $10M', min: 10000000, max: Infinity }
                    ];

                    const counts = BUCKETS.map(() => 0);

                    function getMinNumericValue(str) {
                        if (!str) return -1;
                        const match = String(str).match(/([0-9.]+)\s*([kKmM])/);
                        if (!match) return -1;
                        let val = parseFloat(match[1]);
                        const unit = match[2].toUpperCase();
                        if (unit === 'K') val *= 1000;
                        if (unit === 'M') val *= 1000000;
                        return val;
                    }

                    combined.forEach(r => {
                        const raw = r["Size of Investment"] || r["Investment Size"];
                        const val = getMinNumericValue(raw);
                        if (val !== -1) {
                            const idx = BUCKETS.findIndex(b => val >= b.min && val < b.max);
                            if (idx !== -1) counts[idx]++;
                        }
                    });

                    const maxCount = Math.max(...counts, 1);
                    const container = document.getElementById('range-bars');
                    container.innerHTML = '';

                    BUCKETS.forEach((bucket, i) => {
                        const count = counts[i];
                        const perc = (count / maxCount) * 100;
                        container.innerHTML += `
                    <div>
                        <div style="display:flex; justify-content:space-between; font-size:0.85rem; margin-bottom:6px;">
                            <span style="font-weight:500;">${bucket.label}</span>
                            <span style="color:var(--text-dim);">${count}</span>
                        </div>
                        <div style="height:6px; background:rgba(255,255,255,0.05); border-radius:10px; overflow:hidden;">
                            <div style="width:${perc}%; height:100%; background:var(--accent); border-radius:10px;"></div>
                        </div>
                    </div>
                `;
                    });
                }

                function switchTab(type, btn) {
                    activeType = type;
                    if (btn) {
                        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                        btn.classList.add('active');
                    }

                    const header = document.getElementById('table-head');
                    if (type === 'vc') {
                        header.innerHTML = '<tr><th>VC Funds Name</th><th>Investment Size</th><th>Stage</th><th>Contact Email</th></tr>';
                    } else {
                        header.innerHTML = '<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th><th>Stage</th></tr>';
                    }
                    renderTable();
                }

                function toggleFilter(filter) {
                    // Reset others if 'all' is clicked
                    if (filter === 'all') {
                        activeFilters = { all: true, calls: false, meetings: false, forward: false };
                    } else {
                        activeFilters.all = false;
                        activeFilters[filter] = !activeFilters[filter];

                        // If all sub-filters are off, turn 'all' back on
                        if (!activeFilters.calls && !activeFilters.meetings && !activeFilters.forward) {
                            activeFilters.all = true;
                        }
                    }

                    // Update Buttons UI
                    Object.keys(activeFilters).forEach(f => {
                        const btn = document.getElementById(`f-${f}`);
                        if (activeFilters[f]) btn.classList.add('active');
                        else btn.classList.remove('active');
                    });

                    renderTable(document.getElementById('search').value.toLowerCase());
                }

                function getStatusBadge(item) {
                    const forward = String(item['Moving Forward'] || item['Replied'] || '').toLowerCase();
                    const meeting = String(item['Meeting with Company'] || item['Meeting'] || '').toLowerCase();
                    const call = String(item['Call/Meeting'] || item['Call'] || item['Contact'] || '').toLowerCase();

                    if (forward === 'yes') return `<span class="stage-pill badge-green">Replied / Mov. Forward</span>`;
                    if (forward === 'waiting' || meeting === 'waiting' || call === 'waiting') return `<span class="stage-pill badge-amber">Waiting / Ongoing</span>`;
                    if (forward === 'no') return `<span class="stage-pill badge-red">Passed</span>`;
                    if (meeting === 'yes') return `<span class="stage-pill badge-neutral">Contacted / Met</span>`;
                    if (call === 'yes') return `<span class="stage-pill badge-neutral">Contact Started</span>`;
                    return `<span class="stage-pill" style="color:var(--text-dim); opacity:0.5;">Target</span>`;
                }

                function renderTable(keyword = "") {
                    const tbody = document.getElementById('table-body');
                    tbody.innerHTML = '';
                    const data = activeType === 'vc' ? rawData.vc : rawData.fo;

                    data.forEach(item => {
                        const name = item["Investor"] || "";
                        const size = item["Size of Investment"] || item["Investment Size"] || "–";
                        const email = item["Email"] || "–";
                        const note = item["Description"] || "–";

                        // Filter logic
                        if (keyword && !JSON.stringify(item).toLowerCase().includes(keyword)) return;
                        if (!name) return;

                        if (!activeFilters.all) {
                            let match = false;
                            const isYes = (col) => String(item[col]).toLowerCase() === 'yes';
                            const isWait = (col) => String(item[col]).toLowerCase() === 'waiting';

                            if (activeFilters.calls && (isYes('Call/Meeting') || isWait('Call/Meeting'))) match = true;
                            if (activeFilters.meetings && isYes('Meeting with Company')) match = true;
                            if (activeFilters.forward && isYes('Moving Forward')) match = true;
                            if (!match) return;
                        }

                        const statusBadge = getStatusBadge(item);
                        const iType = item["Type"] || "Unknown";
                        const typeColor = typeColors[iType] || '#94a3b8';

                        if (activeType === 'vc') {
                            tbody.innerHTML += `
                        <tr>
                            <td><b>${name}</b></td>
                            <td><span class="size-tag">${size}</span></td>
                            <td>${statusBadge}</td>
                            <td><a href="mailto:${email}" style="color:var(--accent); text-decoration:none;">${email}</a></td>
                        </tr>`;
                        } else {
                            tbody.innerHTML += `
                        <tr>
                            <td><b>${name}</b></td>
                            <td><span style="color:${typeColor}; font-weight:600;">${iType}</span></td>
                            <td><span class="size-tag">${size}</span></td>
                            <td>${statusBadge}</td>
                        </tr>`;
                        }
                    });
                }

                function filterData() {
                    renderTable(document.getElementById('search').value.toLowerCase());
                }

                async function exportStandaloneHtml() {
                    try {
                        const dashboardEl = document.getElementById('dashboard');
                        const shellEl = document.querySelector('.shell');
                        const loaderEl = document.getElementById('loader');
                        if (!dashboardEl || dashboardEl.style.display === 'none') {
                            alert('Dashboard is still loading. Please wait, then try exporting again.');
                            return;
                        }
                        if (!shellEl) {
                            alert('Could not find the dashboard shell to export.');
                            return;
                        }

                        // Clone the visible shell so we keep the full layout/containers
                        const clone = shellEl.cloneNode(true);

                        // Remove navigation for the exported snapshot
                        const nav = clone.querySelector('.sidebar-nav');
                        if (nav) nav.remove();
                        const layout = clone.querySelector('.layout');
                        if (layout) layout.style.gridTemplateColumns = '1fr';

                        // Drop the “Add dashboard” controls from the snapshot
                        const addPanel = clone.querySelector('#add-dashboard-panel');
                        if (addPanel) addPanel.remove();

                        // Snapshot the donut chart into an <img> to preserve visuals without JS
                        const liveDonut = document.getElementById('donutChart');
                        const cloneDonut = clone.querySelector('#donutChart');
                        if (liveDonut && cloneDonut) {
                            const dataUrl = liveDonut.toDataURL('image/png');
                            const img = document.createElement('img');
                            img.src = dataUrl;
                            img.alt = 'Investor composition chart';
                            img.style.width = '100%';
                            img.style.height = '100%';
                            img.style.objectFit = 'contain';
                            cloneDonut.parentNode.replaceChild(img, cloneDonut);
                        }

                        // Bundle the CSS used on this page so the export is fully standalone
                        const cssPaths = Array.from(
                            document.querySelectorAll('link[rel="stylesheet"][href]')
                        )
                            .map((link) => link.getAttribute('href'))
                            .filter((href) => href && !/^https?:/i.test(href));
                        const cssText = await Promise.all(
                            cssPaths.map((href) =>
                                fetch(new URL(href, window.location.href))
                                    .then(r => r.text())
                                    .catch(() => `/* missing: ${href} */`)
                            )
                        );

                        const snapshotState = {
                            rawData,
                            typeColors,
                            activeType,
                            activeFilters,
                            searchTerm: (document.getElementById('search')?.value || '').toLowerCase()
                        };

                        const snapshotHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${document.title} – Snapshot</title>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
  <style>${cssText.join('\\n\\n')}</style>
</head>
<body>
${clone.outerHTML}
<script>
(() => {
  const state = ${JSON.stringify(snapshotState)};
  const rawData = state.rawData || { vc: [], fo: [] };
  const typeColors = state.typeColors || {};
  let activeType = state.activeType || 'vc';
  let activeFilters = state.activeFilters || { all: true, calls: false, meetings: false, forward: false };
  const searchTerm = state.searchTerm || '';

  const searchInput = document.getElementById('search');
  if (searchInput && searchTerm) searchInput.value = searchTerm;

  function getStatusBadge(item) {
    const forward = String(item['Moving Forward'] || item['Replied'] || '').toLowerCase();
    const meeting = String(item['Meeting with Company'] || item['Meeting'] || '').toLowerCase();
    const call = String(item['Call/Meeting'] || item['Call'] || item['Contact'] || '').toLowerCase();

    if (forward === 'yes') return '<span class="stage-pill badge-green">Replied / Mov. Forward</span>';
    if (forward === 'waiting' || meeting === 'waiting' || call === 'waiting') return '<span class="stage-pill badge-amber">Waiting / Ongoing</span>';
    if (forward === 'no') return '<span class="stage-pill badge-red">Passed</span>';
    if (meeting === 'yes') return '<span class="stage-pill badge-neutral">Contacted / Met</span>';
    if (call === 'yes') return '<span class="stage-pill badge-neutral">Contact Started</span>';
    return '<span class="stage-pill" style="color:var(--text-dim); opacity:0.5;">Target</span>';
  }

  function renderTable(keyword = '') {
    const tbody = document.getElementById('table-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const data = activeType === 'vc' ? (rawData.vc || []) : (rawData.fo || []);

    data.forEach(item => {
      const name = item['Investor'] || '';
      const size = item['Size of Investment'] || item['Investment Size'] || '–';
      const email = item['Email'] || '–';

      if (keyword && !JSON.stringify(item).toLowerCase().includes(keyword)) return;
      if (!name) return;

      if (!activeFilters.all) {
        let match = false;
        const isYes = col => String(item[col]).toLowerCase() === 'yes';
        const isWait = col => String(item[col]).toLowerCase() === 'waiting';
        if (activeFilters.calls && (isYes('Call/Meeting') || isWait('Call/Meeting'))) match = true;
        if (activeFilters.meetings && isYes('Meeting with Company')) match = true;
        if (activeFilters.forward && isYes('Moving Forward')) match = true;
        if (!match) return;
      }

      const statusBadge = getStatusBadge(item);
      const iType = item['Type'] || 'Unknown';
      const typeColor = typeColors[iType] || '#94a3b8';

      if (activeType === 'vc') {
        tbody.innerHTML += \`
          <tr>
            <td><b>\${name}</b></td>
            <td><span class="size-tag">\${size}</span></td>
            <td>\${statusBadge}</td>
            <td><a href="mailto:\${email}" style="color:var(--accent); text-decoration:none;">\${email}</a></td>
          </tr>\`;
      } else {
        tbody.innerHTML += \`
          <tr>
            <td><b>\${name}</b></td>
            <td><span style="color:\${typeColor}; font-weight:600;">\${iType}</span></td>
            <td><span class="size-tag">\${size}</span></td>
            <td>\${statusBadge}</td>
          </tr>\`;
      }
    });
  }

  function switchTab(type, btn) {
    activeType = type;
    if (btn) {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    }
    const header = document.getElementById('table-head');
    if (header) {
      if (type === 'vc') {
        header.innerHTML = '<tr><th>VC Funds Name</th><th>Investment Size</th><th>Stage</th><th>Contact Email</th></tr>';
      } else {
        header.innerHTML = '<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th><th>Stage</th></tr>';
      }
    }
    renderTable(searchInput ? searchInput.value.toLowerCase() : '');
  }

  function toggleFilter(filter) {
    if (filter === 'all') {
      activeFilters = { all: true, calls: false, meetings: false, forward: false };
    } else {
      activeFilters.all = false;
      activeFilters[filter] = !activeFilters[filter];
      if (!activeFilters.calls && !activeFilters.meetings && !activeFilters.forward) {
        activeFilters.all = true;
      }
    }
    Object.keys(activeFilters).forEach(f => {
      const btn = document.getElementById(\`f-\${f}\`);
      if (btn) {
        if (activeFilters[f]) btn.classList.add('active');
        else btn.classList.remove('active');
      }
    });
    renderTable(searchInput ? searchInput.value.toLowerCase() : '');
  }

  function filterData() {
    renderTable(searchInput ? searchInput.value.toLowerCase() : '');
  }

  document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('[data-tab-type]').forEach(btn => {
      btn.addEventListener('click', () => switchTab(btn.getAttribute('data-tab-type'), btn));
    });
    document.querySelectorAll('[data-filter]').forEach(btn => {
      btn.addEventListener('click', () => toggleFilter(btn.getAttribute('data-filter')));
    });
    if (searchInput) searchInput.addEventListener('input', filterData);

    // Restore active tab/filter visual state
    const activeTabBtn = document.querySelector(\`[data-tab-type=\"\${activeType}\"]\`);
    if (activeTabBtn) activeTabBtn.classList.add('active');
    Object.keys(activeFilters).forEach(f => {
      if (activeFilters[f]) {
        const btn = document.getElementById(\`f-\${f}\`);
        if (btn) btn.classList.add('active');
      }
    });

    // Seed header and table content
    switchTab(activeType);
    if (searchTerm) renderTable(searchTerm);
  });
})();
</script>
</body>
</html>`;

                        const blob = new Blob([snapshotHtml], { type: 'text/html' });
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        const dashId = (currentDashboard && currentDashboard.id) || 'snapshot';
                        a.download = `investor-dashboard-${dashId}.html`;
                        a.href = url;
                        a.click();
                        setTimeout(() => URL.revokeObjectURL(url), 1000);
                    } catch (err) {
                        console.error('Export failed', err);
                        alert('Could not export a snapshot. Please try again.');
                    }
                }
