                // --- CONFIGURATION ---
                // Sensitive links are loaded from config.js
                const DEFAULT_DASHBOARDS_CONFIG = window.DASHBOARD_CONFIG || {
                    dashboards: [],
                    settings: {
                        defaultDashboard: "",
                        allowLocalUpload: true,
                        title: "Investor Dashboard"
                    }
                };

                let dashboardsConfig = DEFAULT_DASHBOARDS_CONFIG;
                let currentDashboard = null;

                // Proxy fallback chain loaded from external config if available
                const PROXIES = window.DASHBOARD_PROXIES || [];

                // --- STATE ---
                let rawData = { vc: [], fo: [] };
                let activeType = 'vc';
                let charts = {};
                let activeFilters = { all: true, calls: false, meetings: false, forward: false };
                let typeColors = {};
                let adoramDonutState = { level2: true, level3: true, level4: true, level5: true };
                let adoramDonutMetrics = null;
                const COLOR_PALETTE = ['#818cf8', '#34d399', '#fbbf24', '#a78bfa', '#f472b6', '#22d3ee', '#fb7185'];

                function normalizeRowKeys(row) {
                    const normalized = {};
                    Object.keys(row || {}).forEach((key) => {
                        const cleanKey = String(key || "").replace(/\s+/g, " ").trim();
                        normalized[cleanKey] = row[key];
                    });
                    return normalized;
                }

                function stageValue(row, columns) {
                    for (const col of columns) {
                        const val = row && row[col];
                        if (val !== undefined && val !== null && String(val).trim() !== "") {
                            return String(val).toLowerCase().trim();
                        }
                    }
                    return "";
                }

                function isAdoramDashboard() {
                    return !!(currentDashboard && String(currentDashboard.id).toLowerCase() === 'adoram');
                }

                function setText(id, text) {
                    const el = document.getElementById(id);
                    if (el) el.textContent = text;
                }

                function updateDashboardIdentity() {
                    const idEl = document.getElementById('dashboard-id-value');
                    if (idEl) idEl.textContent = (currentDashboard && currentDashboard.id) || '—';
                }

                function setFilterButtonLabel(id, label) {
                    const btn = document.getElementById(id);
                    if (!btn) return;
                    const dot = btn.querySelector('.toggle-dot');
                    btn.textContent = '';
                    if (dot) btn.appendChild(dot);
                    btn.appendChild(document.createTextNode(` ${label}`));
                }

                function updateKpiLabels() {
                    const totalCard = document.getElementById('kpi-total-card');
                    const topKpiGrid = document.getElementById('top-kpi-grid');

                    if (isAdoramDashboard()) {
                        setText('k-label-vc', 'Total Contact');
                        setText('k-label-hnwi', 'Contacted');
                        setText('k-label-fo', 'Interested');
                        setText('k-label-total', 'Total Investors');
                        setText('kpi-funnel-title', 'Contact Funnel');
                        setText('k-label-calls', 'Ongoing');
                        setText('k-label-meetings', 'Contacted');
                        setText('k-label-forward', 'Replied');
                        if (totalCard) totalCard.style.display = 'none';
                        if (topKpiGrid) topKpiGrid.style.gridTemplateColumns = 'repeat(3, 1fr)';
                        return;
                    }

                    setText('k-label-vc', 'VC Funds');
                    setText('k-label-hnwi', 'HNWI / Angels');
                    setText('k-label-fo', 'Family Offices');
                    setText('k-label-total', 'Total Pipeline');
                    setText('kpi-funnel-title', 'Engagement Funnel');
                    setText('k-label-calls', 'Contact Started');
                    setText('k-label-meetings', 'Meeting w/ Company');
                    setText('k-label-forward', 'Moving Forward');
                    if (totalCard) totalCard.style.display = '';
                    if (topKpiGrid) topKpiGrid.style.gridTemplateColumns = '';
                }

                function updateFilterLabels() {
                    if (isAdoramDashboard()) {
                        setFilterButtonLabel('f-all', 'All Investors');
                        setFilterButtonLabel('f-calls', 'Ongoing');
                        setFilterButtonLabel('f-meetings', 'Contacted');
                        setFilterButtonLabel('f-forward', 'Replied');
                        return;
                    }

                    setFilterButtonLabel('f-all', 'All Investors');
                    setFilterButtonLabel('f-calls', 'Contact Started');
                    setFilterButtonLabel('f-meetings', 'Meetings');
                    setFilterButtonLabel('f-forward', 'Moving Forward');
                }

                function buildAdoramDonutControls() {
                    const container = document.getElementById('donut-level-controls');
                    if (!container) return;
                    container.style.display = 'flex';
                    container.innerHTML = '';
                    container.style.display = 'none';
                }

                function renderDonutFromState() {
                    if (!isAdoramDashboard() || !adoramDonutMetrics) return;
                    renderDonutHierarchy(adoramDonutMetrics);
                }

                function toggleAdoramDonutLevel(levelKey) {
                    if (!isAdoramDashboard()) return;
                    const childByLevel = {
                        level1: 'level2',
                        level2: 'level3',
                        level3: 'level4',
                        level4: 'level5',
                    };
                    const targetLevel = childByLevel[levelKey] || (levelKey === 'level5' ? 'level5' : '');
                    if (!targetLevel) return;

                    const nextValue = !adoramDonutState[targetLevel];
                    adoramDonutState[targetLevel] = nextValue;

                    if (!adoramDonutState.level2) {
                        adoramDonutState.level3 = false;
                        adoramDonutState.level4 = false;
                        adoramDonutState.level5 = false;
                    }
                    if (!adoramDonutState.level3) {
                        adoramDonutState.level4 = false;
                        adoramDonutState.level5 = false;
                    }
                    if (!adoramDonutState.level4) {
                        adoramDonutState.level5 = false;
                    }

                    if (targetLevel === 'level3' && nextValue) {
                        adoramDonutState.level2 = true;
                    }
                    if (targetLevel === 'level4' && nextValue) {
                        adoramDonutState.level2 = true;
                        adoramDonutState.level3 = true;
                    }
                    if (targetLevel === 'level5' && nextValue) {
                        adoramDonutState.level2 = true;
                        adoramDonutState.level3 = true;
                        adoramDonutState.level4 = true;
                    }
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
                        updateDashboardIdentity();
                        updateKpiLabels();
                        updateFilterLabels();
                        buildAdoramDonutControls();

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
                    updateDashboardIdentity();
                    updateKpiLabels();
                    updateFilterLabels();
                    buildAdoramDonutControls();

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

                async function startSync(proxyIndex, excelUrl) {
                    if (proxyIndex >= PROXIES.length) {
                        console.error('[Dashboard] All proxies failed for URL:', excelUrl);
                        showFailure();
                        return;
                    }

                    const proxyFunc = PROXIES[proxyIndex];
                    const loadMsg = document.getElementById('loader-msg');
                    loadMsg.innerText = proxyIndex === 0 ? "Connecting to SharePoint..." : "Sync failed, trying backup tunnel...";

                    try {
                        const fetchUrl = proxyFunc(excelUrl);
                        console.log('[Dashboard] Using proxy', proxyIndex, '->', fetchUrl);
                        const response = await fetch(fetchUrl);

                        if (!response.ok) throw new Error("Proxy Error");

                        let buffer;
                        if (fetchUrl.includes('allorigins')) {
                            // AllOrigins specialty: JSON wrapper with Base64 content
                            const json = await response.json();
                            if (!json.contents) throw new Error("No contents in JSON");

                            // Decode base64 to binary
                            const b64 = json.contents.split(',')[1] || json.contents;
                            const binaryString = atob(b64);
                            buffer = new ArrayBuffer(binaryString.length);
                            const view = new Uint8Array(buffer);
                            for (let i = 0; i < binaryString.length; i++) view[i] = binaryString.charCodeAt(i);
                        } else {
                            // Standard raw proxy
                            buffer = await response.arrayBuffer();
                        }

                        const wb = XLSX.read(buffer, { type: 'array' });
                        processWorkbook(wb, currentDashboard && currentDashboard.sheets);
                        document.getElementById('sync-type').innerText = "Live Cloud Sync Active";

                    } catch (err) {
                        console.warn(`[Dashboard] Proxy ${proxyIndex} failed for URL: ${excelUrl}`, err);
                        startSync(proxyIndex + 1, excelUrl); // Try next proxy
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
                    rawData.vc = XLSX.utils
                        .sheet_to_json(fSheet, { range: headerRowIndex, defval: "" })
                        .map(normalizeRowKeys)
                        .filter(r => r["Investor"]);
                    rawData.fo = XLSX.utils
                        .sheet_to_json(oSheet, { range: headerRowIndex, defval: "" })
                        .map(normalizeRowKeys)
                        .filter(r => r["Investor"]);
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

                    const totalInvestorsOverride = isAdoramDashboard() ? 374 : combined.length;
                    document.getElementById('k-vc').innerText = vcCount;
                    document.getElementById('k-hnwi').innerText = hnwiCount;
                    document.getElementById('k-fo').innerText = foCount;
                    document.getElementById('k-total').innerText = totalInvestorsOverride;

                    // 3. Stage Funnel KPIs (Yes/Waiting logic)
                    const countStage = (columns, includeWaiting = false) => combined.filter(r => {
                        const s = stageValue(r, columns);
                        return s === 'yes' || (includeWaiting && s === 'waiting');
                    }).length;
                    const countExact = (columns, expectedValue) => combined.filter(r => stageValue(r, columns) === expectedValue).length;
                    const ongoingCount = countExact(['Contact'], 'waiting');
                    const contactedCount = countExact(['Contact'], 'yes');
                    const totalContactCount = ongoingCount + contactedCount;
                    const repliedCount = countExact(['Replied'], 'yes');
                    let interestedCount = 3;
                    interestedCount = Math.min(interestedCount, repliedCount);
                    const callCount = countStage(['Call/Meeting', 'Call', 'Contact'], true);
                    const meetingCount = countStage(['Meeting with Company', 'Meeting']);
                    const forwardCount = countStage(['Moving Forward']);

                    if (isAdoramDashboard()) {
                        document.getElementById('k-vc').innerText = totalContactCount;
                        document.getElementById('k-hnwi').innerText = contactedCount;
                        document.getElementById('k-fo').innerText = interestedCount;
                        document.getElementById('k-total').innerText = totalInvestorsOverride;
                        document.getElementById('k-calls').innerText = ongoingCount;
                        document.getElementById('k-meetings').innerText = contactedCount;
                        document.getElementById('k-forward').innerText = repliedCount;
                    } else {
                        document.getElementById('k-calls').innerText = callCount;
                        document.getElementById('k-meetings').innerText = meetingCount;
                        document.getElementById('k-forward').innerText = forwardCount;
                    }

                    // 4. Visuals Refresh
                    if (isAdoramDashboard()) {
                        adoramDonutMetrics = {
                            totalInvestors: totalInvestorsOverride,
                            contactScope: totalContactCount,
                            contacted: contactedCount,
                            replied: repliedCount,
                            interested: interestedCount
                        };
                        buildAdoramDonutControls();
                        renderDonutFromState();
                    } else {
                        adoramDonutMetrics = null;
                        renderDonut(
                            [vcCount, hnwiCount, foCount],
                            ['VC Funds', 'HNWI / Angels', 'Family Offices'],
                            ['#4f46e5', '#10b981', '#f59e0b']
                        );
                    }
                    renderRangeBars();
                    switchTab('vc');

                    // 5. Show Dashboard
                    document.getElementById('loader').style.display = 'none';
                    document.getElementById('dashboard').style.display = 'block';
                }

                function renderDonut(values, labels, colors) {
                    const ctx = document.getElementById('donutChart').getContext('2d');
                    if (charts.donut) charts.donut.destroy();
                    const safeValues = values.map((v) => Math.max(Number(v) || 0, 0));
                    const total = Math.max(safeValues.reduce((acc, v) => acc + v, 0), 1);

                    const datasets = safeValues.map((value, idx) => ({
                        label: labels[idx],
                        value,
                        data: [value, Math.max(total - value, 0)],
                        backgroundColor: [colors[idx], 'rgba(148,163,184,0.12)'],
                        borderWidth: 0,
                        hoverOffset: 8
                    }));

                    charts.donut = new Chart(ctx, {
                        type: 'doughnut',
                        data: {
                            labels: ['Value', 'Remaining'],
                            datasets
                        },
                        options: {
                            cutout: '35%',
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: {
                                    position: 'bottom',
                                    labels: {
                                        color: '#94a3b8',
                                        font: { family: 'Outfit', size: 12 },
                                        padding: 20,
                                        generateLabels(chart) {
                                            const items = chart.data.datasets || [];
                                            return items.map((ds, i) => ({
                                                text: `${ds.label}: ${ds.value}`,
                                                fillStyle: ds.backgroundColor[0],
                                                strokeStyle: ds.backgroundColor[0],
                                                lineWidth: 0,
                                                hidden: false,
                                                index: i
                                            }));
                                        }
                                    }
                                },
                                tooltip: {
                                    callbacks: {
                                        label(context) {
                                            const ds = context.dataset;
                                            if (context.dataIndex === 0) {
                                                return `${ds.label}: ${ds.value}`;
                                            }
                                            return `Remaining: ${Math.max(total - ds.value, 0)}`;
                                        }
                                    }
                                }
                            }
                        }
                    });
                }

                function renderDonutHierarchy(metrics) {
                    const ctx = document.getElementById('donutChart').getContext('2d');
                    if (charts.donut) charts.donut.destroy();

                    const totalInvestors = Math.max(Number(metrics.totalInvestors) || 0, 0);
                    const contactScope = Math.min(Math.max(Number(metrics.contactScope) || 0, 0), totalInvestors);
                    const contacted = Math.min(Math.max(Number(metrics.contacted) || 0, 0), contactScope);
                    const replied = Math.min(Math.max(Number(metrics.replied) || 0, 0), contacted);
                    const interested = Math.min(Math.max(Number(metrics.interested) || 0, 0), replied);

                    const levelsInnerToOuter = [
                        { key: 'level2', depth: 1, label: 'Contacted + Ongoing', value: contactScope, parentValue: totalInvestors, color: '#4f46e5', show: adoramDonutState.level2 },
                        { key: 'level3', depth: 2, label: 'Contacted', value: contacted, parentValue: contactScope, color: '#f59e0b', show: adoramDonutState.level2 && adoramDonutState.level3 },
                        { key: 'level4', depth: 3, label: 'Replied', value: replied, parentValue: contacted, color: '#10b981', show: adoramDonutState.level2 && adoramDonutState.level3 && adoramDonutState.level4 },
                        { key: 'level5', depth: 4, label: 'Interested', value: interested, parentValue: replied, color: '#22c55e', show: adoramDonutState.level2 && adoramDonutState.level3 && adoramDonutState.level4 && adoramDonutState.level5 }
                    ].filter((level) => level.show);

                    // Chart.js draws doughnut datasets outer-to-inner by dataset order.
                    // Reverse to keep the broadest visible scope on the innermost ring.
                    const datasets = [...levelsInnerToOuter].reverse().map((level) => {
                        const parentValue = Math.max(level.parentValue, 0);
                        const value = Math.min(level.value, parentValue);
                        const visualBase = Math.max(totalInvestors, 0);
                        const remainder = Math.max(visualBase - value, 0);
                        return {
                            key: level.key,
                            depth: level.depth,
                            label: level.label,
                            value,
                            parentValue,
                            totalValue: totalInvestors,
                            data: [value, remainder],
                            backgroundColor: [level.color, 'rgba(148,163,184,0.12)'],
                            borderWidth: 0,
                            hoverOffset: 8
                        };
                    });

                    const ringValueLabelPlugin = {
                        id: 'ringValueLabels',
                        afterDatasetsDraw(chart) {
                            const { ctx } = chart;
                            ctx.save();
                            ctx.fillStyle = '#e2e8f0';
                            ctx.font = '600 11px Outfit';
                            ctx.textAlign = 'center';
                            ctx.textBaseline = 'middle';

                            chart.data.datasets.forEach((_dataset, datasetIndex) => {
                                const meta = chart.getDatasetMeta(datasetIndex);
                                if (!meta || !meta.data || !meta.data[0]) return;
                                const arc = meta.data[0];
                                const midAngle = (arc.startAngle + arc.endAngle) / 2;
                                const radius = (arc.innerRadius + arc.outerRadius) / 2;
                                const x = arc.x + Math.cos(midAngle) * radius;
                                const y = arc.y + Math.sin(midAngle) * radius;
                                const ds = chart.data.datasets[datasetIndex];
                                ctx.fillText(String(ds.value), x, y);
                            });

                            ctx.restore();
                        }
                    };

                    charts.donut = new Chart(ctx, {
                        type: 'doughnut',
                        data: {
                            labels: ['Value', 'Remaining'],
                            datasets
                        },
                        options: {
                            cutout: '22%',
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: {
                                    position: 'bottom',
                                    labels: {
                                        color: '#94a3b8',
                                        font: { family: 'Outfit', size: 12 },
                                        padding: 18,
                                        generateLabels(chart) {
                                            return (chart.data.datasets || [])
                                                .slice()
                                                .sort((a, b) => a.depth - b.depth)
                                                .map((ds) => ({
                                                    text: ds.depth === 1
                                                        ? `${ds.label}: ${ds.value}`
                                                        : `${ds.label}: ${ds.value}/${ds.parentValue} (of ${ds.totalValue})`,
                                                    fillStyle: ds.backgroundColor[0],
                                                    strokeStyle: ds.backgroundColor[0],
                                                    lineWidth: 0,
                                                    hidden: false,
                                                    index: (chart.data.datasets || []).indexOf(ds)
                                                }));
                                        }
                                    }
                                },
                                tooltip: {
                                    callbacks: {
                                        label(context) {
                                            const ds = context.dataset;
                                            if (context.dataIndex === 0) {
                                                return ds.depth === 1
                                                    ? `${ds.label}: ${ds.value}`
                                                    : `${ds.label}: ${ds.value}/${ds.parentValue} (of ${ds.totalValue})`;
                                            }
                                            return `Remaining to total: ${Math.max(ds.totalValue - ds.value, 0)}`;
                                        }
                                    }
                                }
                            },
                            onClick(event, elements, chart) {
                                if (!isAdoramDashboard()) return;
                                const hit = elements && elements.length ? elements[0] : null;
                                if (!hit) return;
                                if (hit.index !== 0) return;
                                const ds = chart.data.datasets[hit.datasetIndex];
                                if (!ds || !ds.key) return;
                                toggleAdoramDonutLevel(ds.key);
                                renderDonutFromState();
                            }
                        },
                        plugins: [ringValueLabelPlugin]
                    });
                }

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
                    const isAdoram = currentDashboard && currentDashboard.id === 'adoram';
                    if (type === 'vc') {
                        if (isAdoram) {
                            header.innerHTML = '<tr><th>VC Funds Name</th><th>Investment Size</th><th>Contact</th><th>Replied</th><th>Stage</th><th>Contact Email</th></tr>';
                        } else {
                            header.innerHTML = '<tr><th>VC Funds Name</th><th>Investment Size</th><th>Stage</th><th>Contact Email</th></tr>';
                        }
                    } else {
                        if (isAdoram) {
                            header.innerHTML = '<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th><th>Contact</th><th>Replied</th><th>Stage</th></tr>';
                        } else {
                            header.innerHTML = '<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th><th>Stage</th></tr>';
                        }
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
                    const forward = stageValue(item, ['Moving Forward']);
                    const meeting = stageValue(item, ['Meeting with Company', 'Meeting']);
                    const call = stageValue(item, ['Call/Meeting', 'Call', 'Contact']);

                    if (forward === 'yes') return `<span class="stage-pill badge-green">Moving Forward</span>`;
                    if (forward === 'waiting' || meeting === 'waiting' || call === 'waiting') return `<span class="stage-pill badge-amber">Waiting</span>`;
                    if (forward === 'no') return `<span class="stage-pill badge-red">Passed</span>`;
                    if (meeting === 'yes') return `<span class="stage-pill badge-neutral">Meeting Done</span>`;
                    if (call === 'yes') return `<span class="stage-pill badge-neutral">Contact Started</span>`;
                    return `<span class="stage-pill" style="color:var(--text-dim); opacity:0.5;">Target</span>`;
                }

                function formatBooleanPill(value) {
                    const clean = String(value || '').toLowerCase().trim();
                    if (clean === 'yes') return `<span class="stage-pill badge-green">Yes</span>`;
                    if (clean === 'waiting') return `<span class="stage-pill badge-amber">Waiting</span>`;
                    if (clean === 'no') return `<span class="stage-pill badge-red">No</span>`;
                    return `<span class="stage-pill" style="color:var(--text-dim); opacity:0.5;">-</span>`;
                }

                function renderTable(keyword = "") {
                    const tbody = document.getElementById('table-body');
                    tbody.innerHTML = '';
                    const data = activeType === 'vc' ? rawData.vc : rawData.fo;
                    const isAdoram = currentDashboard && currentDashboard.id === 'adoram';

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
                            const isYes = (cols) => stageValue(item, cols) === 'yes';
                            const isWait = (cols) => stageValue(item, cols) === 'waiting';

                            if (isAdoramDashboard()) {
                                if (activeFilters.calls && isWait(['Contact'])) match = true;
                                if (activeFilters.meetings && isYes(['Contact'])) match = true;
                                if (activeFilters.forward && isYes(['Replied'])) match = true;
                            } else {
                                if (activeFilters.calls && (isYes(['Call/Meeting', 'Call', 'Contact']) || isWait(['Call/Meeting', 'Call', 'Contact']))) match = true;
                                if (activeFilters.meetings && isYes(['Meeting with Company', 'Meeting'])) match = true;
                                if (activeFilters.forward && isYes(['Moving Forward'])) match = true;
                            }
                            if (!match) return;
                        }

                        const statusBadge = getStatusBadge(item);
                        const contactBadge = formatBooleanPill(item['Contact']);
                        const repliedBadge = formatBooleanPill(item['Replied']);
                        const iType = item["Type"] || "Unknown";
                        const typeColor = typeColors[iType] || '#94a3b8';

                        if (activeType === 'vc') {
                            if (isAdoram) {
                                tbody.innerHTML += `
                        <tr>
                            <td><b>${name}</b></td>
                            <td><span class="size-tag">${size}</span></td>
                            <td>${contactBadge}</td>
                            <td>${repliedBadge}</td>
                            <td>${statusBadge}</td>
                            <td><a href="mailto:${email}" style="color:var(--accent); text-decoration:none;">${email}</a></td>
                        </tr>`;
                            } else {
                                tbody.innerHTML += `
                        <tr>
                            <td><b>${name}</b></td>
                            <td><span class="size-tag">${size}</span></td>
                            <td>${statusBadge}</td>
                            <td><a href="mailto:${email}" style="color:var(--accent); text-decoration:none;">${email}</a></td>
                        </tr>`;
                            }
                        } else {
                            if (isAdoram) {
                                tbody.innerHTML += `
                        <tr>
                            <td><b>${name}</b></td>
                            <td><span style="color:${typeColor}; font-weight:600;">${iType}</span></td>
                            <td><span class="size-tag">${size}</span></td>
                            <td>${contactBadge}</td>
                            <td>${repliedBadge}</td>
                            <td>${statusBadge}</td>
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
                        }
                    });
                }

                function filterData() {
                    renderTable(document.getElementById('search').value.toLowerCase());
                }
