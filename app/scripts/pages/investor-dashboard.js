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
                let currentDashboardMode = 'home';
                let dashboardFormMode = 'edit';
                let refreshButtonBusy = false;

                // Proxy fallback chain loaded from external config if available
                const PROXIES = window.DASHBOARD_PROXIES || [];

                // --- STATE ---
                let rawData = { vc: [], fo: [] };
                let activeType = 'vc';
                let activeFilters = { all: true, calls: false, meetings: false, forward: false };
                let compositionSelection = { groupKey: '', stageLabel: '' };
                let typeColors = {};
                let pendingTableRenderFrame = 0;
                let pendingSearchFilterTimer = 0;
                const COLOR_PALETTE = ['#818cf8', '#34d399', '#fbbf24', '#a78bfa', '#f472b6', '#22d3ee', '#fb7185'];
                const DASHBOARD_SEARCH_FIELDS = [
                    'Investor',
                    'Investor Name',
                    'Name',
                    'Email',
                    'Type',
                    'Type of Client',
                    'Description',
                    'Size of Investment',
                    'Investment Size',
                    'Replied',
                    'Moving Forward',
                    'Meeting with Company',
                    'Meeting',
                    'Call/Meeting',
                    'Call',
                    'Contact',
                ];
                const COMPOSITION_STAGE_ORDER = [
                    'Target',
                    'Contact Started',
                    'Contacted / Meeting Done',
                    'Waiting / Ongoing',
                    'Replied / Moving Forward',
                ];
                const COMPOSITION_GROUPS = [
                    { key: 'vc', label: 'VC Funds', shortLabel: 'VC', color: '#4f46e5', childColors: ['#312e81', '#3730a3', '#4338ca', '#6366f1', '#818cf8', '#a5b4fc'] },
                    { key: 'hnwi', label: 'HNWI / Angels', shortLabel: 'HNWI', color: '#10b981', childColors: ['#065f46', '#047857', '#059669', '#10b981', '#34d399', '#6ee7b7'] },
                    { key: 'fo', label: 'Family Offices', shortLabel: 'FO', color: '#f59e0b', childColors: ['#92400e', '#b45309', '#d97706', '#f59e0b', '#fbbf24', '#fcd34d'] },
                ];
                const NAVRO_KEY_TYPE_LABELS = {
                    CVC: 'Corporate Venture Capital',
                    VC: 'Venture Capital',
                    PB: 'Private Bank',
                    SFO: 'Single Family Office',
                    MFO: 'Multi Family Office',
                    AS: 'Angel Syndicate',
                    PF: 'Pension Funds',
                    AM: 'Asset Manager',
                    INS: 'Insurance Company',
                    HNW: 'High Net Worth Individual',
                };
                const NAVRO_CLIENT_TYPE_TO_KEY = Object.entries(NAVRO_KEY_TYPE_LABELS).reduce((acc, [key, label]) => {
                    acc[normalizeDashboardText(label)] = key;
                    return acc;
                }, {});
                const NAVRO_VC_KEYS = new Set(['CVC', 'VC', 'PB', 'PF', 'INS', 'AM']);
                const NAVRO_HNWI_KEYS = new Set(['AS', 'HNW']);
                const CONTACT_FIELD_KEYS = ['Contact\u00a0', 'Contact', 'Contact ', 'Contact/Call', 'Contact / Call'];
                const CALL_FIELD_KEYS = ['Call/Meeting', 'Call'];
                const MEETING_FIELD_KEYS = ['Meeting\u00a0', 'Meeting with Company', 'Meeting'];
                const FORWARD_FIELD_KEYS = ['Reply', 'Replied', 'Moving Forward'];

                function normalizeDashboardText(value) {
                    return String(value || '').trim().toLowerCase();
                }

                function escapeHtml(value) {
                    return String(value == null ? '' : value)
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;')
                        .replace(/'/g, '&#39;');
                }

                function isDashboardLowPowerDevice() {
                    const nav = window.navigator || {};
                    const connection = nav.connection || nav.mozConnection || nav.webkitConnection || null;
                    const prefersReducedMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
                    const hardwareConcurrency = Number(nav.hardwareConcurrency || 0);
                    const deviceMemory = Number(nav.deviceMemory || 0);
                    return Boolean(
                        prefersReducedMotion ||
                        (connection && connection.saveData) ||
                        (hardwareConcurrency && hardwareConcurrency <= 4) ||
                        (deviceMemory && deviceMemory <= 4)
                    );
                }

                function applyDashboardPerformanceMode() {
                    const body = document.body;
                    if (!body) return;
                    if (isDashboardLowPowerDevice()) {
                        body.setAttribute('data-dashboard-performance', 'lite');
                        return;
                    }
                    body.removeAttribute('data-dashboard-performance');
                }

                function normalizeDealLifecycleStatus(value) {
                    const normalized = normalizeDashboardText(value);
                    if (normalized === 'finished') return 'finished';
                    if (normalized === 'closed') return 'closed';
                    return 'active';
                }

                function isActiveDeal(deal) {
                    return normalizeDealLifecycleStatus(deal && (deal.lifecycleStatus || deal.dealStatus)) === 'active';
                }

                function isNavroDashboard() {
                    return normalizeDashboardId(currentDashboard && currentDashboard.id) === 'navro';
                }

                function getDashboardTypeKey(row) {
                    return String(
                        row && (
                            row["KEY"] ||
                            row["Key"] ||
                            row["key"] ||
                            ''
                        )
                    ).trim().toUpperCase();
                }

                function getNavroClientType(row) {
                    return String(
                        row && (
                            row["Type of Client"] ||
                            row["Type of client"] ||
                            row["Type Of Client"] ||
                            row["type of client"] ||
                            row["Type"] ||
                            ''
                        )
                    ).trim();
                }

                function normalizeNavroClientTypeCode(value) {
                    const normalized = String(value || '').trim().toUpperCase();
                    if (NAVRO_KEY_TYPE_LABELS[normalized]) {
                        return normalized;
                    }

                    const normalizedLabel = normalizeDashboardText(value);
                    return NAVRO_CLIENT_TYPE_TO_KEY[normalizedLabel] || '';
                }

                function getNavroClientTypeCode(row) {
                    const clientTypeCode = normalizeNavroClientTypeCode(getNavroClientType(row));
                    if (clientTypeCode) return clientTypeCode;
                    return normalizeNavroClientTypeCode(getDashboardTypeKey(row));
                }

                function getInvestorTypeLabel(row) {
                    if (isNavroDashboard()) {
                        const navroCode = getNavroClientTypeCode(row);
                        if (navroCode && NAVRO_KEY_TYPE_LABELS[navroCode]) {
                            return NAVRO_KEY_TYPE_LABELS[navroCode];
                        }
                        const navroClientType = getNavroClientType(row);
                        if (navroClientType) {
                            return navroClientType;
                        }
                    }
                    return String(row && row["Type"] || 'Unknown').trim() || 'Unknown';
                }

                function isNavroVcRow(row) {
                    return NAVRO_VC_KEYS.has(getNavroClientTypeCode(row));
                }

                function isHNWIOrAngelRow(row) {
                    if (isNavroDashboard()) {
                        return NAVRO_HNWI_KEYS.has(getNavroClientTypeCode(row));
                    }
                    const t = String(row && row["Type"] || "").toLowerCase();
                    const d = String(row && row["Description"] || "").toLowerCase();
                    return t.includes('angel') || d.includes('angel') || t.includes('individual') || d.includes('individual') || t.includes('hnwi') || d.includes('hnwi');
                }

                function getCompositionGroupMeta(groupKey) {
                    return COMPOSITION_GROUPS.find((group) => group.key === groupKey) || null;
                }

                function getDashboardFieldValue(row, keys) {
                    for (const key of keys) {
                        if (row && Object.prototype.hasOwnProperty.call(row, key)) {
                            const value = String(row[key]).trim().toLowerCase();
                            if (value) return value;
                        }
                    }
                    return '';
                }

                function getCompositionRowsForGroup(groupKey) {
                    if (groupKey === 'vc') return rawData.vc.slice();
                    if (groupKey === 'hnwi') return rawData.fo.filter((row) => isHNWIOrAngelRow(row));
                    if (groupKey === 'fo') return rawData.fo.filter((row) => !isHNWIOrAngelRow(row));
                    return [];
                }

                function polarPoint(cx, cy, radius, angle) {
                    const radians = ((angle - 90) * Math.PI) / 180;
                    return {
                        x: cx + radius * Math.cos(radians),
                        y: cy + radius * Math.sin(radians),
                    };
                }

                function buildArcPath(cx, cy, innerRadius, outerRadius, startAngle, endAngle) {
                    const outerStart = polarPoint(cx, cy, outerRadius, startAngle);
                    const outerEnd = polarPoint(cx, cy, outerRadius, endAngle);
                    const innerEnd = polarPoint(cx, cy, innerRadius, endAngle);
                    const innerStart = polarPoint(cx, cy, innerRadius, startAngle);
                    const largeArc = endAngle - startAngle > 180 ? 1 : 0;

                    return [
                        'M', outerStart.x.toFixed(3), outerStart.y.toFixed(3),
                        'A', outerRadius, outerRadius, 0, largeArc, 1, outerEnd.x.toFixed(3), outerEnd.y.toFixed(3),
                        'L', innerEnd.x.toFixed(3), innerEnd.y.toFixed(3),
                        'A', innerRadius, innerRadius, 0, largeArc, 0, innerStart.x.toFixed(3), innerStart.y.toFixed(3),
                        'Z',
                    ].join(' ');
                }

                function buildCompositionHierarchy() {
                    const groups = [
                        { meta: COMPOSITION_GROUPS[0], rows: getCompositionRowsForGroup('vc') },
                        { meta: COMPOSITION_GROUPS[1], rows: getCompositionRowsForGroup('hnwi') },
                        { meta: COMPOSITION_GROUPS[2], rows: getCompositionRowsForGroup('fo') },
                    ];

                    return groups
                        .map(({ meta, rows }) => {
                            const stageCounts = COMPOSITION_STAGE_ORDER.map((label) => ({ label, count: 0 }));
                            rows.forEach((row) => {
                                const status = getStatusText(row);
                                const match = stageCounts.find((entry) => entry.label === status);
                                if (match) match.count += 1;
                            });
                            return {
                                ...meta,
                                count: rows.length,
                                children: stageCounts
                                    .filter((entry) => entry.count > 0)
                                    .map((entry, index) => ({
                                        ...entry,
                                        shortLabel: entry.label
                                            .replace(' / Moving Forward', '')
                                            .replace(' / Meeting Done', '')
                                            .replace(' / Ongoing', ''),
                                        color: meta.childColors[index % meta.childColors.length],
                                    })),
                            };
                        })
                        .filter((group) => group.count > 0);
                }

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
                    applyDashboardPerformanceMode();
                    setupUiBindings();
                    initializeDashboard();
                });

                function setupUiBindings() {
                    const openLocalBtn = document.getElementById('btn-open-local-file');
                    const fileInput = document.getElementById('file-input');
                    const retryBtn = document.getElementById('btn-retry-sync');
                    const refreshBtn = document.getElementById('btn-refresh-dashboard');
                    const dashboardSwitcher = document.getElementById('dashboard-switcher');
                    const dashboardSwitcherTrigger = document.getElementById('dashboard-switcher-trigger');
                    const dashboardSwitcherMenu = document.getElementById('dashboard-switcher-menu');
                    const dashboardHomePanel = document.getElementById('dashboard-home-panel');
                    const searchInput = document.getElementById('search');
                    const compositionResetBtn = document.getElementById('composition-reset');
                    const dashboardForm = document.getElementById('dashboard-form');
                    const editCurrentBtn = document.getElementById('btn-edit-current-dashboard');
                    const resetDashboardFormBtn = document.getElementById('btn-reset-dashboard-form');
                    const exportBtn = document.getElementById('btn-export-html');
                    const exportMenu = document.getElementById('dashboard-export-menu');
                    const exportConfirmBtn = document.getElementById('btn-export-html-confirm');
                    const exportNamesOnlyBtn = document.getElementById('btn-export-html-names-only');
                    const exportContainer = document.getElementById('dashboard-export');

                    if (openLocalBtn && fileInput) {
                        openLocalBtn.addEventListener('click', () => fileInput.click());
                    }
                    if (fileInput) {
                        fileInput.addEventListener('change', (event) => loadLocalFile(event));
                    }
                    if (retryBtn) {
                        retryBtn.addEventListener('click', retrySync);
                    }
                    if (refreshBtn) {
                        refreshBtn.addEventListener('click', refreshCurrentDashboard);
                    }
                    if (dashboardSwitcherTrigger) {
                        dashboardSwitcherTrigger.addEventListener('click', () => {
                            const isOpen = dashboardSwitcherMenu && !dashboardSwitcherMenu.hidden;
                            if (!dashboardSwitcherMenu) return;
                            dashboardSwitcherMenu.hidden = isOpen;
                            dashboardSwitcherTrigger.setAttribute('aria-expanded', isOpen ? 'false' : 'true');
                        });
                    }
                    if (dashboardSwitcherMenu) {
                        dashboardSwitcherMenu.addEventListener('click', (event) => {
                            const dashboardBtn = event.target.closest('[data-dashboard-switcher-dashboard]');
                            if (dashboardBtn) {
                                const dashboardId = String(dashboardBtn.getAttribute('data-dashboard-switcher-dashboard') || '').trim();
                                closeDashboardSwitcher();
                                if (dashboardId) selectDashboardById(dashboardId);
                                return;
                            }

                            const homeBtn = event.target.closest('[data-dashboard-switcher-home]');
                            if (homeBtn) {
                                closeDashboardSwitcher();
                                showDashboardHome();
                                return;
                            }

                            const missingBtn = event.target.closest('[data-dashboard-switcher-missing]');
                            if (missingBtn) {
                                const dealId = String(missingBtn.getAttribute('data-dashboard-switcher-missing') || '').trim();
                                closeDashboardSwitcher();
                                if (dealId) {
                                    window.location.href = buildPageUrl('deal-details', { id: dealId });
                                }
                            }
                        });
                    }
                    if (dashboardSwitcher) {
                        document.addEventListener('click', (event) => {
                            if (dashboardSwitcher.contains(event.target)) return;
                            closeDashboardSwitcher();
                        });
                        document.addEventListener('keydown', (event) => {
                            if (event.key === 'Escape') closeDashboardSwitcher();
                        });
                    }
                    if (dashboardHomePanel) {
                        dashboardHomePanel.addEventListener('click', (event) => {
                            const openBtn = event.target.closest('[data-dashboard-switcher-dashboard]');
                            if (openBtn) {
                                const dashboardId = String(openBtn.getAttribute('data-dashboard-switcher-dashboard') || '').trim();
                                if (dashboardId) {
                                    selectDashboardById(dashboardId);
                                }
                            }
                        });
                    }
                    if (searchInput) {
                        searchInput.addEventListener('input', filterData);
                    }
                    if (compositionResetBtn) {
                        compositionResetBtn.addEventListener('click', () => clearCompositionSelection());
                    }
                    if (dashboardForm) {
                        dashboardForm.addEventListener('submit', onAddDashboard);
                    }
                    if (editCurrentBtn) {
                        editCurrentBtn.addEventListener('click', () => {
                            if (!currentDashboard) return;
                            populateDashboardForm(currentDashboard);
                            openDashboardSettingsPanel();
                        });
                    }
                    if (resetDashboardFormBtn) {
                        resetDashboardFormBtn.addEventListener('click', () => {
                            resetDashboardFormForCreate();
                            openDashboardSettingsPanel();
                        });
                    }
                    if (exportBtn && exportMenu) {
                        exportBtn.addEventListener('click', () => {
                            const isOpen = !exportMenu.hidden;
                            exportMenu.hidden = isOpen;
                            exportBtn.setAttribute('aria-expanded', isOpen ? 'false' : 'true');
                        });
                    }
                    if (exportConfirmBtn) {
                        exportConfirmBtn.addEventListener('click', exportStandaloneHtml);
                    }
                    if (exportNamesOnlyBtn) {
                        exportNamesOnlyBtn.addEventListener('click', () => exportStandaloneHtml({ exportPreset: 'names-no-emails' }));
                    }
                    if (exportContainer) {
                        document.addEventListener('click', (event) => {
                            if (exportContainer.contains(event.target)) return;
                            closeExportMenu();
                        });
                        document.addEventListener('keydown', (event) => {
                            if (event.key === 'Escape') closeExportMenu();
                        });
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

                    window.addEventListener('appcore:dashboard-config-updated', () => {
                        dashboardsConfig = getMergedDashboardsConfig();
                        syncCurrentDashboardReference();
                        renderDashboardSwitcher();
                        if (currentDashboardMode === 'dashboard' && currentDashboard) {
                            applyDashboardPageChrome(currentDashboard);
                            if (dashboardFormMode !== 'create') {
                                populateDashboardForm(currentDashboard);
                            }
                        } else {
                            applyDashboardPageChrome(null);
                            renderDashboardHome();
                        }
                        updateDashboardActionButtons();
                    });

                    window.addEventListener('appcore:deals-updated', () => {
                        renderDashboardSwitcher();
                        if (currentDashboardMode === 'home') {
                            renderDashboardHome();
                        }
                    });
                }

                function openDashboardSettingsPanel() {
                    const panel = document.getElementById('add-dashboard-panel');
                    if (!panel) return;
                    panel.open = true;
                    panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }

                function getDashboardFormElements() {
                    return {
                        panel: document.getElementById('add-dashboard-panel'),
                        modeEl: document.getElementById('dashboard-form-mode'),
                        helpEl: document.getElementById('dashboard-form-help'),
                        submitBtn: document.getElementById('btn-submit-dashboard-form'),
                        resetBtn: document.getElementById('btn-reset-dashboard-form'),
                        idInput: document.getElementById('dashboard-id'),
                        nameInput: document.getElementById('dashboard-name'),
                        descInput: document.getElementById('dashboard-description'),
                        urlInput: document.getElementById('dashboard-excel-url'),
                        fundsInput: document.getElementById('dashboard-sheet-funds'),
                        foInput: document.getElementById('dashboard-sheet-fo'),
                        figInput: document.getElementById('dashboard-sheet-figures'),
                        makeDefaultInput: document.getElementById('dashboard-make-default'),
                    };
                }

                function applyDashboardPageChrome(dashboard) {
                    const settings = (dashboardsConfig && dashboardsConfig.settings) || {};
                    const pageTitle = settings.title || 'Investor Dashboard';
                    document.title = pageTitle;

                    const titleEl = document.getElementById('dashboard-title');
                    const subtitleEl = document.getElementById('dashboard-subtitle');
                    if (!dashboard) {
                        if (titleEl) titleEl.textContent = pageTitle;
                        if (subtitleEl) subtitleEl.textContent = 'Choose a linked dashboard, or review active deals that still need one.';
                        return;
                    }
                    if (titleEl) titleEl.textContent = dashboard.name || pageTitle;
                    if (subtitleEl) subtitleEl.textContent = dashboard.description || 'Live Dashboard';
                }

                function syncCurrentDashboardReference() {
                    if (!currentDashboard || !dashboardsConfig || !Array.isArray(dashboardsConfig.dashboards)) return;
                    const next = dashboardsConfig.dashboards.find(
                        (dashboard) => normalizeDashboardId(dashboard.id) === normalizeDashboardId(currentDashboard.id),
                    );
                    if (next) {
                        currentDashboard = next;
                        return;
                    }
                    currentDashboard = null;
                    currentDashboardMode = 'home';
                }

                function setDashboardFormMode(mode, dashboard) {
                    dashboardFormMode = mode === 'create' ? 'create' : 'edit';
                    const {
                        modeEl,
                        helpEl,
                        submitBtn,
                        resetBtn,
                        idInput,
                    } = getDashboardFormElements();

                    if (dashboardFormMode === 'create') {
                        if (modeEl) modeEl.textContent = 'Create a new dashboard profile with its Excel link and source sheet names.';
                        if (helpEl) helpEl.textContent = 'This saves the dashboard profile in the app and syncs it to ShareDrive when shared config is available.';
                        if (submitBtn) submitBtn.textContent = 'Create dashboard';
                        if (resetBtn) resetBtn.textContent = 'Clear form';
                        if (idInput) idInput.readOnly = false;
                        return;
                    }

                    if (modeEl) modeEl.textContent = `Editing ${dashboard && (dashboard.name || dashboard.id) ? (dashboard.name || dashboard.id) : 'selected dashboard'}. Update the Excel link or sheet names below.`;
                    if (helpEl) helpEl.textContent = 'This saves the dashboard profile in the app and syncs it to ShareDrive when shared config is available. In edit mode the dashboard ID stays locked so linked deals keep working.';
                    if (submitBtn) submitBtn.textContent = 'Save dashboard changes';
                    if (resetBtn) resetBtn.textContent = 'New dashboard';
                    if (idInput) idInput.readOnly = true;
                }

                function resetDashboardFormForCreate() {
                    const {
                        idInput,
                        nameInput,
                        descInput,
                        urlInput,
                        fundsInput,
                        foInput,
                        figInput,
                        makeDefaultInput,
                    } = getDashboardFormElements();

                    if (idInput) idInput.value = '';
                    if (nameInput) nameInput.value = '';
                    if (descInput) descInput.value = '';
                    if (urlInput) urlInput.value = '';
                    if (fundsInput) fundsInput.value = DEFAULT_SHEET_NAMES.funds;
                    if (foInput) foInput.value = DEFAULT_SHEET_NAMES.familyOffices;
                    if (figInput) figInput.value = DEFAULT_SHEET_NAMES.figures;
                    if (makeDefaultInput) makeDefaultInput.checked = false;
                    setDashboardFormStatus('', false);
                    setDashboardFormMode('create');
                }

                function populateDashboardForm(dashboard) {
                    if (!dashboard) {
                        resetDashboardFormForCreate();
                        return;
                    }

                    const {
                        idInput,
                        nameInput,
                        descInput,
                        urlInput,
                        fundsInput,
                        foInput,
                        figInput,
                        makeDefaultInput,
                    } = getDashboardFormElements();

                    if (idInput) idInput.value = String(dashboard.id || '').trim();
                    if (nameInput) nameInput.value = String(dashboard.name || '').trim();
                    if (descInput) descInput.value = String(dashboard.description || '').trim();
                    if (urlInput) urlInput.value = String(dashboard.excelUrl || '').trim();
                    if (fundsInput) fundsInput.value = String(dashboard.sheets && dashboard.sheets.funds || '').trim() || DEFAULT_SHEET_NAMES.funds;
                    if (foInput) foInput.value = String(dashboard.sheets && dashboard.sheets.familyOffices || '').trim() || DEFAULT_SHEET_NAMES.familyOffices;
                    if (figInput) figInput.value = String(dashboard.sheets && dashboard.sheets.figures || '').trim() || DEFAULT_SHEET_NAMES.figures;
                    if (makeDefaultInput) {
                        const settings = (dashboardsConfig && dashboardsConfig.settings) || {};
                        makeDefaultInput.checked = normalizeDashboardId(settings.defaultDashboard) === normalizeDashboardId(dashboard.id);
                    }
                    setDashboardFormStatus('', false);
                    setDashboardFormMode('edit', dashboard);
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

                        console.log('[Dashboard] Available dashboards:', dashboards.map(d => d.id));
                        console.log('[Dashboard] Requested dashboard ID from URL:', requestedId);

                        renderDashboardSwitcher();
                        updateDashboardActionButtons();

                        if (String(params.get('edit') || '').trim() === '1') {
                            openDashboardSettingsPanel();
                        }

                        if (!dashboards.length) {
                            resetDashboardFormForCreate();
                            showDashboardHome();
                            return;
                        }

                        if (requestedId) {
                            await selectDashboardById(requestedId);
                            return;
                        }

                        if (dashboardFormMode !== 'create') {
                            resetDashboardFormForCreate();
                        }
                        showDashboardHome();
                    } catch (err) {
                        console.error(err);
                        showFailure();
                    }
                }

                async function selectDashboardById(selectedId, options = {}) {
                    const forceReload = Boolean(options && options.forceReload);
                    dashboardsConfig = getMergedDashboardsConfig();
                    if (!dashboardsConfig || !dashboardsConfig.dashboards) return;

                    const normalizedSelectedId = normalizeDashboardId(selectedId);
                    const next = dashboardsConfig.dashboards.find((dashboard) => normalizeDashboardId(dashboard.id) === normalizedSelectedId);
                    if (!next) {
                        showDashboardHome();
                        return;
                    }
                    if (!forceReload && currentDashboardMode === 'dashboard' && currentDashboard && next.id === currentDashboard.id) {
                        return;
                    }

                    currentDashboard = next;
                    currentDashboardMode = 'dashboard';

                    console.log('[Dashboard] Switched dashboard to:', currentDashboard.id, currentDashboard);

                    applyDashboardPageChrome(currentDashboard);
                    renderDashboardSwitcher();
                    updateDashboardActionButtons();
                    if (dashboardFormMode !== 'create') {
                        populateDashboardForm(currentDashboard);
                    }

                    // Show loader while switching datasets
                    document.getElementById('dashboard').style.display = 'none';
                    document.getElementById('loader').style.display = 'flex';
                    const loaderMsg = document.getElementById('loader-msg');
                    if (loaderMsg) loaderMsg.textContent = 'Switching dashboard data...';

                    document.getElementById('sync-type').innerText = "Cloud Sync Active";

                    // Restart sync with new Excel URL
                    if (currentDashboard.excelUrl) {
                        console.log('[Dashboard] Restarting sync for', currentDashboard.id, 'with URL:', currentDashboard.excelUrl);
                        replaceDashboardUrl({ dashboard: currentDashboard.id });
                        await startSync(0, currentDashboard.excelUrl);
                    }
                }

                function normalizeDashboardId(value) {
                    const raw = String(value || "").trim().toLowerCase();
                    if (!raw) return "";
                    return raw
                        .replace(/[^a-z0-9]+/g, "-")
                        .replace(/^-+|-+$/g, "");
                }

                function escapeHtml(value) {
                    return String(value == null ? '' : value)
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;')
                        .replace(/'/g, '&#39;');
                }

                function getDealsData() {
                    if (window.AppCore && typeof window.AppCore.loadDealsData === 'function') {
                        return window.AppCore.loadDealsData();
                    }
                    return Array.isArray(window.DEALS) ? window.DEALS : [];
                }

                function getActiveDealsData() {
                    return (Array.isArray(getDealsData()) ? getDealsData() : []).filter((deal) => isActiveDeal(deal));
                }

                function getDashboardDirectoryModel() {
                    const dashboards = (dashboardsConfig && Array.isArray(dashboardsConfig.dashboards)) ? dashboardsConfig.dashboards : [];
                    const settings = (dashboardsConfig && dashboardsConfig.settings) || {};
                    const activeDeals = getActiveDealsData();
                    const linkedDashboardIds = new Set();
                    const linkedDeals = [];
                    const missingDeals = [];

                    activeDeals.forEach((deal) => {
                        const dashboardId = normalizeDashboardId(deal && deal.fundraisingDashboardId);
                        const dashboard = dashboards.find((entry) => normalizeDashboardId(entry && entry.id) === dashboardId) || null;
                        if (dashboard) {
                            linkedDashboardIds.add(normalizeDashboardId(dashboard.id));
                            linkedDeals.push({ deal, dashboard });
                            return;
                        }
                        missingDeals.push({ deal });
                    });

                    const otherDashboards = dashboards
                        .filter((dashboard) => !linkedDashboardIds.has(normalizeDashboardId(dashboard && dashboard.id)))
                        .map((dashboard) => ({ dashboard }));

                    linkedDeals.sort((a, b) => {
                        const aLabel = String(a.deal && (a.deal.name || a.dashboard && a.dashboard.name || a.dashboard && a.dashboard.id) || '').trim();
                        const bLabel = String(b.deal && (b.deal.name || b.dashboard && b.dashboard.name || b.dashboard && b.dashboard.id) || '').trim();
                        return aLabel.localeCompare(bLabel);
                    });
                    missingDeals.sort((a, b) => {
                        const aLabel = String(a.deal && (a.deal.name || a.deal.company || a.deal.id) || '').trim();
                        const bLabel = String(b.deal && (b.deal.name || b.deal.company || b.deal.id) || '').trim();
                        return aLabel.localeCompare(bLabel);
                    });
                    otherDashboards.sort((a, b) => {
                        const aLabel = String(a.dashboard && (a.dashboard.name || a.dashboard.id) || '').trim();
                        const bLabel = String(b.dashboard && (b.dashboard.name || b.dashboard.id) || '').trim();
                        return aLabel.localeCompare(bLabel);
                    });

                    return {
                        defaultDashboardId: normalizeDashboardId(settings.defaultDashboard),
                        linkedDeals,
                        missingDeals,
                        otherDashboards,
                    };
                }

                function closeDashboardSwitcher() {
                    const menu = document.getElementById('dashboard-switcher-menu');
                    const trigger = document.getElementById('dashboard-switcher-trigger');
                    if (menu) menu.hidden = true;
                    if (trigger) trigger.setAttribute('aria-expanded', 'false');
                }

                function closeExportMenu() {
                    const menu = document.getElementById('dashboard-export-menu');
                    const trigger = document.getElementById('btn-export-html');
                    if (menu) menu.hidden = true;
                    if (trigger) trigger.setAttribute('aria-expanded', 'false');
                }

                async function refreshCurrentDashboard() {
                    if (refreshButtonBusy) return;

                    refreshButtonBusy = true;
                    updateDashboardActionButtons();
                    closeDashboardSwitcher();
                    closeExportMenu();

                    const loader = document.getElementById('loader');
                    const spinner = document.getElementById('loading-spinner');
                    const errorPanel = document.getElementById('error-panel');
                    const loaderMsg = document.getElementById('loader-msg');
                    const dashboardEl = document.getElementById('dashboard');

                    if (errorPanel) errorPanel.style.display = 'none';
                    if (spinner) spinner.style.display = 'block';
                    if (loader) loader.style.display = 'flex';
                    if (dashboardEl && currentDashboardMode === 'dashboard') {
                        dashboardEl.style.display = 'none';
                    }
                    if (loaderMsg) {
                        loaderMsg.textContent = currentDashboardMode === 'dashboard'
                            ? 'Refreshing dashboard data...'
                            : 'Refreshing dashboards...';
                    }

                    try {
                        if (window.AppCore && typeof window.AppCore.refreshDashboardConfigFromShareDrive === "function") {
                            await window.AppCore.refreshDashboardConfigFromShareDrive();
                        }
                        dashboardsConfig = getMergedDashboardsConfig();
                        renderDashboardSwitcher();
                        updateDashboardActionButtons();

                        if (currentDashboardMode === 'dashboard' && currentDashboard && currentDashboard.id) {
                            await selectDashboardById(currentDashboard.id, { forceReload: true });
                        } else {
                            await initializeDashboard();
                        }
                    } catch (err) {
                        console.error(err);
                        showFailure();
                    } finally {
                        refreshButtonBusy = false;
                        updateDashboardActionButtons();
                    }
                }

                function updateDashboardActionButtons() {
                    const refreshBtn = document.getElementById('btn-refresh-dashboard');
                    const exportBtn = document.getElementById('btn-export-html');
                    const exportConfirmBtn = document.getElementById('btn-export-html-confirm');
                    const exportNamesOnlyBtn = document.getElementById('btn-export-html-names-only');
                    const editBtn = document.getElementById('btn-edit-current-dashboard');
                    const canExport = currentDashboardMode === 'dashboard' && currentDashboard;

                    if (refreshBtn) {
                        refreshBtn.disabled = refreshButtonBusy;
                        refreshBtn.textContent = refreshButtonBusy
                            ? 'Refreshing...'
                            : (currentDashboardMode === 'dashboard' && currentDashboard
                                ? 'Refresh data'
                                : 'Refresh dashboards');
                    }
                    if (exportBtn) {
                        exportBtn.disabled = !canExport;
                        if (!canExport) closeExportMenu();
                    }
                    if (exportConfirmBtn) {
                        exportConfirmBtn.disabled = !canExport;
                    }
                    if (exportNamesOnlyBtn) {
                        exportNamesOnlyBtn.disabled = !canExport;
                    }
                    if (editBtn) {
                        editBtn.disabled = !canExport;
                        editBtn.textContent = canExport
                            ? 'Edit selected dashboard'
                            : 'Select a dashboard to edit';
                    }
                }

                function renderDashboardSwitcher() {
                    const menu = document.getElementById('dashboard-switcher-menu');
                    const valueEl = document.getElementById('dashboard-switcher-value');
                    const warningEl = document.getElementById('dashboard-switcher-warning');
                    const directory = getDashboardDirectoryModel();
                    if (!menu || !valueEl || !warningEl) return;

                    const missingCount = directory.missingDeals.length;
                    warningEl.hidden = !missingCount;
                    warningEl.textContent = missingCount === 1 ? '1 missing link' : `${missingCount} missing links`;

                    if (currentDashboardMode === 'dashboard' && currentDashboard) {
                        const linkedDeal = directory.linkedDeals.find((entry) => normalizeDashboardId(entry.dashboard && entry.dashboard.id) === normalizeDashboardId(currentDashboard.id));
                        valueEl.textContent = linkedDeal
                            ? `${linkedDeal.deal.name || linkedDeal.dashboard.name || linkedDeal.dashboard.id}`
                            : (currentDashboard.name || currentDashboard.id || 'Dashboard');
                    } else {
                        valueEl.textContent = 'Investor Dashboard Home';
                    }

                    const renderLinkedItem = ({ deal, dashboard }) => {
                        const isActive = currentDashboardMode === 'dashboard' && currentDashboard && normalizeDashboardId(currentDashboard.id) === normalizeDashboardId(dashboard.id);
                        const isDefault = directory.defaultDashboardId && normalizeDashboardId(dashboard.id) === directory.defaultDashboardId;
                        return `
                            <button class="dashboard-switcher-item${isActive ? ' is-active' : ''}" type="button" data-dashboard-switcher-dashboard="${escapeHtml(dashboard.id || '')}">
                                <span class="dashboard-switcher-item-copy">
                                    <span class="dashboard-switcher-item-title">${escapeHtml(deal.name || dashboard.name || dashboard.id)}</span>
                                    <span class="dashboard-switcher-item-meta">${escapeHtml(deal.company || 'No company')} · ${escapeHtml(dashboard.name || dashboard.id || '')}</span>
                                </span>
                                ${isDefault ? '<span class="dashboard-switcher-pill">Default</span>' : ''}
                            </button>
                        `;
                    };

                    const renderMissingItem = ({ deal }) => `
                        <button class="dashboard-switcher-item is-missing" type="button" data-dashboard-switcher-missing="${escapeHtml(deal.id || '')}">
                            <span class="dashboard-switcher-item-copy">
                                <span class="dashboard-switcher-item-title">${escapeHtml(deal.name || deal.company || deal.id || 'Deal')}</span>
                                <span class="dashboard-switcher-item-meta">${escapeHtml(deal.company || 'No company')} · No linked dashboard</span>
                            </span>
                            <span class="dashboard-switcher-pill is-danger">Needs link</span>
                        </button>
                    `;

                    const renderOtherItem = ({ dashboard }) => {
                        const isActive = currentDashboardMode === 'dashboard' && currentDashboard && normalizeDashboardId(currentDashboard.id) === normalizeDashboardId(dashboard.id);
                        const isDefault = directory.defaultDashboardId && normalizeDashboardId(dashboard.id) === directory.defaultDashboardId;
                        return `
                            <button class="dashboard-switcher-item${isActive ? ' is-active' : ''}" type="button" data-dashboard-switcher-dashboard="${escapeHtml(dashboard.id || '')}">
                                <span class="dashboard-switcher-item-copy">
                                    <span class="dashboard-switcher-item-title">${escapeHtml(dashboard.name || dashboard.id)}</span>
                                    <span class="dashboard-switcher-item-meta">Standalone dashboard profile</span>
                                </span>
                                ${isDefault ? '<span class="dashboard-switcher-pill">Default</span>' : ''}
                            </button>
                        `;
                    };

                    menu.innerHTML = `
                        <div class="dashboard-switcher-section">
                            <div class="dashboard-switcher-section-label">Home</div>
                            <button class="dashboard-switcher-item${currentDashboardMode === 'home' ? ' is-active' : ''}" type="button" data-dashboard-switcher-home="1">
                                <span class="dashboard-switcher-item-copy">
                                    <span class="dashboard-switcher-item-title">Investor Dashboard Home</span>
                                    <span class="dashboard-switcher-item-meta">Overview, linked dashboards, and missing dashboard links</span>
                                </span>
                            </button>
                        </div>
                        <div class="dashboard-switcher-section">
                            <div class="dashboard-switcher-section-label">Linked dashboards</div>
                            ${directory.linkedDeals.length
                                ? directory.linkedDeals.map(renderLinkedItem).join('')
                                : '<div class="dashboard-switcher-empty">No active deals currently have a linked dashboard.</div>'}
                        </div>
                        <div class="dashboard-switcher-section is-alert">
                            <div class="dashboard-switcher-section-label">Needs dashboard link</div>
                            ${directory.missingDeals.length
                                ? directory.missingDeals.map(renderMissingItem).join('')
                                : '<div class="dashboard-switcher-empty">Every active deal is linked to a dashboard.</div>'}
                        </div>
                        ${directory.otherDashboards.length ? `
                            <div class="dashboard-switcher-section">
                                <div class="dashboard-switcher-section-label">Other dashboards</div>
                                ${directory.otherDashboards.map(renderOtherItem).join('')}
                            </div>
                        ` : ''}
                    `;
                }

                function renderDashboardHome() {
                    const panel = document.getElementById('dashboard-home-panel');
                    if (!panel) return;
                    const directory = getDashboardDirectoryModel();
                    const defaultLinked = directory.linkedDeals.find((entry) => normalizeDashboardId(entry.dashboard && entry.dashboard.id) === directory.defaultDashboardId) || null;

                    const renderLinkedCards = directory.linkedDeals.length
                        ? directory.linkedDeals.slice(0, 8).map(({ deal, dashboard }) => `
                            <button class="dashboard-home-item" type="button" data-dashboard-switcher-dashboard="${escapeHtml(dashboard.id || '')}">
                                <span class="dashboard-home-item-title">${escapeHtml(deal.name || dashboard.name || dashboard.id)}</span>
                                <span class="dashboard-home-item-meta">${escapeHtml(deal.company || 'No company')} · ${escapeHtml(dashboard.name || dashboard.id || '')}</span>
                            </button>
                        `).join('')
                        : '<div class="dashboard-home-empty">No linked dashboards yet. Create one or attach one from a deal.</div>';

                    const renderMissingCards = directory.missingDeals.length
                        ? directory.missingDeals.map(({ deal }) => `
                            <a class="dashboard-home-item is-missing" href="${buildPageUrl('deal-details', { id: deal.id })}">
                                <span class="dashboard-home-item-title">${escapeHtml(deal.name || deal.company || deal.id || 'Deal')}</span>
                                <span class="dashboard-home-item-meta">${escapeHtml(deal.company || 'No company')} · Open deal to link a dashboard</span>
                            </a>
                        `).join('')
                        : '<div class="dashboard-home-empty">No missing dashboard links across active deals.</div>';

                    panel.innerHTML = `
                        <div class="dashboard-home-hero">
                            <div>
                                <div class="card-title">Investor Dashboard Home</div>
                                <p class="dashboard-home-copy">Pick a linked dashboard to load live investor data, or use the red list to catch active deals that still need a fundraising dashboard attached.</p>
                            </div>
                            <div class="dashboard-home-actions">
                                ${defaultLinked ? `<button class="btn btn-primary" type="button" data-dashboard-switcher-dashboard="${escapeHtml(defaultLinked.dashboard.id || '')}">Open default dashboard</button>` : ''}
                                <button class="btn" id="btn-dashboard-home-create" type="button">Create dashboard</button>
                            </div>
                        </div>
                        <div class="dashboard-home-stats">
                            <div class="dashboard-home-stat">
                                <span class="dashboard-home-stat-label">Linked dashboards</span>
                                <strong>${directory.linkedDeals.length}</strong>
                            </div>
                            <div class="dashboard-home-stat${directory.missingDeals.length ? ' is-danger' : ''}">
                                <span class="dashboard-home-stat-label">Deals missing dashboard</span>
                                <strong>${directory.missingDeals.length}</strong>
                            </div>
                            <div class="dashboard-home-stat">
                                <span class="dashboard-home-stat-label">Standalone dashboards</span>
                                <strong>${directory.otherDashboards.length}</strong>
                            </div>
                        </div>
                        <div class="dashboard-home-grid">
                            <div class="dashboard-home-column">
                                <div class="dashboard-home-column-title">Ready to open</div>
                                ${renderLinkedCards}
                            </div>
                            <div class="dashboard-home-column is-alert">
                                <div class="dashboard-home-column-title">Needs dashboard link</div>
                                ${renderMissingCards}
                            </div>
                        </div>
                    `;

                    const createBtn = document.getElementById('btn-dashboard-home-create');
                    if (createBtn) {
                        createBtn.addEventListener('click', () => {
                            resetDashboardFormForCreate();
                            openDashboardSettingsPanel();
                        });
                    }
                }

                function replaceDashboardUrl(params) {
                    const nextHref = buildPageUrl('investor-dashboard', params);
                    if (window.history && typeof window.history.replaceState === 'function') {
                        window.history.replaceState({}, '', nextHref);
                    }
                }

                function showDashboardHome() {
                    currentDashboard = null;
                    currentDashboardMode = 'home';
                    const homePanel = document.getElementById('dashboard-home-panel');
                    const dataView = document.getElementById('dashboard-data-view');
                    if (homePanel) homePanel.hidden = false;
                    if (dataView) dataView.hidden = true;
                    document.getElementById('loader').style.display = 'none';
                    document.getElementById('dashboard').style.display = 'block';
                    applyDashboardPageChrome(null);
                    renderDashboardHome();
                    renderDashboardSwitcher();
                    updateDashboardActionButtons();
                    replaceDashboardUrl({});
                }

                function showDashboardDatasetView() {
                    const homePanel = document.getElementById('dashboard-home-panel');
                    const dataView = document.getElementById('dashboard-data-view');
                    if (homePanel) homePanel.hidden = true;
                    if (dataView) dataView.hidden = false;
                    document.getElementById('dashboard').style.display = 'block';
                    updateDashboardActionButtons();
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
                        currentDashboard = (updatedConfig && Array.isArray(updatedConfig.dashboards))
                            ? (updatedConfig.dashboards.find(d => normalizeDashboardId(d.id) === dashboardId) || currentDashboard)
                            : currentDashboard;

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
                        return startSync(proxyIndex + 1, excelUrl, true); // Try next proxy
                    }
                }

                function showFailure() {
                    document.getElementById('loading-spinner').style.display = 'none';
                    document.getElementById('error-panel').style.display = 'block';
                }

                function retrySync() {
                    document.getElementById('error-panel').style.display = 'none';
                    document.getElementById('loading-spinner').style.display = 'block';
                    if (currentDashboardMode === 'dashboard' && currentDashboard && currentDashboard.excelUrl) {
                        startSync(0, currentDashboard.excelUrl);
                    } else {
                        showDashboardHome();
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
                    const sameSheetConfigured =
                        normalizeDashboardText(sheetNames.funds) &&
                        normalizeDashboardText(sheetNames.funds) === normalizeDashboardText(sheetNames.familyOffices);

                    if (isNavroDashboard() && sameSheetConfigured && fSheet) {
                        const navroRows = XLSX.utils.sheet_to_json(fSheet, { range: headerRowIndex, defval: "" })
                            .filter(r => r["Investor"]);
                        rawData.vc = navroRows.filter((row) => isNavroVcRow(row));
                        rawData.fo = navroRows.filter((row) => {
                            const clientTypeCode = getNavroClientTypeCode(row);
                            if (clientTypeCode) return !NAVRO_VC_KEYS.has(clientTypeCode);
                            return !isNavroVcRow(row);
                        });
                    } else {
                        rawData.vc = XLSX.utils.sheet_to_json(fSheet, { range: headerRowIndex, defval: "" }).filter(r => r["Investor"]);
                        rawData.fo = XLSX.utils.sheet_to_json(oSheet, { range: headerRowIndex, defval: "" }).filter(r => r["Investor"]);
                    }

                    prepareDashboardRows(rawData.vc);
                    prepareDashboardRows(rawData.fo);
                    const figures = gSheet ? XLSX.utils.sheet_to_json(gSheet, { defval: "" }) : [];

                    const combined = [...rawData.vc, ...rawData.fo];

                    console.log('[Dashboard] Parsed row counts -> vc:', rawData.vc.length, 'fo:', rawData.fo.length, 'combined:', combined.length);

                    // 1. Dynamic Type Color Mapping
                    const uniqueTypes = [...new Set(combined.map(r => getInvestorTypeLabel(r)).filter(t => t))];
                    typeColors = {};
                    uniqueTypes.forEach((type, idx) => {
                        typeColors[type] = COLOR_PALETTE[idx % COLOR_PALETTE.length];
                    });

                    // 2. Robust KPIs (Exhaustive Source-Aware Count)
                    // Rule: all rows on the Funds sheet are VC. FO/HNWI are split only within the Family Offices sheet.
                    let vcCount = rawData.vc.length;
                    let hnwiCount = 0;
                    let foCount = 0;

                    const hasTypeSignals = rawData.fo.some(r =>
                        Object.prototype.hasOwnProperty.call(r, "Type of Client") ||
                        Object.prototype.hasOwnProperty.call(r, "Type Of Client") ||
                        Object.prototype.hasOwnProperty.call(r, "type of client") ||
                        Object.prototype.hasOwnProperty.call(r, "Type") ||
                        Object.prototype.hasOwnProperty.call(r, "KEY") ||
                        Object.prototype.hasOwnProperty.call(r, "Key") ||
                        Object.prototype.hasOwnProperty.call(r, "key")
                    );
                    console.log('[Dashboard] Has type signal on FO sheet:', hasTypeSignals);

                    if (!hasTypeSignals) {
                        // If there's no "Type" column on the FO sheet at all, treat all FO rows as Family Offices
                        foCount = rawData.fo.length;
                        hnwiCount = 0;
                    } else {
                        rawData.fo.forEach(r => {
                            if (isHNWIOrAngelRow(r)) hnwiCount++;
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
                    const getContactStatus = (row) => getDashboardFieldValue(row, CONTACT_FIELD_KEYS);
                    const getCallStatus = (row) => getDashboardFieldValue(row, CALL_FIELD_KEYS);

                    const ongoingCount = combined.filter(r => {
                        const s = getCallStatus(r);
                        return s === 'yes' || s === 'waiting';
                    }).length;
                    const totalContactCount = combined.filter(r => {
                        const s = getContactStatus(r);
                        return s === 'yes' || s === 'waiting';
                    }).length;
                    const repliedCount = countExact(FORWARD_FIELD_KEYS, 'yes');
                    const meetingsCount = countExact(MEETING_FIELD_KEYS, 'yes');
                    const isCurrentDashboardNavro = isNavroDashboard();
                    const contactedCardValue = totalContactCount;
                    const meetingsCardLabel = isCurrentDashboardNavro ? 'Ongoing' : 'Meetings';
                    const meetingsCardValue = isCurrentDashboardNavro ? ongoingCount : meetingsCount;

                    const poolSize = combined.length;

                    const setText = (id, value) => {
                        const el = document.getElementById(id);
                        if (el) el.innerText = value;
                    };

                    setText('k-pool', poolSize);
                    setText('k-contacted', contactedCardValue);
                    setText('k-replied', repliedCount);
                    setText('k-meetings-label', meetingsCardLabel);
                    setText('k-meetings', meetingsCardValue);

                    // 4. Visuals Refresh
                    activeFilters = normalizeQuickFilters(activeFilters);
                    applyQuickFilterButtons();
                    renderRangeBars();
                    resetCompositionSelectionState();
                    switchTab('vc', document.querySelector('[data-tab-type="vc"]'), { keepCompositionSelection: true });

                    // 5. Show Dashboard
                    document.getElementById('loader').style.display = 'none';
                    showDashboardDatasetView();
                    renderDashboardSwitcher();
                }

                function getStatusText(item) {
                    const forward = getDashboardFieldValue(item, FORWARD_FIELD_KEYS);
                    const meeting = getDashboardFieldValue(item, MEETING_FIELD_KEYS);
                    const call = getDashboardFieldValue(item, [...CALL_FIELD_KEYS, ...CONTACT_FIELD_KEYS]);

                    if (forward === 'yes') return 'Replied / Moving Forward';
                    if (forward === 'waiting' || meeting === 'waiting' || call === 'waiting') return 'Waiting / Ongoing';
                    if (forward === 'no') return '';
                    if (meeting === 'yes') return 'Contacted / Meeting Done';
                    if (call === 'yes') return 'Contact Started';
                    return 'Target';
                }

                function buildDashboardSearchIndex(item) {
                    return DASHBOARD_SEARCH_FIELDS
                        .map((field) => String(item && item[field] || '').trim())
                        .filter(Boolean)
                        .join(' ')
                        .toLowerCase();
                }

                function prepareDashboardRows(rows) {
                    (rows || []).forEach((item) => {
                        if (!item || typeof item !== 'object') return;
                        item.__dashboardName = String(item["Investor"] || item["Investor Name"] || item["Name"] || '').trim();
                        item.__dashboardSize = String(item["Size of Investment"] || item["Investment Size"] || '–').trim() || '–';
                        item.__dashboardEmail = String(item["Email"] || '–').trim() || '–';
                        item.__dashboardType = getInvestorTypeLabel(item);
                        item.__dashboardStatusText = getStatusText(item);
                        item.__dashboardStatusBadge = getStatusBadge(item);
                        item.__dashboardSearchIndex = buildDashboardSearchIndex(item);
                    });
                }

                function resetCompositionSelectionState() {
                    compositionSelection = { groupKey: '', stageLabel: '' };
                }

                function getNormalizedCompositionStage(stageLabel) {
                    const normalized = normalizeDashboardText(stageLabel);
                    return COMPOSITION_STAGE_ORDER.find((label) => normalizeDashboardText(label) === normalized) || '';
                }

                function getCurrentSearchKeyword() {
                    return String(document.getElementById('search') && document.getElementById('search').value || '').toLowerCase();
                }

                function getCompositionSelectionCount() {
                    if (!compositionSelection.groupKey) return 0;
                    const rows = getCompositionRowsForGroup(compositionSelection.groupKey);
                    if (!compositionSelection.stageLabel) return rows.length;
                    return rows.filter((row) => getStatusText(row) === compositionSelection.stageLabel).length;
                }

                function updateCompositionSelectionUi() {
                    const caption = document.getElementById('composition-caption');
                    const resetBtn = document.getElementById('composition-reset');
                    const hasSelection = Boolean(compositionSelection.groupKey);

                    if (caption) {
                        if (!hasSelection) {
                            caption.textContent = 'Click a ring segment to filter the investor table.';
                        } else {
                            const groupMeta = getCompositionGroupMeta(compositionSelection.groupKey);
                            const groupLabel = groupMeta ? groupMeta.label : 'Selected investors';
                            const count = getCompositionSelectionCount();
                            const countLabel = `${count} investor${count === 1 ? '' : 's'}`;
                            caption.textContent = compositionSelection.stageLabel
                                ? `Filtering ${groupLabel} / ${compositionSelection.stageLabel} / ${countLabel}.`
                                : `Filtering ${groupLabel} / ${countLabel}.`;
                        }
                    }

                    if (resetBtn) {
                        resetBtn.hidden = !hasSelection;
                        resetBtn.disabled = !hasSelection;
                    }
                }

                function clearCompositionSelection(options = {}) {
                    resetCompositionSelectionState();
                    if (options.render === false) return;
                    renderCompositionChart();
                    scheduleTableRender(getCurrentSearchKeyword());
                    updateCompositionSelectionUi();
                }

                function setCompositionSelection(groupKey, stageLabel) {
                    const groupMeta = getCompositionGroupMeta(groupKey);
                    if (!groupMeta) {
                        clearCompositionSelection();
                        return;
                    }

                    const normalizedStage = getNormalizedCompositionStage(stageLabel);
                    const isSameSelection =
                        compositionSelection.groupKey === groupKey &&
                        compositionSelection.stageLabel === normalizedStage;

                    if (isSameSelection) {
                        clearCompositionSelection();
                        return;
                    }

                    compositionSelection = {
                        groupKey,
                        stageLabel: normalizedStage,
                    };

                    const nextType = groupKey === 'vc' ? 'vc' : 'fo';
                    const nextButton = document.querySelector(`[data-tab-type="${nextType}"]`);
                    switchTab(nextType, nextButton, { keepCompositionSelection: true });
                }

                function matchesCompositionSelection(item) {
                    if (!compositionSelection.groupKey) return true;

                    if (compositionSelection.groupKey === 'vc') {
                        if (activeType !== 'vc') return false;
                    } else {
                        if (activeType !== 'fo') return false;
                        const isHnwi = isHNWIOrAngelRow(item);
                        if (compositionSelection.groupKey === 'hnwi' && !isHnwi) return false;
                        if (compositionSelection.groupKey === 'fo' && isHnwi) return false;
                    }

                    if (compositionSelection.stageLabel && getStatusText(item) !== compositionSelection.stageLabel) {
                        return false;
                    }

                    return true;
                }

                function handleCompositionSelectionClick(event) {
                    const target = event.currentTarget;
                    const groupKey = String(target && target.getAttribute('data-composition-group') || '').trim();
                    const stageLabel = String(target && target.getAttribute('data-composition-stage') || '').trim();
                    if (!groupKey) return;
                    setCompositionSelection(groupKey, stageLabel);
                }

                function handleCompositionSelectionKeydown(event) {
                    if (event.key !== 'Enter' && event.key !== ' ') return;
                    event.preventDefault();
                    handleCompositionSelectionClick(event);
                }

                function renderCompositionChart() {
                    const svg = document.getElementById('composition-sunburst');
                    const legend = document.getElementById('composition-legend');
                    if (!svg || !legend) return;

                    const hierarchy = buildCompositionHierarchy();
                    const total = hierarchy.reduce((sum, group) => sum + group.count, 0);
                    if (!total) {
                        svg.innerHTML = '';
                        legend.innerHTML = '';
                        updateCompositionSelectionUi();
                        return;
                    }

                    const hasSelection = Boolean(compositionSelection.groupKey);
                    let startAngle = 0;
                    let markup = '';

                    hierarchy.forEach((group) => {
                        const span = (group.count / total) * 360;
                        const endAngle = startAngle + span;
                        const groupIsActive = compositionSelection.groupKey === group.key;
                        const groupIsMuted = hasSelection && !groupIsActive;
                        markup += `<path class="composition-segment is-clickable${groupIsActive ? ' is-active' : ''}${groupIsMuted ? ' is-muted' : ''}" data-composition-group="${group.key}" d="${buildArcPath(50, 50, 18, 34, startAngle, endAngle)}" fill="${group.color}" aria-label="${group.label}" />`;

                        if (span >= 18) {
                            const innerPoint = polarPoint(50, 50, 26, startAngle + span / 2);
                            markup += `<text x="${innerPoint.x.toFixed(2)}" y="${innerPoint.y.toFixed(2)}" class="composition-label" font-size="${span > 70 ? 4.2 : 3.2}">${group.shortLabel} (${group.count})</text>`;
                        }

                        let childStartAngle = startAngle;
                        group.children.forEach((child) => {
                            const childSpan = (child.count / total) * 360;
                            const childEndAngle = childStartAngle + childSpan;
                            const childIsActive =
                                compositionSelection.groupKey === group.key &&
                                (compositionSelection.stageLabel
                                    ? compositionSelection.stageLabel === child.label
                                    : true);
                            const childIsMuted =
                                hasSelection &&
                                (
                                    compositionSelection.groupKey !== group.key ||
                                    (compositionSelection.stageLabel && compositionSelection.stageLabel !== child.label)
                                );
                            markup += `<path class="composition-segment is-clickable${childIsActive ? ' is-active' : ''}${childIsMuted ? ' is-muted' : ''}" data-composition-group="${group.key}" data-composition-stage="${child.label}" d="${buildArcPath(50, 50, 36, 49.5, childStartAngle, childEndAngle)}" fill="${child.color}" aria-label="${group.label} ${child.label}" />`;

                            if (childSpan >= 14) {
                                const outerPoint = polarPoint(50, 50, 42.5, childStartAngle + childSpan / 2);
                                markup += `<text x="${outerPoint.x.toFixed(2)}" y="${outerPoint.y.toFixed(2)}" class="composition-label" font-size="${childSpan > 28 ? 2.8 : 2.2}">${child.shortLabel}</text>`;
                            }

                            childStartAngle = childEndAngle;
                        });

                        startAngle = endAngle;
                    });

                    markup += `<circle cx="50" cy="50" r="11.5" fill="rgba(255,255,255,0.06)" stroke="rgba(148,163,184,0.16)" />`;
                    markup += `<text x="50" y="48.6" class="composition-center-note">pool</text>`;
                    markup += `<text x="50" y="55.3" class="composition-center-total">${total}</text>`;
                    svg.innerHTML = markup;

                    legend.innerHTML = hierarchy.map((group) => `
                        <div class="composition-legend-group">
                            <div class="composition-legend-head composition-interactive${compositionSelection.groupKey === group.key ? ' is-active' : ''}${hasSelection && compositionSelection.groupKey !== group.key ? ' is-muted' : ''}" data-composition-group="${group.key}" role="button" tabindex="0" aria-pressed="${compositionSelection.groupKey === group.key && !compositionSelection.stageLabel ? 'true' : 'false'}">
                                <div class="composition-legend-title">
                                    <span class="composition-legend-dot" style="background:${group.color}"></span>
                                    <span>${group.label}</span>
                                </div>
                                <span class="composition-legend-total">${group.count}</span>
                            </div>
                            ${group.children.map((child) => `
                                <div class="composition-legend-item composition-interactive${compositionSelection.groupKey === group.key && (!compositionSelection.stageLabel || compositionSelection.stageLabel === child.label) ? ' is-active' : ''}${hasSelection && (compositionSelection.groupKey !== group.key || (compositionSelection.stageLabel && compositionSelection.stageLabel !== child.label)) ? ' is-muted' : ''}" data-composition-group="${group.key}" data-composition-stage="${child.label}" role="button" tabindex="0" aria-pressed="${compositionSelection.groupKey === group.key && compositionSelection.stageLabel === child.label ? 'true' : 'false'}">
                                    <span class="composition-legend-swatch" style="background:${child.color}"></span>
                                    <span>${child.label}</span>
                                    <strong>${child.count}</strong>
                                </div>
                            `).join('')}
                        </div>
                    `).join('');

                    svg.querySelectorAll('[data-composition-group]').forEach((segment) => {
                        segment.addEventListener('click', handleCompositionSelectionClick);
                    });
                    legend.querySelectorAll('[data-composition-group]').forEach((item) => {
                        item.addEventListener('click', handleCompositionSelectionClick);
                        item.addEventListener('keydown', handleCompositionSelectionKeydown);
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
                    if (!container) return;

                    const markup = BUCKETS.map((bucket, i) => {
                        const count = counts[i];
                        const perc = (count / maxCount) * 100;
                        return `
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
                    }).join('');

                    container.innerHTML = markup;
                }

                function switchTab(type, btn, options = {}) {
                    activeType = type;
                    if (!options.keepCompositionSelection) {
                        resetCompositionSelectionState();
                    }
                    const activeButton = btn || document.querySelector(`[data-tab-type="${type}"]`);
                    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                    if (activeButton) activeButton.classList.add('active');

                    const header = document.getElementById('table-head');
                    if (type === 'vc') {
                        header.innerHTML = '<tr><th>VC Funds Name</th><th>Investment Size</th><th>Contact Email</th></tr>';
                    } else {
                        header.innerHTML = '<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th></tr>';
                    }
                    renderCompositionChart();
                    scheduleTableRender(getCurrentSearchKeyword());
                    updateCompositionSelectionUi();
                }

                function toggleFilter(filter) {
                    const normalizedFilter = ['all', 'calls', 'meetings', 'forward'].includes(filter) ? filter : 'all';
                    activeFilters = normalizeQuickFilters(activeFilters);
                    activeFilters = normalizedFilter === 'all'
                        ? normalizeQuickFilters({ all: true })
                        : normalizeQuickFilters({ [normalizedFilter]: !activeFilters[normalizedFilter] });
                    applyQuickFilterButtons();
                    scheduleTableRender(getCurrentSearchKeyword());
                }

                function getStatusBadge(item) {
                    const forward = getDashboardFieldValue(item, FORWARD_FIELD_KEYS);
                    const meeting = getDashboardFieldValue(item, MEETING_FIELD_KEYS);
                    const call = getDashboardFieldValue(item, [...CALL_FIELD_KEYS, ...CONTACT_FIELD_KEYS]);

                    if (forward === 'yes') return `<span class="stage-pill badge-green">Replied / Mov. Forward</span>`;
                    if (forward === 'waiting' || meeting === 'waiting' || call === 'waiting') return `<span class="stage-pill badge-amber">Waiting / Ongoing</span>`;
                    if (forward === 'no') return '';
                    if (meeting === 'yes') return `<span class="stage-pill badge-neutral">Contacted / Met</span>`;
                    if (call === 'yes') return `<span class="stage-pill badge-neutral">Contact Started</span>`;
                    return `<span class="stage-pill" style="color:var(--text-dim); opacity:0.5;">Target</span>`;
                }

                function normalizeQuickFilters(filters) {
                    const source = filters && typeof filters === 'object' ? filters : {};
                    const next = { all: false, calls: false, meetings: false, forward: false };
                    if (source.calls) next.calls = true;
                    else if (source.meetings) next.meetings = true;
                    else if (source.forward) next.forward = true;
                    else next.all = true;
                    return next;
                }

                function applyQuickFilterButtons() {
                    Object.keys(activeFilters).forEach((filterKey) => {
                        const btn = document.getElementById(`f-${filterKey}`);
                        if (!btn) return;
                        btn.classList.toggle('active', Boolean(activeFilters[filterKey]));
                    });
                }

                function matchesQuickFilter(item, filterKey) {
                    const contact = getDashboardFieldValue(item, CONTACT_FIELD_KEYS);
                    const call = getDashboardFieldValue(item, CALL_FIELD_KEYS);
                    const forward = getDashboardFieldValue(item, FORWARD_FIELD_KEYS);

                    if (filterKey === 'calls') {
                        return call === 'yes' || call === 'waiting';
                    }
                    if (filterKey === 'meetings') {
                        return contact === 'yes' || contact === 'waiting';
                    }
                    if (filterKey === 'forward') {
                        return forward === 'yes';
                    }
                    return true;
                }

                function matchesActiveFilters(item) {
                    activeFilters = normalizeQuickFilters(activeFilters);
                    if (activeFilters.all) return true;
                    const activeFilterKey = Object.keys(activeFilters).find((filterKey) => filterKey !== 'all' && activeFilters[filterKey]) || 'all';
                    return matchesQuickFilter(item, activeFilterKey);
                }

                function scheduleTableRender(keyword = getCurrentSearchKeyword()) {
                    const normalizedKeyword = String(keyword || '').toLowerCase();
                    if (pendingTableRenderFrame) {
                        window.cancelAnimationFrame(pendingTableRenderFrame);
                    }
                    pendingTableRenderFrame = window.requestAnimationFrame(() => {
                        pendingTableRenderFrame = 0;
                        renderTable(normalizedKeyword);
                    });
                }

                function renderTable(keyword = "") {
                    const tbody = document.getElementById('table-body');
                    if (!tbody) return;
                    const data = activeType === 'vc' ? rawData.vc : rawData.fo;
                    const normalizedKeyword = String(keyword || '').toLowerCase();
                    const rows = [];

                    data.forEach(item => {
                        const name = item.__dashboardName || "";
                        if (!name) return;
                        if (normalizedKeyword && !(item.__dashboardSearchIndex || '').includes(normalizedKeyword)) return;
                        if (!matchesCompositionSelection(item)) return;
                        if (!matchesActiveFilters(item)) return;

                        const size = escapeHtml(item.__dashboardSize || '–');
                        const email = escapeHtml(item.__dashboardEmail || '–');

                        if (activeType === 'vc') {
                            rows.push(`
                        <tr>
                            <td><b>${escapeHtml(name)}</b></td>
                            <td><span class="size-tag">${size}</span></td>
                            <td><a href="mailto:${email}" style="color:var(--accent); text-decoration:none;">${email}</a></td>
                        </tr>`);
                            return;
                        }

                        const investorType = item.__dashboardType || 'Unknown';
                        const typeColor = typeColors[investorType] || '#94a3b8';
                        rows.push(`
                        <tr>
                            <td><b>${escapeHtml(name)}</b></td>
                            <td><span style="color:${typeColor}; font-weight:600;">${escapeHtml(investorType)}</span></td>
                            <td><span class="size-tag">${size}</span></td>
                        </tr>`);
                    });

                    if (!rows.length) {
                        const emptyStateColspan = activeType === 'vc' ? 3 : 3;
                        tbody.innerHTML = `
                        <tr class="dashboard-empty-row">
                            <td colspan="${emptyStateColspan}">No investors match the current filters.</td>
                        </tr>`;
                        return;
                    }

                    tbody.innerHTML = rows.join('');
                }

                function filterData() {
                    if (pendingSearchFilterTimer) {
                        window.clearTimeout(pendingSearchFilterTimer);
                    }
                    pendingSearchFilterTimer = window.setTimeout(() => {
                        pendingSearchFilterTimer = 0;
                        scheduleTableRender(getCurrentSearchKeyword());
                    }, 90);
                }

                function getCapacitorPlugin(name) {
                    const capacitor = window.Capacitor;
                    if (!capacitor || !name) return null;
                    if (capacitor.Plugins && capacitor.Plugins[name]) return capacitor.Plugins[name];
                    if (typeof capacitor.registerPlugin === 'function') {
                        try {
                            return capacitor.registerPlugin(name);
                        } catch (error) {
                            console.warn(`[Dashboard] Failed to register Capacitor plugin: ${name}`, error);
                        }
                    }
                    return null;
                }

                function isNativeMobileApp() {
                    return Boolean(
                        window.Capacitor &&
                        typeof window.Capacitor.isNativePlatform === 'function' &&
                        window.Capacitor.isNativePlatform()
                    );
                }

                function sanitizeExportFileName(value) {
                    return String(value || 'snapshot')
                        .trim()
                        .replace(/[^a-z0-9._-]+/gi, '-')
                        .replace(/-+/g, '-')
                        .replace(/^-|-$/g, '') || 'snapshot';
                }

                function getExportOptions(exportPreset = 'default') {
                    const includeInvestorDetailsInput = document.getElementById('export-include-investor-details');
                    if (exportPreset === 'names-no-emails') {
                        return {
                            includeInvestorNames: true,
                            includeInvestorEmails: false,
                            exportPreset,
                        };
                    }
                    const includeInvestorDetails = !includeInvestorDetailsInput || includeInvestorDetailsInput.checked;
                    return {
                        includeInvestorNames: includeInvestorDetails,
                        includeInvestorEmails: includeInvestorDetails,
                        exportPreset: includeInvestorDetails ? 'full' : 'anonymized',
                    };
                }

                function isSensitiveInvestorNameField(fieldName) {
                    const normalized = String(fieldName || '')
                        .trim()
                        .toLowerCase()
                        .replace(/[^a-z0-9]+/g, ' ')
                        .trim();
                    return [
                        'investor',
                        'investor name',
                        'name',
                        'contact name',
                    ].includes(normalized);
                }

                function isSensitiveInvestorEmailField(fieldName) {
                    const normalized = String(fieldName || '')
                        .trim()
                        .toLowerCase()
                        .replace(/[^a-z0-9]+/g, ' ')
                        .trim();
                    return [
                        'email',
                        'contact email',
                        'e mail',
                    ].includes(normalized);
                }

                function buildSnapshotRawData(options) {
                    const exportOptions = options || {};
                    const includeInvestorNames = exportOptions.includeInvestorNames !== false;
                    const includeInvestorEmails = exportOptions.includeInvestorEmails !== false;
                    const buildCollection = (rows) => (Array.isArray(rows) ? rows : []).map((row, index) => {
                        const next = Object.assign({}, row || {});
                        next.__snapshotInvestorLabel = `Investor ${index + 1}`;
                        if (!includeInvestorNames || !includeInvestorEmails) {
                            Object.keys(next).forEach((key) => {
                                if (!includeInvestorNames && isSensitiveInvestorNameField(key)) next[key] = '';
                                if (!includeInvestorEmails && isSensitiveInvestorEmailField(key)) next[key] = '';
                            });
                        }
                        if (!includeInvestorNames) {
                            next.__dashboardName = '';
                        }
                        if (!includeInvestorEmails) {
                            next.__dashboardEmail = '';
                        }
                        if (!includeInvestorNames || !includeInvestorEmails) {
                            next.__dashboardSearchIndex = '';
                        }
                        return next;
                    });

                    return {
                        vc: buildCollection(rawData && rawData.vc),
                        fo: buildCollection(rawData && rawData.fo),
                    };
                }

                function sanitizeCloneForExport(clone) {
                    if (!clone) return;
                    const header = clone.querySelector('header.dashboard-hero');
                    if (header) header.remove();

                    const nav = clone.querySelector('.sidebar-nav');
                    if (nav) nav.remove();

                    const layout = clone.querySelector('.layout');
                    if (layout) layout.style.gridTemplateColumns = '1fr';

                    const addPanel = clone.querySelector('#add-dashboard-panel');
                    if (addPanel) addPanel.remove();

                    const searchInput = clone.querySelector('#search');
                    if (searchInput) {
                        searchInput.value = '';
                        searchInput.setAttribute('value', '');
                    }

                    const tableHead = clone.querySelector('#table-head');
                    if (tableHead) tableHead.innerHTML = '';

                    const tableBody = clone.querySelector('#table-body');
                    if (tableBody) tableBody.innerHTML = '';
                }

                async function blobToBase64(blob) {
                    const buffer = await blob.arrayBuffer();
                    const bytes = new Uint8Array(buffer);
                    const chunkSize = 0x8000;
                    let binary = '';
                    for (let index = 0; index < bytes.length; index += chunkSize) {
                        const chunk = bytes.subarray(index, index + chunkSize);
                        binary += String.fromCharCode(...chunk);
                    }
                    return btoa(binary);
                }

                async function exportBlobForCurrentPlatform(blob, fileName) {
                    const SharePlugin = getCapacitorPlugin('Share');
                    const FilesystemPlugin = getCapacitorPlugin('Filesystem');

                    if (
                        isNativeMobileApp() &&
                        SharePlugin &&
                        FilesystemPlugin &&
                        typeof FilesystemPlugin.writeFile === 'function' &&
                        typeof SharePlugin.share === 'function'
                    ) {
                        const base64Data = await blobToBase64(blob);
                        const saved = await FilesystemPlugin.writeFile({
                            path: fileName,
                            data: base64Data,
                            directory: 'Cache',
                            recursive: true,
                        });

                        await SharePlugin.share({
                            title: 'Export HTML',
                            text: 'Investor dashboard snapshot',
                            url: saved.uri,
                            dialogTitle: 'Share investor dashboard snapshot',
                        });
                        return 'shared';
                    }

                    if (typeof File === 'function' && navigator.share) {
                        const shareFile = new File([blob], fileName, { type: 'text/html' });
                        if (!navigator.canShare || navigator.canShare({ files: [shareFile] })) {
                            await navigator.share({
                                title: 'Export HTML',
                                text: 'Investor dashboard snapshot',
                                files: [shareFile],
                            });
                            return 'shared';
                        }
                    }

                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.download = fileName;
                    a.href = url;
                    a.rel = 'noopener';
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    setTimeout(() => URL.revokeObjectURL(url), 1000);
                    return 'downloaded';
                }

                async function exportStandaloneHtml(options = {}) {
                    try {
                        if (!(currentDashboardMode === 'dashboard' && currentDashboard)) {
                            alert('Open a linked investor dashboard first, then export the live snapshot.');
                            return;
                        }
                        closeExportMenu();
                        const exportOptions = getExportOptions(options && options.exportPreset ? options.exportPreset : 'default');
                        const dashboardEl = document.getElementById('dashboard');
                        const shellEl = document.querySelector('.shell');
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
                        sanitizeCloneForExport(clone);

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
                            rawData: buildSnapshotRawData(exportOptions),
                            typeColors,
                            activeType,
                            activeFilters,
                            searchTerm: (exportOptions.includeInvestorNames || exportOptions.includeInvestorEmails)
                                ? (document.getElementById('search')?.value || '').toLowerCase()
                                : '',
                            exportOptions,
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
  const exportOptions = state.exportOptions || { includeInvestorNames: true, includeInvestorEmails: true };
  const includeInvestorNames = exportOptions.includeInvestorNames !== false;
  const includeInvestorEmails = exportOptions.includeInvestorEmails !== false;
  const CONTACT_FIELD_KEYS = ['Contact\\u00a0', 'Contact', 'Contact ', 'Contact/Call', 'Contact / Call'];
  const CALL_FIELD_KEYS = ['Call/Meeting', 'Call'];
  const MEETING_FIELD_KEYS = ['Meeting\\u00a0', 'Meeting with Company', 'Meeting'];
  const FORWARD_FIELD_KEYS = ['Reply', 'Replied', 'Moving Forward'];

  const searchInput = document.getElementById('search');
  if (searchInput && searchTerm) searchInput.value = searchTerm;

  function getFieldValue(item, keys) {
    for (const key of keys) {
      if (item && Object.prototype.hasOwnProperty.call(item, key)) {
        const value = String(item[key]).trim().toLowerCase();
        if (value) return value;
      }
    }
    return '';
  }

  function getStatusBadge(item) {
    const forward = getFieldValue(item, FORWARD_FIELD_KEYS);
    const meeting = getFieldValue(item, MEETING_FIELD_KEYS);
    const call = getFieldValue(item, [...CALL_FIELD_KEYS, ...CONTACT_FIELD_KEYS]);

    if (forward === 'yes') return '<span class="stage-pill badge-green">Replied / Mov. Forward</span>';
    if (forward === 'waiting' || meeting === 'waiting' || call === 'waiting') return '<span class="stage-pill badge-amber">Waiting / Ongoing</span>';
    if (forward === 'no') return '';
    if (meeting === 'yes') return '<span class="stage-pill badge-neutral">Contacted / Met</span>';
    if (call === 'yes') return '<span class="stage-pill badge-neutral">Contact Started</span>';
    return '<span class="stage-pill" style="color:var(--text-dim); opacity:0.5;">Target</span>';
  }

  function normalizeQuickFilters(filters) {
    const source = filters && typeof filters === 'object' ? filters : {};
    const next = { all: false, calls: false, meetings: false, forward: false };
    if (source.calls) next.calls = true;
    else if (source.meetings) next.meetings = true;
    else if (source.forward) next.forward = true;
    else next.all = true;
    return next;
  }

  function applyQuickFilterButtons() {
    Object.keys(activeFilters).forEach(filterKey => {
      const btn = document.getElementById(\`f-\${filterKey}\`);
      if (!btn) return;
      btn.classList.toggle('active', Boolean(activeFilters[filterKey]));
    });
  }

  function matchesQuickFilter(item, filterKey) {
    const contact = getFieldValue(item, CONTACT_FIELD_KEYS);
    const call = getFieldValue(item, CALL_FIELD_KEYS);
    const forward = getFieldValue(item, FORWARD_FIELD_KEYS);

    if (filterKey === 'calls') {
      return call === 'yes' || call === 'waiting';
    }
    if (filterKey === 'meetings') {
      return contact === 'yes' || contact === 'waiting';
    }
    if (filterKey === 'forward') {
      return forward === 'yes';
    }
    return true;
  }

  function renderTable(keyword = '') {
    const tbody = document.getElementById('table-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const data = activeType === 'vc' ? (rawData.vc || []) : (rawData.fo || []);

    data.forEach(item => {
      const name = includeInvestorNames
        ? (item['Investor'] || item['Investor Name'] || item['Name'] || '')
        : (item.__snapshotInvestorLabel || '');
      const size = item['Size of Investment'] || item['Investment Size'] || '–';
      const email = includeInvestorEmails ? (item['Email'] || '–') : '';

      if (keyword && !JSON.stringify(item).toLowerCase().includes(keyword)) return;
      if (!name) return;

      activeFilters = normalizeQuickFilters(activeFilters);
      if (!activeFilters.all) {
        const activeFilterKey = Object.keys(activeFilters).find(filterKey => filterKey !== 'all' && activeFilters[filterKey]) || 'all';
        if (!matchesQuickFilter(item, activeFilterKey)) return;
      }

      const statusBadge = getStatusBadge(item);
      const iType = item['Type'] || 'Unknown';
      const typeColor = typeColors[iType] || '#94a3b8';

      if (activeType === 'vc') {
        if (includeInvestorEmails) {
          tbody.innerHTML += \`
            <tr>
              <td><b>\${name}</b></td>
              <td><span class="size-tag">\${size}</span></td>
              <td><a href="mailto:\${email}" style="color:var(--accent); text-decoration:none;">\${email}</a></td>
            </tr>\`;
        } else {
          tbody.innerHTML += \`
            <tr>
              <td><b>\${name}</b></td>
              <td><span class="size-tag">\${size}</span></td>
            </tr>\`;
        }
      } else {
        tbody.innerHTML += \`
          <tr>
            <td><b>\${name}</b></td>
            <td><span style="color:\${typeColor}; font-weight:600;">\${iType}</span></td>
            <td><span class="size-tag">\${size}</span></td>
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
        header.innerHTML = includeInvestorNames
          ? '<tr><th>VC Funds Name</th><th>Investment Size</th><th>Contact Email</th></tr>'
          : '<tr><th>Investor</th><th>Investment Size</th></tr>';
        if (includeInvestorNames && !includeInvestorEmails) {
          header.innerHTML = '<tr><th>VC Funds Name</th><th>Investment Size</th></tr>';
        }
      } else {
        header.innerHTML = includeInvestorNames
          ? '<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th></tr>'
          : '<tr><th>Investor</th><th>Investor Type</th><th>Investment Size</th></tr>';
      }
    }
    renderTable(searchInput ? searchInput.value.toLowerCase() : '');
  }

  function toggleFilter(filter) {
    const normalizedFilter = ['all', 'calls', 'meetings', 'forward'].includes(filter) ? filter : 'all';
    activeFilters = normalizeQuickFilters(activeFilters);
    activeFilters = normalizedFilter === 'all'
      ? normalizeQuickFilters({ all: true })
      : normalizeQuickFilters({ [normalizedFilter]: !activeFilters[normalizedFilter] });
    applyQuickFilterButtons();
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
    activeFilters = normalizeQuickFilters(activeFilters);
    applyQuickFilterButtons();

    // Seed header and table content
    switchTab(activeType);
    if (searchTerm) renderTable(searchTerm);
  });
})();
</script>
</body>
</html>`;

                        const blob = new Blob([snapshotHtml], { type: 'text/html' });
                        const dashId = (currentDashboard && currentDashboard.id) || 'snapshot';
                        const exportSuffix = exportOptions.includeInvestorNames
                            ? (exportOptions.includeInvestorEmails ? '' : '-names-no-emails')
                            : '-anonymized';
                        const fileName = `investor-dashboard-${sanitizeExportFileName(dashId)}${exportSuffix}.html`;
                        await exportBlobForCurrentPlatform(blob, fileName);
                    } catch (err) {
                        if (err && err.name === 'AbortError') {
                            return;
                        }
                        console.error('Export failed', err);
                        alert('Could not export a snapshot on this device. Please try again.');
                    }
                }
