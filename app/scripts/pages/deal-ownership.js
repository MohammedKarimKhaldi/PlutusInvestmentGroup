(function initDealOwnership() {
    const AppCore = window.AppCore;
    const DASHBOARD_ID = "deal-ownership";
    const STAGE_LABELS = {
        prospect: "Prospect",
        signing: "Signing",
        onboarding: "Onboarding",
        "contacting investors": "Contacting investors",
    };
    
    let rawData = [];
    let dealsSource = [];
    let dealsMap = new Map();
    let ownershipContext = null;
    let ownershipFilters = {
        keyword: "",
        view: "summary",
        link: "all",
        coverage: "all",
        stage: "all",
        sector: "all",
        location: "all",
        fundingStage: "all",
        revenue: "all",
        thematic: "all",
    };
    let charts = {};
    const COLOR_PALETTE = ['#818cf8', '#34d399', '#fbbf24', '#a78bfa', '#f472b6', '#22d3ee', '#fb7185'];

    function normalizeValue(value) {
        if (AppCore && typeof AppCore.normalizeValue === 'function') {
            return AppCore.normalizeValue(value);
        }
        return String(value || '').trim().toLowerCase();
    }

    function normalizeDealLifecycleStatus(value) {
        const normalized = normalizeValue(value);
        if (normalized === 'finished') return 'finished';
        if (normalized === 'closed') return 'closed';
        return 'active';
    }

    function isActiveDeal(deal) {
        return normalizeDealLifecycleStatus(deal && (deal.lifecycleStatus || deal.dealStatus)) === 'active';
    }

    window.addEventListener('load', () => {
        setupUiBindings();
        initializeDashboard();
    });

    function setupUiBindings() {
        const searchInput = document.getElementById('search');
        const retryBtn = document.getElementById('btn-retry-sync');
        const viewFilter = document.getElementById('ownership-view-filter');
        const linkFilter = document.getElementById('ownership-link-filter');
        const coverageFilter = document.getElementById('ownership-coverage-filter');
        const stageFilter = document.getElementById('ownership-stage-filter');
        const resetBtn = document.getElementById('ownership-reset-filters');
        const visualFilters = document.getElementById('ownership-visual-filters');

        if (searchInput) {
            searchInput.addEventListener('input', () => {
                syncOwnershipFiltersFromUi();
                renderDashboardState();
            });
        }
        [viewFilter, linkFilter, coverageFilter, stageFilter].forEach((control) => {
            if (!control) return;
            control.addEventListener('change', () => {
                syncOwnershipFiltersFromUi();
                renderDashboardState();
            });
        });
        if (resetBtn) {
            resetBtn.addEventListener('click', () => {
                ownershipFilters = {
                    keyword: "",
                    view: "summary",
                    link: "all",
                    coverage: "all",
                    stage: "all",
                    sector: "all",
                    location: "all",
                    fundingStage: "all",
                    revenue: "all",
                    thematic: "all",
                };
                applyOwnershipFiltersToUi();
                renderDashboardState();
            });
        }
        if (visualFilters) {
            visualFilters.addEventListener('click', (event) => {
                const chip = event.target.closest('[data-ownership-visual-filter]');
                if (!chip) return;
                const filterKey = String(chip.getAttribute('data-ownership-visual-filter') || '').trim();
                const nextValue = String(chip.getAttribute('data-ownership-visual-value') || 'all').trim();
                if (!filterKey || !(filterKey in ownershipFilters)) return;
                ownershipFilters[filterKey] = ownershipFilters[filterKey] === nextValue && nextValue !== 'all' ? 'all' : nextValue;
                renderDashboardState();
            });
        }
        if (retryBtn) {
            retryBtn.addEventListener('click', initializeDashboard);
        }

        window.addEventListener('appcore:deals-updated', () => {
            buildDealsMap();
            rebuildOwnershipContext();
            renderDashboardState();
        });

        window.addEventListener('appcore:dashboard-config-updated', () => {
            rebuildOwnershipContext();
            renderDashboardState();
        });
    }

    function syncOwnershipFiltersFromUi() {
        ownershipFilters = {
            keyword: String(document.getElementById('search') && document.getElementById('search').value || '').trim().toLowerCase(),
            view: String(document.getElementById('ownership-view-filter') && document.getElementById('ownership-view-filter').value || 'summary').trim().toLowerCase() || 'summary',
            link: String(document.getElementById('ownership-link-filter') && document.getElementById('ownership-link-filter').value || 'all').trim().toLowerCase() || 'all',
            coverage: String(document.getElementById('ownership-coverage-filter') && document.getElementById('ownership-coverage-filter').value || 'all').trim().toLowerCase() || 'all',
            stage: String(document.getElementById('ownership-stage-filter') && document.getElementById('ownership-stage-filter').value || 'all').trim().toLowerCase() || 'all',
            sector: ownershipFilters.sector || 'all',
            location: ownershipFilters.location || 'all',
            fundingStage: ownershipFilters.fundingStage || 'all',
            revenue: ownershipFilters.revenue || 'all',
            thematic: ownershipFilters.thematic || 'all',
        };
    }

    function applyOwnershipFiltersToUi() {
        const searchInput = document.getElementById('search');
        const viewFilter = document.getElementById('ownership-view-filter');
        const linkFilter = document.getElementById('ownership-link-filter');
        const coverageFilter = document.getElementById('ownership-coverage-filter');
        const stageFilter = document.getElementById('ownership-stage-filter');

        if (searchInput) searchInput.value = ownershipFilters.keyword || '';
        if (viewFilter) viewFilter.value = ownershipFilters.view || 'summary';
        if (linkFilter) linkFilter.value = ownershipFilters.link || 'all';
        if (coverageFilter) coverageFilter.value = ownershipFilters.coverage || 'all';
        if (stageFilter) stageFilter.value = ownershipFilters.stage || 'all';
    }

    function rebuildOwnershipContext() {
        ownershipContext = rawData.length ? buildOwnershipContext() : null;
        return ownershipContext;
    }

    function renderDashboardState() {
        applyOwnershipFiltersToUi();
        if (!ownershipContext) return;
        renderOwnershipVisualFilters(ownershipContext);
        const scopedContext = buildScopedOwnershipContext(ownershipContext);
        renderKPIs(scopedContext);
        renderCharts(scopedContext);
        renderWatchlist(scopedContext);
        renderTable(ownershipFilters.keyword, scopedContext);
    }

    async function initializeDashboard() {
        showLoader("Syncing Staffing Data...");
        try {
            if (AppCore && typeof AppCore.refreshDashboardConfigFromShareDrive === "function") {
                await AppCore.refreshDashboardConfigFromShareDrive();
            }
            buildDealsMap();
            const config = AppCore.getDashboardConfig();
            const dash = config.dashboards.find(d => d.id === DASHBOARD_ID);
            
            if (!dash || !dash.excelUrl) {
                throw new Error("Deal ownership configuration not found in config.json");
            }

            console.log('[DealOwnership] Selected dashboard:', dash);
            
            // Resolve download URL
            const downloadUrl = await resolveUrl(dash.excelUrl);
            const wb = await fetchWorkbook(downloadUrl);
            
            processWorkbook(wb, dash.sheets);
            showDashboard();
        } catch (err) {
            console.error('[DealOwnership] Init failed:', err);
            showError(err.message);
        }
    }

    async function resolveUrl(url) {
        const normalized = String(url || "").toLowerCase();
        const isSharePointLink = normalized.includes("sharepoint.com") || normalized.includes("1drv.ms");
        if (isSharePointLink && AppCore && typeof AppCore.resolveShareDriveDownloadUrl === "function") {
            try {
                const resolved = await AppCore.resolveShareDriveDownloadUrl(url);
                if (resolved) return resolved;
            } catch (error) {
                console.warn("[DealOwnership] ShareDrive URL resolution failed, using original URL", error);
            }
        }
        return url;
    }

    function buildDownloadHintUrl(url) {
        const raw = String(url || "").trim();
        if (!raw) return raw;
        if (/([?&])download=1(?:&|$)/i.test(raw)) return raw;
        return raw.includes("?") ? `${raw}&download=1` : `${raw}?download=1`;
    }

    function looksLikeHtmlBuffer(buffer) {
        try {
            const bytes = new Uint8Array(buffer || new ArrayBuffer(0));
            const sample = bytes.slice(0, 512);
            const text = new TextDecoder("utf-8").decode(sample).trim().toLowerCase();
            return text.startsWith("<!doctype html") || text.startsWith("<html") || text.includes("<head") || text.includes("<body");
        } catch {
            return false;
        }
    }

    async function downloadWorkbookBuffer(url) {
        return AppCore && typeof AppCore.downloadBinary === "function"
            ? AppCore.downloadBinary(url, { cache: "no-store" })
            : fetch(url).then((response) => {
                if (!response.ok) throw new Error("Failed to download staffing file");
                return response.arrayBuffer();
            });
    }

    async function fetchWorkbook(url) {
        const attempts = [String(url || "").trim()].filter(Boolean);
        const hintedUrl = buildDownloadHintUrl(url);
        if (hintedUrl && hintedUrl !== attempts[0]) attempts.push(hintedUrl);

        let lastError = null;
        for (let i = 0; i < attempts.length; i += 1) {
            const attemptUrl = attempts[i];
            try {
                const buffer = await downloadWorkbookBuffer(attemptUrl);
                if (looksLikeHtmlBuffer(buffer)) {
                    throw new Error("Downloaded content is HTML (SharePoint sign-in/permissions page) instead of an Excel file.");
                }
                return XLSX.read(buffer, { type: 'array' });
            } catch (error) {
                lastError = error;
            }
        }

        if (lastError && lastError.message && lastError.message.includes("Downloaded content is HTML")) {
            throw new Error("Failed to download DealOwnership workbook: SharePoint returned an HTML page. Reconnect Sharedrive and confirm the excelUrl points to a file you can access.");
        }
        throw lastError || new Error("Failed to download DealOwnership workbook.");
    }

    function sanitizeWorkbookRows(rows) {
        if (!Array.isArray(rows)) return [];
        return rows.map((row) => {
            const cleaned = {};
            Object.entries(row || {}).forEach(([key, value]) => {
                const header = String(key || '').trim();
                if (!header || /^__EMPTY/i.test(header)) return;
                cleaned[header] = value;
            });
            return cleaned;
        });
    }

    function processWorkbook(wb, sheetsConfig) {
        const sheetName = sheetsConfig.funds || wb.SheetNames[0]; // Reuse 'funds' mapping for generic staffing sheet or take first
        const sheet = wb.Sheets[sheetName] || wb.Sheets[wb.SheetNames[0]];
        
        if (!sheet) throw new Error("Staffing sheet not found");

        // Simple row detection
        rawData = sanitizeWorkbookRows(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
        console.log('[DealOwnership] Raw data rows:', rawData.length);
        
        if (rawData.length === 0) return;

        rebuildOwnershipContext();
        renderDashboardState();
    }

    function buildDealsMap() {
        dealsMap = new Map();
        const source = AppCore && typeof AppCore.loadDealsData === "function"
            ? AppCore.loadDealsData()
            : (Array.isArray(window.DEALS) ? window.DEALS : []);
        dealsSource = Array.isArray(source) ? source : [];

        const normalize = (v) => String(v || "").trim().toLowerCase();
        dealsSource.forEach(d => {
            const keys = [d.name, d.company, d.id, d.fundraisingDashboardId];
            keys.filter(Boolean).forEach(k => dealsMap.set(normalize(k), d));
        });
    }

    function findDeal(reference) {
        if (AppCore && typeof AppCore.findDealByReference === "function") {
            return AppCore.findDealByReference(dealsSource, reference);
        }

        const references = Array.isArray(reference) ? reference : [reference];
        const norm = (v) => String(v || "").trim().toLowerCase();
        for (const entry of references) {
            const key = norm(entry);
            if (key && dealsMap.has(key)) return dealsMap.get(key);
        }
        return null;
    }

    function buildPageUrl(pageId, params) {
        if (AppCore && typeof AppCore.getPageUrl === "function") {
            return AppCore.getPageUrl(pageId, params);
        }
        if (window.PlutusAppConfig && typeof window.PlutusAppConfig.buildPageHref === "function") {
            return window.PlutusAppConfig.buildPageHref(pageId, params);
        }
        const query = new URLSearchParams(params || {});
        const queryString = query.toString();
        return queryString ? `${pageId}.html?${queryString}` : `${pageId}.html`;
    }

    function findDealForRow(row, headers) {
        const references = headers
            .filter((header) => /deal|company|project/i.test(header))
            .map((header) => row[header])
            .filter(Boolean);

        return findDeal(references);
    }

    function escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function buildLinksCell(match) {
        if (!match || !match.id) return '–';

        const overviewHref = buildPageUrl("deal-details", { id: match.id });
        const config = AppCore && typeof AppCore.getDashboardConfig === "function"
            ? AppCore.getDashboardConfig()
            : (window.DASHBOARD_CONFIG || {});
        let dashboard = AppCore && typeof AppCore.getDashboardForDeal === "function"
            ? AppCore.getDashboardForDeal(match, config)
            : null;
        if (!dashboard) {
            const dashboardId = String(match.fundraisingDashboardId || "").trim().toLowerCase();
            const dashboards = Array.isArray(config.dashboards) ? config.dashboards : [];
            dashboard = dashboards.find((entry) => String(entry.id || "").trim().toLowerCase() === dashboardId) || null;
        }

        let html = `<a class="action-link" href="${overviewHref}">Deal overview</a>`;
        if (dashboard && dashboard.id) {
            const dashHref = buildPageUrl("investor-dashboard", { dashboard: dashboard.id });
            html += ` <span style="color: var(--text-dim);">·</span> <a class="action-link" href="${dashHref}">Dashboard</a>`;
        }
        return html;
    }

    function getRowValueByHeaderPattern(row, patterns) {
        const headers = Object.keys(row || {});
        const matchHeader = headers.find((header) => patterns.some((pattern) => pattern.test(header)));
        return matchHeader ? String(row[matchHeader] || '').trim() : '';
    }

    function getPrimaryDealLabel(row, headers) {
        const prioritizedHeader = headers.find((header) => /deal/i.test(header))
            || headers.find((header) => /company/i.test(header))
            || headers.find((header) => /project/i.test(header));
        return prioritizedHeader ? String(row[prioritizedHeader] || '').trim() : '';
    }

    function getStaffLabel(row) {
        return getRowValueByHeaderPattern(row, [/^staff$/i, /^name$/i, /^lead$/i, /staff/i, /owner/i]) || 'Unassigned';
    }

    function getRoleLabel(row) {
        return getRowValueByHeaderPattern(row, [/^role$/i, /title/i, /position/i, /function/i, /team/i]);
    }

    function getSubOwners(deal) {
        const source = deal && deal.subOwners;
        const values = Array.isArray(source)
            ? source
            : typeof source === 'string'
                ? source.split(/[\n,;]+/)
                : [];
        return Array.from(new Set(values.map((entry) => String(entry || '').trim()).filter(Boolean)));
    }

    function getDealKeywords(deal) {
        const source = deal && deal.keywords;
        const values = Array.isArray(source)
            ? source
            : typeof source === 'string'
                ? source.split(/[\n,;]+/)
                : [];
        return Array.from(new Set(values.map((entry) => String(entry || '').trim()).filter(Boolean)));
    }

    function getDealSectors(deal) {
        const source = deal && (
            Array.isArray(deal.sectors)
                ? deal.sectors
                : typeof deal.sectors === 'string'
                    ? deal.sectors.split(/[\n,;]+/)
                    : typeof deal.sector === 'string'
                        ? deal.sector.split(/[\n,;]+/)
                        : []
        );
        const normalized = Array.from(new Set(
            (Array.isArray(source) ? source : [])
                .map((entry) => String(entry || '').trim())
                .filter(Boolean)
        ));
        if (deal && typeof deal === 'object') {
            deal.sectors = normalized;
            deal.sector = normalized.join(', ');
        }
        return normalized;
    }

    function getDealProfileValue(deal, key) {
        if (key === 'sector') {
            return getDealSectors(deal).join(', ');
        }
        return String(deal && deal[key] || '').trim();
    }

    function getDealProfileSummary(deal) {
        return [
            getDealProfileValue(deal, 'sector'),
            getDealProfileValue(deal, 'location'),
            getDealProfileValue(deal, 'fundingStage'),
            getDealProfileValue(deal, 'revenue'),
        ].filter(Boolean);
    }

    function buildDealProfileLabel(deal) {
        const profileParts = getDealProfileSummary(deal);
        if (!profileParts.length) return '';
        return profileParts.join(' · ');
    }

    function addSubOwnersToSummary(summary) {
        if (!summary || !summary.match || !summary.staff || !summary.roles) return;
        const subOwners = getSubOwners(summary.match);
        if (!subOwners.length) return;

        const existingKeys = new Set(Array.from(summary.staff).map((entry) => normalizeValue(entry)));
        let addedCount = 0;

        subOwners.forEach((owner) => {
            const key = normalizeValue(owner);
            if (!key || existingKeys.has(key)) return;
            existingKeys.add(key);
            summary.staff.add(owner);
            addedCount += 1;
        });

        if (addedCount > 0) {
            summary.roles.set('Sub owner', (summary.roles.get('Sub owner') || 0) + addedCount);
        }
    }

    function buildOwnerRosterLabel(deal) {
        if (!deal) return 'Owners: –';
        const primaryOwners = [
            String(deal.seniorOwner || deal.owner || '').trim(),
            String(deal.juniorOwner || '').trim(),
        ].filter(Boolean);
        const subOwners = getSubOwners(deal);
        const parts = [];
        parts.push(`Owners: ${primaryOwners.join(' / ') || '–'}`);
        if (subOwners.length) {
            parts.push(`Sub owners: ${subOwners.join(', ')}`);
        }
        return parts.join(' · ');
    }

    function getDealOwners(deal) {
        const primaryOwners = [
            String(deal && (deal.seniorOwner || deal.owner) || '').trim(),
            String(deal && deal.juniorOwner || '').trim(),
            ...getSubOwners(deal),
        ].filter(Boolean);
        return Array.from(new Set(primaryOwners));
    }

    function formatAmountCell(amount, currency) {
        if (amount == null || amount === '') {
            return `<span class="amount-unknown">??</span>`;
        }
        if (typeof amount === 'string') {
            return escapeHtml(amount);
        }
        if (typeof amount === 'number' && !Number.isNaN(amount)) {
            try {
                const formatted = new Intl.NumberFormat('en-US', {
                    style: 'currency',
                    currency: currency || 'USD',
                    maximumFractionDigits: 0,
                }).format(amount);
                return escapeHtml(formatted);
            } catch {
                return escapeHtml(`${currency || 'USD'} ${Number(amount).toLocaleString()}`);
            }
        }
        return `<span class="amount-unknown">??</span>`;
    }

    function normalizeStageKey(stage) {
        return String(stage || '').trim().toLowerCase();
    }

    function getStageLabel(stage) {
        const normalized = normalizeStageKey(stage);
        return STAGE_LABELS[normalized] || (normalized ? normalized.charAt(0).toUpperCase() + normalized.slice(1) : 'No stage');
    }

    function getStageBadge(stage) {
        const normalized = normalizeStageKey(stage);
        const label = getStageLabel(stage);
        const className = normalized === 'prospect'
            ? 'stage-prospect'
            : normalized === 'signing'
                ? 'stage-signing'
                : normalized === 'onboarding'
                    ? 'stage-onboarding'
                    : normalized === 'contacting investors'
                        ? 'stage-contacting'
                        : 'stage-prospect';
        return `<span class="stage-badge"><span class="stage-dot ${className}"></span>${escapeHtml(label)}</span>`;
    }

    function getSummaryDisplayLabel(entry) {
        if (!entry) return 'Staffing entry';
        if (entry.match) return entry.match.name || entry.match.company || entry.match.id || entry.label || 'Deal';
        return entry.label || 'Staffing entry';
    }

    function buildCoverageSignal(entry) {
        const staffCount = entry.staffCount || 0;

        if (entry.linkStatus === 'unlinked') {
            return {
                status: 'unlinked',
                bucket: 'risk',
                label: 'Unlinked',
                detail: 'No matching deal overview record for this staffing entry yet.',
                actionTitle: 'Link or create the deal',
                actionDetail: 'Rename the staffing line or create a deal overview record so this work is trackable.',
            };
        }

        if (entry.linkStatus === 'missing-staffing') {
            return {
                status: 'missing-staffing',
                bucket: 'risk',
                label: 'Needs staffing',
                detail: 'This pipeline deal has no staffing rows linked yet.',
                actionTitle: 'Assign coverage',
                actionDetail: 'Add staffing coverage before the next pipeline milestone.',
            };
        }

        return {
            status: 'healthy',
            bucket: 'healthy',
            label: 'Covered',
            detail: `${staffCount} team member${staffCount === 1 ? '' : 's'} covering ${entry.assignments || 0} assignment${entry.assignments === 1 ? '' : 's'}.`,
            actionTitle: 'Coverage looks healthy',
            actionDetail: 'No immediate ownership action is flagged from staffing data.',
        };
    }

    function buildOwnershipContext() {
        const headers = rawData.length ? Object.keys(rawData[0]) : [];
        const summaries = new Map();
        const linkedDealIds = new Set();
        let unlinkedRows = 0;

        const enrichedRows = rawData.map((row) => {
            const match = findDealForRow(row, headers);
            if (match && !isActiveDeal(match)) {
                return null;
            }
            const primaryDealLabel = getPrimaryDealLabel(row, headers) || 'Unlabelled staffing line';
            const summaryKey = match && match.id
                ? `deal:${String(match.id)}`
                : `unlinked:${String(primaryDealLabel || '').trim().toLowerCase()}`;

            if (!summaries.has(summaryKey)) {
                summaries.set(summaryKey, {
                    key: summaryKey,
                    assignments: 0,
                    staff: new Set(),
                    roles: new Map(),
                    label: primaryDealLabel,
                    match,
                    isPipelineOnly: false,
                    rows: [],
                });
            }

            const summary = summaries.get(summaryKey);
            summary.assignments += 1;
            const staffLabel = getStaffLabel(row);
            const roleLabel = getRoleLabel(row);
            if (staffLabel) summary.staff.add(staffLabel);
            if (roleLabel) {
                summary.roles.set(roleLabel, (summary.roles.get(roleLabel) || 0) + 1);
            }
            summary.rows.push(row);

            if (match && match.id) linkedDealIds.add(String(match.id));
            if (!match) unlinkedRows += 1;

            return {
                row,
                match,
                summaryKey,
                primaryDealLabel,
            };
        }).filter(Boolean);

        dealsSource.filter((deal) => isActiveDeal(deal)).forEach((deal) => {
            const dealId = String(deal && deal.id || '').trim();
            if (!dealId || linkedDealIds.has(dealId)) return;
            const summaryKey = `deal:${dealId}`;
            if (summaries.has(summaryKey)) return;

            summaries.set(summaryKey, {
                key: summaryKey,
                assignments: 0,
                staff: new Set(),
                roles: new Map(),
                label: deal.name || deal.company || deal.id || 'Deal',
                match: deal,
                isPipelineOnly: true,
                rows: [],
            });
        });

        const summaryEntries = Array.from(summaries.values())
            .map((summary) => {
                addSubOwnersToSummary(summary);
                const staffNames = Array.from(summary.staff || []).sort((a, b) => a.localeCompare(b));
                const hasCoveragePeople = staffNames.length > 0;
                const linkStatus = !summary.match
                    ? 'unlinked'
                    : !hasCoveragePeople
                        ? 'missing-staffing'
                        : 'linked';
                const stageKey = normalizeStageKey(summary.match && summary.match.stage);
                const entry = {
                    ...summary,
                    staffNames,
                    staffCount: staffNames.length,
                    roleSummary: getTopRoleSummary(summary, 3),
                    linkStatus,
                    stageKey,
                    stageLabel: getStageLabel(stageKey),
                    sectors: getDealSectors(summary.match),
                    sector: getDealProfileValue(summary.match, 'sector'),
                    location: getDealProfileValue(summary.match, 'location'),
                    fundingStage: getDealProfileValue(summary.match, 'fundingStage'),
                    revenue: getDealProfileValue(summary.match, 'revenue'),
                    keywords: getDealKeywords(summary.match),
                };
                const signal = buildCoverageSignal(entry);
                entry.coverageStatus = signal.status;
                entry.coverageBucket = signal.bucket;
                entry.coverageLabel = signal.label;
                entry.coverageDetail = signal.detail;
                entry.actionTitle = signal.actionTitle;
                entry.actionDetail = signal.actionDetail;
                return entry;
            })
            .sort((a, b) => {
                const riskOrder = a.coverageBucket === b.coverageBucket ? 0 : a.coverageBucket === 'risk' ? -1 : 1;
                if (riskOrder) return riskOrder;
                const aAssignments = Number(a.assignments || 0);
                const bAssignments = Number(b.assignments || 0);
                if (aAssignments !== bAssignments) return bAssignments - aAssignments;
                return getSummaryDisplayLabel(a).localeCompare(getSummaryDisplayLabel(b));
            });

        const summaryEntriesByKey = new Map(summaryEntries.map((entry) => [entry.key, entry]));

        return {
            headers,
            enrichedRows,
            summaries,
            summaryEntries,
            summaryEntriesByKey,
            linkedDealCount: linkedDealIds.size,
            unlinkedRows,
        };
    }

    function getTopRoleSummary(summary, limit = 2) {
        if (!summary || !summary.roles || !summary.roles.size) return 'Roles not tagged';
        return Array.from(summary.roles.entries())
            .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
            .slice(0, limit)
            .map(([role, count]) => count > 1 ? `${role} x${count}` : role)
            .join(', ');
    }

    function buildLinkedDealCell(match, summary, fallbackLabel) {
        if (!match || !match.id) {
            return `
                <div class="ownership-stack">
                    <span class="ownership-badge ownership-badge-warn">Unlinked</span>
                    <span class="ownership-primary">${escapeHtml(fallbackLabel || summary.label || 'Staffing entry')}</span>
                    <span class="ownership-muted">No matching deal found in deals overview yet.</span>
                </div>
            `;
        }

        return `
            <div class="ownership-stack">
                <span class="ownership-badge ownership-badge-good">Linked</span>
                <a class="action-link ownership-primary-link" href="${buildPageUrl("deal-details", { id: match.id })}">${escapeHtml(match.name || match.company || match.id)}</a>
                <span class="ownership-muted">${escapeHtml(match.company || 'No company')} · ${escapeHtml(match.id || '')}</span>
            </div>
        `;
    }

    function buildPipelineCell(match) {
        if (!match) {
            return `<div class="ownership-stack"><span class="ownership-muted">Create or rename a deal overview record to link this staffing line.</span></div>`;
        }

        const profileLabel = buildDealProfileLabel(match);
        const keywordLabel = getDealKeywords(match).slice(0, 4).join(', ');

        return `
            <div class="ownership-stack">
                ${getStageBadge(match.stage)}
                <span class="ownership-muted">${escapeHtml(buildOwnerRosterLabel(match))}</span>
                ${profileLabel ? `<span class="ownership-muted">${escapeHtml(profileLabel)}</span>` : ''}
                ${keywordLabel ? `<span class="ownership-muted">Thematics: ${escapeHtml(keywordLabel)}</span>` : ''}
                <span class="ownership-muted">Target: ${formatAmountCell(match.targetAmount, match.currency)} · Raised: ${formatAmountCell(match.raisedAmount, match.currency)}</span>
            </div>
        `;
    }

    function buildStaffingSnapshotCell(summary) {
        if (!summary) return '–';
        const teamSize = summary.staff ? summary.staff.size : 0;
        const assignmentCount = summary.assignments || 0;
        const roleSummary = getTopRoleSummary(summary);
        return `
            <div class="ownership-stack">
                <span class="ownership-primary">${teamSize} team member${teamSize === 1 ? '' : 's'} · ${assignmentCount} assignment${assignmentCount === 1 ? '' : 's'}</span>
                <span class="ownership-muted">${escapeHtml(roleSummary)}</span>
            </div>
        `;
    }

    function renderKPIs(context = ownershipContext) {
        const container = document.getElementById('staffing-kpis');
        if (!container || !context) return;

        const enrichedRows = Array.isArray(context.enrichedRows) ? context.enrichedRows : [];
        const totalRows = enrichedRows.length;
        const uniqueStaff = new Set(enrichedRows.map(({ row }) => getStaffLabel(row)).filter(Boolean)).size;
        const coveredDeals = context.summaryEntries.filter((entry) => entry.linkStatus === 'linked').length;
        const missingStaffingCount = context.summaryEntries.filter((entry) => entry.linkStatus === 'missing-staffing').length;
        const atRiskCount = context.summaryEntries.filter((entry) => entry.linkStatus === 'linked' && entry.coverageBucket === 'risk').length;
        const unlinkedClusters = context.summaryEntries.filter((entry) => entry.linkStatus === 'unlinked').length;

        container.innerHTML = `
            <div class="kpi-card">
                <div class="kpi-label">Active Staff</div>
                <div class="kpi-value">${uniqueStaff}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Total Assignments</div>
                <div class="kpi-value">${totalRows}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Covered Deals</div>
                <div class="kpi-value">${coveredDeals}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Needs Staffing</div>
                <div class="kpi-value" style="color: ${missingStaffingCount ? 'var(--warning)' : 'var(--success)'};">${missingStaffingCount}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">At-Risk Coverage</div>
                <div class="kpi-value" style="color: ${atRiskCount ? 'var(--warning)' : 'var(--success)'};">${atRiskCount}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Unlinked Clusters</div>
                <div class="kpi-value" style="color: ${unlinkedClusters ? 'var(--warning)' : 'var(--success)'};">${unlinkedClusters}</div>
            </div>
        `;
    }

    function renderCharts(context = ownershipContext) {
        if (!context) return;
        renderAllocationChart(context);
        renderDealDistChart(context);
    }

    function renderAllocationChart(context = ownershipContext) {
        const ctx = document.getElementById('staffChart').getContext('2d');
        if (charts.staff) charts.staff.destroy();

        const staffMap = {};
        (Array.isArray(context && context.enrichedRows) ? context.enrichedRows : []).forEach(({ row }) => {
            const name = getStaffLabel(row) || "Unknown";
            staffMap[name] = (staffMap[name] || 0) + 1;
        });

        const labels = Object.keys(staffMap);
        const data = Object.values(staffMap);

        charts.staff = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'Assignments',
                    data,
                    backgroundColor: COLOR_PALETTE[0],
                    borderRadius: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
                    x: { grid: { display: false } }
                }
            }
        });
    }

    function renderDealDistChart(context = ownershipContext) {
        const ctx = document.getElementById('dealChart').getContext('2d');
        if (charts.deal) charts.deal.destroy();

        const labels = [];
        const data = [];
        const colors = [];
        const buckets = [
            { key: 'healthy', label: 'Covered', color: '#34d399' },
            { key: 'missing-staffing', label: 'Needs staffing', color: '#f97316' },
            { key: 'unlinked', label: 'Unlinked', color: '#94a3b8' },
        ];

        buckets.forEach((bucket) => {
            const count = context.summaryEntries.filter((entry) => entry.coverageStatus === bucket.key).length;
            if (!count) return;
            labels.push(bucket.label);
            data.push(count);
            colors.push(bucket.color);
        });

        charts.deal = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{
                    data,
                    backgroundColor: colors,
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '70%',
                plugins: {
                    legend: { position: 'right', labels: { boxWidth: 12, padding: 15, color: '#94a3b8' } }
                }
            }
        });
    }

    function renderWatchlist(context = ownershipContext) {
        const container = document.getElementById('ownership-watchlist');
        if (!container || !context) return;

        const needsStaffing = context.summaryEntries.filter((entry) => entry.linkStatus === 'missing-staffing').slice(0, 4);
        const atRisk = context.summaryEntries
            .filter((entry) => entry.linkStatus === 'linked' && entry.coverageBucket === 'risk')
            .slice(0, 4);
        const unlinked = context.summaryEntries.filter((entry) => entry.linkStatus === 'unlinked').slice(0, 4);

        const staffLoadMap = new Map();
        (Array.isArray(context.enrichedRows) ? context.enrichedRows : []).forEach(({ row }) => {
            const staff = getStaffLabel(row);
            if (!staff) return;
            staffLoadMap.set(staff, (staffLoadMap.get(staff) || 0) + 1);
        });
        const busyStaff = Array.from(staffLoadMap.entries())
            .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
            .slice(0, 4);

        const renderEntries = (entries, emptyText, formatter) => entries.length
            ? entries.map(formatter).join('')
            : `<div class="ownership-watchlist-empty">${escapeHtml(emptyText)}</div>`;

        container.innerHTML = `
            <div class="card-title">Smart Watchlist</div>
            <div class="ownership-watchlist-grid">
                <div class="ownership-watchlist-column">
                    <div class="ownership-watchlist-title">Needs staffing</div>
                    ${renderEntries(needsStaffing, 'No pipeline deals are missing staffing coverage.', (entry) => `
                        <div class="ownership-watchlist-item">
                            <strong>${escapeHtml(getSummaryDisplayLabel(entry))}</strong>
                            <span>${escapeHtml(entry.actionDetail)}</span>
                        </div>
                    `)}
                </div>
                <div class="ownership-watchlist-column">
                    <div class="ownership-watchlist-title">At-risk coverage</div>
                    ${renderEntries(atRisk, 'No linked deals are currently flagged for coverage risk.', (entry) => `
                        <div class="ownership-watchlist-item">
                            <strong>${escapeHtml(getSummaryDisplayLabel(entry))}</strong>
                            <span>${escapeHtml(entry.coverageLabel)} / ${escapeHtml(entry.actionDetail)}</span>
                        </div>
                    `)}
                </div>
                <div class="ownership-watchlist-column">
                    <div class="ownership-watchlist-title">Unlinked staffing</div>
                    ${renderEntries(unlinked, 'All staffing clusters are linked to a pipeline deal.', (entry) => `
                        <div class="ownership-watchlist-item">
                            <strong>${escapeHtml(getSummaryDisplayLabel(entry))}</strong>
                            <span>${escapeHtml(entry.staffCount)} team member${entry.staffCount === 1 ? '' : 's'} / ${escapeHtml(entry.assignments)} assignment${entry.assignments === 1 ? '' : 's'}</span>
                        </div>
                    `)}
                </div>
                <div class="ownership-watchlist-column">
                    <div class="ownership-watchlist-title">Busiest staff</div>
                    ${renderEntries(busyStaff, 'No staff assignments were found.', ([name, count]) => `
                        <div class="ownership-watchlist-item">
                            <strong>${escapeHtml(name)}</strong>
                            <span>${count} assignment${count === 1 ? '' : 's'}</span>
                        </div>
                    `)}
                </div>
            </div>
        `;
    }

    function matchesProfileFilterValue(actualValue, selectedValue) {
        if (!selectedValue || selectedValue === 'all') return true;
        if (Array.isArray(actualValue)) {
            return actualValue.some((value) => normalizeValue(value) === normalizeValue(selectedValue));
        }
        return normalizeValue(actualValue) === normalizeValue(selectedValue);
    }

    function matchesKeywordFilter(entryKeywords, selectedValue) {
        if (!selectedValue || selectedValue === 'all') return true;
        return (Array.isArray(entryKeywords) ? entryKeywords : []).some((entry) => normalizeValue(entry) === normalizeValue(selectedValue));
    }

    function collectVisualFilterOptions(entries, filterKey, limit = 10) {
        const counts = new Map();

        const valuesForEntry = (entry) => {
            if (!entry) return [];
            if (filterKey === 'thematic') return Array.isArray(entry.keywords) ? entry.keywords : [];
            if (filterKey === 'sector') return Array.isArray(entry.sectors) ? entry.sectors : [];
            if (filterKey === 'location') return entry.location ? [entry.location] : [];
            if (filterKey === 'fundingStage') return entry.fundingStage ? [entry.fundingStage] : [];
            if (filterKey === 'revenue') return entry.revenue ? [entry.revenue] : [];
            return [];
        };

        (Array.isArray(entries) ? entries : []).forEach((entry) => {
            valuesForEntry(entry).forEach((value) => {
                const label = String(value || '').trim();
                const key = normalizeValue(label);
                if (!key) return;
                if (!counts.has(key)) {
                    counts.set(key, { value: label, count: 0 });
                }
                counts.get(key).count += 1;
            });
        });

        const selectedValue = ownershipFilters[filterKey] || 'all';
        const options = Array.from(counts.values())
            .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));
        const selectedOption = options.find((option) => normalizeValue(option.value) === normalizeValue(selectedValue));
        const limited = options.slice(0, limit);
        if (selectedOption && !limited.some((option) => normalizeValue(option.value) === normalizeValue(selectedOption.value))) {
            limited.push(selectedOption);
        }
        return limited;
    }

    function buildVisualFilterChips(filterKey, label, options) {
        const activeValue = ownershipFilters[filterKey] || 'all';
        const allCount = options.reduce((sum, option) => sum + option.count, 0);
        const chips = [
            `
                <button class="ownership-visual-chip${activeValue === 'all' ? ' is-active' : ''}" type="button" data-ownership-visual-filter="${filterKey}" data-ownership-visual-value="all">
                    <span>${escapeHtml(label)}</span>
                    <strong>${allCount}</strong>
                </button>
            `,
            ...options.map((option) => `
                <button class="ownership-visual-chip${normalizeValue(activeValue) === normalizeValue(option.value) ? ' is-active' : ''}" type="button" data-ownership-visual-filter="${filterKey}" data-ownership-visual-value="${escapeHtml(option.value)}">
                    <span>${escapeHtml(option.value)}</span>
                    <strong>${option.count}</strong>
                </button>
            `),
        ];

        return `
            <div class="ownership-visual-group">
                <div class="ownership-visual-label">${escapeHtml(label)}</div>
                <div class="ownership-visual-chip-row">
                    ${chips.join('')}
                </div>
            </div>
        `;
    }

    function renderOwnershipVisualFilters(context = ownershipContext) {
        const container = document.getElementById('ownership-visual-filters');
        if (!container || !context) return;

        const linkedEntries = (Array.isArray(context.summaryEntries) ? context.summaryEntries : []).filter((entry) => entry && entry.match);
        const groups = [
            { key: 'thematic', label: 'Thematics', limit: 12 },
            { key: 'sector', label: 'Sectors', limit: 10 },
            { key: 'location', label: 'Location', limit: 10 },
            { key: 'fundingStage', label: 'Funding stage', limit: 10 },
            { key: 'revenue', label: 'Revenue', limit: 8 },
        ].map((group) => ({
            ...group,
            options: collectVisualFilterOptions(linkedEntries, group.key, group.limit),
        })).filter((group) => group.options.length);

        if (!groups.length) {
            container.innerHTML = `
                <div class="card-title">Deal profile filters</div>
                <div class="ownership-watchlist-empty">Add sectors, location, funding stage, revenue, or thematics on deals to unlock visual filters here.</div>
            `;
            return;
        }

        container.innerHTML = `
            <div class="card-title">Deal profile filters</div>
            <div class="ownership-visual-subtitle">Use the richer deal profile fields to narrow coverage by theme, geography, company maturity, and revenue profile.</div>
            <div class="ownership-visual-groups">
                ${groups.map((group) => buildVisualFilterChips(group.key, group.label, group.options)).join('')}
            </div>
        `;
    }

    function buildScopedOwnershipContext(context = ownershipContext) {
        if (!context) return null;
        const summaryEntries = (Array.isArray(context.summaryEntries) ? context.summaryEntries : []).filter((entry) => matchesOwnershipEntryScopeFilters(entry));
        const allowedKeys = new Set(summaryEntries.map((entry) => entry.key));
        const enrichedRows = (Array.isArray(context.enrichedRows) ? context.enrichedRows : []).filter(({ summaryKey }) => allowedKeys.has(summaryKey));
        return {
            ...context,
            summaryEntries,
            summaryEntriesByKey: new Map(summaryEntries.map((entry) => [entry.key, entry])),
            enrichedRows,
        };
    }

    function matchesOwnershipEntryScopeFilters(entry) {
        if (!entry) return false;

        if (ownershipFilters.link !== 'all' && entry.linkStatus !== ownershipFilters.link) {
            return false;
        }
        if (ownershipFilters.coverage !== 'all' && entry.coverageBucket !== ownershipFilters.coverage) {
            return false;
        }
        if (ownershipFilters.stage !== 'all' && entry.stageKey !== ownershipFilters.stage) {
            return false;
        }
        if (!matchesProfileFilterValue(entry.sectors, ownershipFilters.sector)) {
            return false;
        }
        if (!matchesProfileFilterValue(entry.location, ownershipFilters.location)) {
            return false;
        }
        if (!matchesProfileFilterValue(entry.fundingStage, ownershipFilters.fundingStage)) {
            return false;
        }
        if (!matchesProfileFilterValue(entry.revenue, ownershipFilters.revenue)) {
            return false;
        }
        if (!matchesKeywordFilter(entry.keywords, ownershipFilters.thematic)) {
            return false;
        }

        return true;
    }

    function matchesOwnershipEntryFilters(entry, keyword) {
        if (!matchesOwnershipEntryScopeFilters(entry)) return false;

        if (!keyword) return true;
        const haystack = [
            getSummaryDisplayLabel(entry),
            entry.label,
            entry.match && entry.match.company,
            entry.match && entry.match.id,
            entry.match && entry.match.fundraisingDashboardId,
            entry.match && entry.match.seniorOwner,
            entry.match && entry.match.juniorOwner,
            ...(entry.match ? getSubOwners(entry.match) : []),
            entry.roleSummary,
            entry.coverageLabel,
            entry.coverageDetail,
            entry.actionTitle,
            entry.actionDetail,
            entry.stageLabel,
            ...(entry.sectors || []),
            entry.location,
            entry.fundingStage,
            entry.revenue,
            ...(entry.keywords || []),
            ...(entry.staffNames || []),
        ].join(' ').toLowerCase();
        return haystack.includes(keyword);
    }

    function buildCoverageCell(entry) {
        const badgeClass = entry.coverageBucket === 'healthy'
            ? 'ownership-badge ownership-badge-good'
            : 'ownership-badge ownership-badge-warn';

        return `
            <div class="ownership-stack">
                <span class="${badgeClass}">${escapeHtml(entry.coverageLabel)}</span>
                <span class="ownership-muted">${escapeHtml(entry.coverageDetail)}</span>
            </div>
        `;
    }

    function buildNextActionCell(entry) {
        return `
            <div class="ownership-stack">
                <span class="ownership-primary">${escapeHtml(entry.actionTitle)}</span>
                <span class="ownership-muted">${escapeHtml(entry.actionDetail)}</span>
            </div>
        `;
    }

    function buildOwnerGroupContexts(entries) {
        const groups = new Map();

        const ensureGroup = (key, label, options = {}) => {
            if (!groups.has(key)) {
                groups.set(key, {
                    key,
                    label,
                    entries: [],
                    isSynthetic: Boolean(options.isSynthetic),
                });
            }
            return groups.get(key);
        };

        (Array.isArray(entries) ? entries : []).forEach((entry) => {
            if (!entry) return;
            if (!entry.match) {
                ensureGroup('__unlinked__', 'Unlinked staffing', { isSynthetic: true }).entries.push(entry);
                return;
            }

            const owners = getDealOwners(entry.match);
            if (!owners.length) {
                ensureGroup('__no_owner__', 'No deal owner', { isSynthetic: true }).entries.push(entry);
                return;
            }

            owners.forEach((owner) => {
                ensureGroup(`owner:${normalizeValue(owner)}`, owner).entries.push(entry);
            });
        });

        return Array.from(groups.values())
            .map((group) => {
                const groupEntries = Array.isArray(group.entries) ? group.entries.slice() : [];
                groupEntries.sort((a, b) => {
                    const riskOrder = a.coverageBucket === b.coverageBucket ? 0 : a.coverageBucket === 'risk' ? -1 : 1;
                    if (riskOrder) return riskOrder;
                    const aAssignments = Number(a.assignments || 0);
                    const bAssignments = Number(b.assignments || 0);
                    if (aAssignments !== bAssignments) return bAssignments - aAssignments;
                    return getSummaryDisplayLabel(a).localeCompare(getSummaryDisplayLabel(b));
                });

                const attentionCount = groupEntries.filter((entry) => entry.coverageBucket !== 'healthy').length;
                const dealCount = groupEntries.length;
                const stages = new Map();
                groupEntries.forEach((entry) => {
                    const stageLabel = entry.stageLabel || 'No stage';
                    stages.set(stageLabel, (stages.get(stageLabel) || 0) + 1);
                });
                const stageSummary = Array.from(stages.entries())
                    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
                    .slice(0, 3)
                    .map(([label, count]) => `${label}${count > 1 ? ` x${count}` : ''}`)
                    .join(' · ');

                return {
                    ...group,
                    entries: groupEntries,
                    dealCount,
                    attentionCount,
                    healthyCount: groupEntries.filter((entry) => entry.coverageBucket === 'healthy').length,
                    stageSummary,
                };
            })
            .sort((a, b) => {
                if (a.attentionCount !== b.attentionCount) return b.attentionCount - a.attentionCount;
                if (a.dealCount !== b.dealCount) return b.dealCount - a.dealCount;
                return a.label.localeCompare(b.label);
            });
    }

    function buildOwnerGroupHeader(group) {
        const summaryParts = [`${group.dealCount} deal${group.dealCount === 1 ? '' : 's'}`];
        if (group.attentionCount) {
            summaryParts.push(`${group.attentionCount} needs attention`);
        } else {
            summaryParts.push('All covered');
        }
        if (group.stageSummary) {
            summaryParts.push(group.stageSummary);
        }

        const ownerTasksLink = group.isSynthetic
            ? ''
            : `<a class="action-link" href="${buildPageUrl('owner-tasks', { owner: group.label })}">Owner tasks</a>`;

        return `
            <div class="ownership-group-summary">
                <div class="ownership-group-copy">
                    <div class="ownership-group-title-row">
                        <span class="ownership-group-title">${escapeHtml(group.label)}</span>
                        <span class="ownership-badge ${group.attentionCount ? 'ownership-badge-warn' : 'ownership-badge-good'}">
                            ${group.attentionCount ? `${group.attentionCount} flagged` : 'On track'}
                        </span>
                    </div>
                    <div class="ownership-group-meta">${escapeHtml(summaryParts.join(' · '))}</div>
                </div>
                <div class="ownership-group-actions">
                    ${ownerTasksLink}
                </div>
            </div>
        `;
    }

    function renderOwnershipFilterSummary(shownCount, totalCount) {
        const summaryEl = document.getElementById('ownership-filter-summary');
        if (!summaryEl) return;
        const linkFilter = document.getElementById('ownership-link-filter');
        const coverageFilter = document.getElementById('ownership-coverage-filter');

        const viewLabel = ownershipFilters.view === 'rows'
            ? 'staffing rows'
            : ownershipFilters.view === 'owners'
                ? 'owner groups'
                : 'coverage rows';
        const parts = [`Showing ${shownCount} of ${totalCount} ${viewLabel}`];
        if (ownershipFilters.link !== 'all') {
            const linkLabel = linkFilter && linkFilter.selectedOptions && linkFilter.selectedOptions[0]
                ? linkFilter.selectedOptions[0].text
                : ownershipFilters.link.replace(/-/g, ' ');
            parts.push(`link: ${linkLabel}`);
        }
        if (ownershipFilters.coverage !== 'all') {
            const coverageLabel = coverageFilter && coverageFilter.selectedOptions && coverageFilter.selectedOptions[0]
                ? coverageFilter.selectedOptions[0].text
                : ownershipFilters.coverage;
            parts.push(`coverage: ${coverageLabel}`);
        }
        if (ownershipFilters.stage !== 'all') parts.push(`stage: ${getStageLabel(ownershipFilters.stage)}`);
        if (ownershipFilters.sector !== 'all') parts.push(`sectors: ${ownershipFilters.sector}`);
        if (ownershipFilters.location !== 'all') parts.push(`location: ${ownershipFilters.location}`);
        if (ownershipFilters.fundingStage !== 'all') parts.push(`funding stage: ${ownershipFilters.fundingStage}`);
        if (ownershipFilters.revenue !== 'all') parts.push(`revenue: ${ownershipFilters.revenue}`);
        if (ownershipFilters.thematic !== 'all') parts.push(`thematic: ${ownershipFilters.thematic}`);
        if (ownershipFilters.keyword) parts.push(`search: "${ownershipFilters.keyword}"`);
        summaryEl.textContent = `${parts.join(' / ')}.`;
    }

    function renderTable(keyword = "", context = ownershipContext) {
        const head = document.getElementById('table-head');
        const body = document.getElementById('table-body');
        if (!head || !body || !context) return;

        if (ownershipFilters.view === 'owners') {
            head.innerHTML = `
                <tr>
                    <th>Deal / Staffing Cluster</th>
                    <th>Pipeline</th>
                    <th>Staffing Snapshot</th>
                    <th>Coverage</th>
                    <th>Next Action</th>
                    <th>Links</th>
                </tr>
            `;

            const scopedGroups = buildOwnerGroupContexts(context.summaryEntries);
            const filteredEntries = context.summaryEntries.filter((entry) => matchesOwnershipEntryFilters(entry, keyword));
            const filteredGroups = buildOwnerGroupContexts(filteredEntries);

            body.innerHTML = filteredGroups.length
                ? filteredGroups.map((group) => `
                    <tr class="ownership-group-row">
                        <td colspan="6">${buildOwnerGroupHeader(group)}</td>
                    </tr>
                    ${group.entries.map((entry) => `
                        <tr class="ownership-group-deal-row">
                            <td class="ownership-rich-cell">${buildLinkedDealCell(entry.match, entry, entry.label)}</td>
                            <td class="ownership-rich-cell">${buildPipelineCell(entry.match)}</td>
                            <td class="ownership-rich-cell">${buildStaffingSnapshotCell(entry)}</td>
                            <td class="ownership-rich-cell">${buildCoverageCell(entry)}</td>
                            <td class="ownership-rich-cell">${buildNextActionCell(entry)}</td>
                            <td class="ownership-rich-cell">${buildLinksCell(entry.match)}</td>
                        </tr>
                    `).join('')}
                `).join('')
                : '<tr><td colspan="6" class="ownership-rich-cell">No owner groups match the current filters.</td></tr>';

            renderOwnershipFilterSummary(filteredGroups.length, scopedGroups.length);
            return;
        }

        if (ownershipFilters.view === 'rows') {
            const headers = context.headers;
            head.innerHTML = `<tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}<th>Linked Deal</th><th>Pipeline</th><th>Staffing Snapshot</th><th>Coverage</th><th>Next Action</th><th>Links</th></tr>`;

            const filteredRows = context.enrichedRows.filter(({ row, summaryKey, primaryDealLabel }) => {
                const entry = context.summaryEntriesByKey.get(summaryKey);
                if (!entry || !matchesOwnershipEntryScopeFilters(entry)) return false;
                if (!keyword) return true;
                const rowHaystack = [
                    ...context.headers.map((header) => String(row[header] || '')),
                    primaryDealLabel,
                ].join(' ').toLowerCase();
                return rowHaystack.includes(keyword) || matchesOwnershipEntryFilters(entry, keyword);
            });

            body.innerHTML = filteredRows.length
                ? filteredRows.map(({ row, match, summaryKey, primaryDealLabel }) => {
                    const summary = context.summaries.get(summaryKey);
                    const entry = context.summaryEntriesByKey.get(summaryKey);
                    return `
                        <tr>
                            ${context.headers.map(h => `<td>${escapeHtml(row[h] || '–')}</td>`).join('')}
                            <td class="ownership-rich-cell">${buildLinkedDealCell(match, summary, primaryDealLabel)}</td>
                            <td class="ownership-rich-cell">${buildPipelineCell(match)}</td>
                            <td class="ownership-rich-cell">${buildStaffingSnapshotCell(summary)}</td>
                            <td class="ownership-rich-cell">${buildCoverageCell(entry)}</td>
                            <td class="ownership-rich-cell">${buildNextActionCell(entry)}</td>
                            <td class="ownership-rich-cell">${buildLinksCell(match)}</td>
                        </tr>
                    `;
                }).join('')
                : `<tr><td colspan="${context.headers.length + 6}" class="ownership-rich-cell">No staffing rows match the current filters.</td></tr>`;

            renderOwnershipFilterSummary(filteredRows.length, context.enrichedRows.length);
            return;
        }

        head.innerHTML = `
            <tr>
                <th>Deal / Staffing Cluster</th>
                <th>Pipeline</th>
                <th>Staffing Snapshot</th>
                <th>Coverage</th>
                <th>Next Action</th>
                <th>Links</th>
            </tr>
        `;

        const filteredEntries = context.summaryEntries.filter((entry) => matchesOwnershipEntryFilters(entry, keyword));
        body.innerHTML = filteredEntries.length
            ? filteredEntries.map((entry) => `
                <tr>
                    <td class="ownership-rich-cell">${buildLinkedDealCell(entry.match, entry, entry.label)}</td>
                    <td class="ownership-rich-cell">${buildPipelineCell(entry.match)}</td>
                    <td class="ownership-rich-cell">${buildStaffingSnapshotCell(entry)}</td>
                    <td class="ownership-rich-cell">${buildCoverageCell(entry)}</td>
                    <td class="ownership-rich-cell">${buildNextActionCell(entry)}</td>
                    <td class="ownership-rich-cell">${buildLinksCell(entry.match)}</td>
                </tr>
            `).join('')
            : '<tr><td colspan="6" class="ownership-rich-cell">No coverage rows match the current filters.</td></tr>';

        renderOwnershipFilterSummary(filteredEntries.length, context.summaryEntries.length);
    }

    function showLoader(msg) {
        document.getElementById('loader').style.display = 'flex';
        document.getElementById('loading-spinner').style.display = 'block';
        document.getElementById('error-panel').style.display = 'none';
        document.getElementById('dashboard').style.display = 'none';
        if (msg) document.getElementById('loader-msg').textContent = msg;
    }

    function showDashboard() {
        document.getElementById('loader').style.display = 'none';
        document.getElementById('dashboard').style.display = 'block';
    }

    function showError(msg) {
        document.getElementById('loader').style.display = 'flex';
        document.getElementById('loading-spinner').style.display = 'none';
        document.getElementById('error-panel').style.display = 'block';
        document.getElementById('error-details').textContent = msg;
    }

})(window);
