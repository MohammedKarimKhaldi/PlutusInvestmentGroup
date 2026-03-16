(function initDealOwnership() {
    const AppCore = window.AppCore;
    const DASHBOARD_ID = "deal-ownership";
    
    let rawData = [];
    let dealsSource = [];
    let dealsMap = new Map();
    let charts = {};
    const COLOR_PALETTE = ['#818cf8', '#34d399', '#fbbf24', '#a78bfa', '#f472b6', '#22d3ee', '#fb7185'];

    window.addEventListener('load', () => {
        initializeDashboard();
        
        const searchInput = document.getElementById('search');
        if (searchInput) {
            searchInput.addEventListener('input', (e) => renderTable(e.target.value.toLowerCase()));
        }
        
        const retryBtn = document.getElementById('btn-retry-sync');
        if (retryBtn) {
            retryBtn.addEventListener('click', initializeDashboard);
        }

        window.addEventListener('appcore:deals-updated', () => {
            buildDealsMap();
            const search = document.getElementById('search');
            renderTable(search ? search.value.toLowerCase() : "");
        });

        window.addEventListener('appcore:dashboard-config-updated', () => {
            const search = document.getElementById('search');
            renderTable(search ? search.value.toLowerCase() : "");
        });
    });

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

        renderKPIs();
        renderCharts();
        renderTable();
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

    function getStageBadge(stage) {
        const normalized = String(stage || '').trim().toLowerCase();
        const label = normalized
            ? normalized.charAt(0).toUpperCase() + normalized.slice(1)
            : 'No stage';
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

    function buildOwnershipContext() {
        const headers = rawData.length ? Object.keys(rawData[0]) : [];
        const summaries = new Map();
        const linkedDealIds = new Set();
        let unlinkedRows = 0;

        const enrichedRows = rawData.map((row) => {
            const match = findDealForRow(row, headers);
            const primaryDealLabel = getPrimaryDealLabel(row, headers) || 'Unlabelled staffing line';
            const summaryKey = match && match.id
                ? `deal:${String(match.id)}`
                : `unlinked:${String(primaryDealLabel || '').trim().toLowerCase()}`;

            if (!summaries.has(summaryKey)) {
                summaries.set(summaryKey, {
                    assignments: 0,
                    staff: new Set(),
                    roles: new Map(),
                    label: primaryDealLabel,
                    match,
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

            if (match && match.id) linkedDealIds.add(String(match.id));
            if (!match) unlinkedRows += 1;

            return {
                row,
                match,
                summaryKey,
                primaryDealLabel,
            };
        });

        return {
            headers,
            enrichedRows,
            summaries,
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

        const senior = escapeHtml(match.seniorOwner || match.owner || '–');
        const junior = escapeHtml(match.juniorOwner || '–');
        return `
            <div class="ownership-stack">
                ${getStageBadge(match.stage)}
                <span class="ownership-muted">Owners: ${senior} / ${junior}</span>
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
                <span class="ownership-primary">${teamSize} staff · ${assignmentCount} assignments</span>
                <span class="ownership-muted">${escapeHtml(roleSummary)}</span>
            </div>
        `;
    }

    function renderKPIs() {
        const container = document.getElementById('staffing-kpis');
        if (!container) return;

        const context = buildOwnershipContext();
        const totalRows = rawData.length;
        const uniqueStaff = new Set(rawData.map(r => r["Staff"] || r["Name"] || r["Lead"] || "").filter(Boolean)).size;
        const uniqueDeals = new Set(rawData.map(r => r["Deal"] || r["Company"] || r["Project"] || "").filter(Boolean)).size;

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
                <div class="kpi-label">Unique Deals</div>
                <div class="kpi-value">${uniqueDeals}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Linked Pipeline Deals</div>
                <div class="kpi-value">${context.linkedDealCount}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Unlinked Staffing Rows</div>
                <div class="kpi-value" style="color: ${context.unlinkedRows ? 'var(--warning)' : 'var(--success)'};">${context.unlinkedRows}</div>
            </div>
        `;
    }

    function renderCharts() {
        renderAllocationChart();
        renderDealDistChart();
    }

    function renderAllocationChart() {
        const ctx = document.getElementById('staffChart').getContext('2d');
        if (charts.staff) charts.staff.destroy();

        const staffMap = {};
        rawData.forEach(r => {
            const name = r["Staff"] || r["Name"] || r["Lead"] || "Unknown";
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

    function renderDealDistChart() {
        const ctx = document.getElementById('dealChart').getContext('2d');
        if (charts.deal) charts.deal.destroy();

        const dealMap = {};
        rawData.forEach(r => {
            const deal = r["Deal"] || r["Company"] || r["Project"] || "Other";
            dealMap[deal] = (dealMap[deal] || 0) + 1;
        });

        const labels = Object.keys(dealMap);
        const data = Object.values(dealMap);

        charts.deal = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{
                    data,
                    backgroundColor: COLOR_PALETTE,
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

    function renderTable(keyword = "") {
        const head = document.getElementById('table-head');
        const body = document.getElementById('table-body');
        if (!head || !body || rawData.length === 0) return;

        const context = buildOwnershipContext();
        const headers = context.headers;
        head.innerHTML = `<tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}<th>Linked Deal</th><th>Pipeline</th><th>Staffing Snapshot</th><th>Links</th></tr>`;

        const filtered = context.enrichedRows.filter(({ row, match, summaryKey, primaryDealLabel }) => {
            if (!keyword) return true;
            const summary = context.summaries.get(summaryKey);
            const searchParts = [
                ...headers.map((h) => String(row[h] || '')),
                primaryDealLabel,
                match && match.name,
                match && match.company,
                match && match.id,
                match && match.fundraisingDashboardId,
                summary && getTopRoleSummary(summary),
            ];
            return searchParts.some((value) => String(value || '').toLowerCase().includes(keyword));
        });

        body.innerHTML = filtered.map(({ row, match, summaryKey, primaryDealLabel }) => {
            const summary = context.summaries.get(summaryKey);
            return `
                <tr>
                    ${headers.map(h => `<td>${escapeHtml(row[h] || '–')}</td>`).join('')}
                    <td class="ownership-rich-cell">${buildLinkedDealCell(match, summary, primaryDealLabel)}</td>
                    <td class="ownership-rich-cell">${buildPipelineCell(match)}</td>
                    <td class="ownership-rich-cell">${buildStaffingSnapshotCell(summary)}</td>
                    <td class="ownership-rich-cell">${buildLinksCell(match)}</td>
                </tr>
            `;
        }).join('');
    }

    function showLoader(msg) {
        document.getElementById('loader').style.display = 'flex';
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
