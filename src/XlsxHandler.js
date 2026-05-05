import workLogHandler from "./WorkLogHandler.js";

const DONE_STATUSES = new Set(['CLOSED', 'Done', 'FIXED/DONE', 'FIXED / DONE']);

const fmtDate = (d) => {
    if (!d) return '—';
    const date = d instanceof Date ? d : new Date(d);
    if (isNaN(date)) return '—';
    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    return `${dd}/${mm}/${date.getFullYear()}`;
};

class XlsxHandler {

    constructor() {
        this.flatGroupable   = null;
        this.flatIssueDaywise = null;
        this._listenEvents();
    }

    _listenEvents() {
        document.getElementById("xlsxFile").addEventListener("change", async (e) => {
            const file = e.target.files[0];
            if (!file) return;
            document.getElementById("xlsxFileName").textContent = file.name;
            await this._loadXlsx(file);
        });

        document.getElementById("exportProjectsBtn").addEventListener("click", () => {
            this._exportProjectsCSV();
        });
    }

    async _loadXlsx(file) {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

        const groupableSheet   = wb.Sheets['Flat (Groupable)'];
        const issueDaywiseSheet = wb.Sheets['Flat - (Issue daywise)'];

        if (groupableSheet) {
            this.flatGroupable = XLSX.utils.sheet_to_json(groupableSheet);
        }
        if (issueDaywiseSheet) {
            // Row 0 = week labels, row 1 = day sub-labels; skip both via defval
            this.flatIssueDaywise = XLSX.utils.sheet_to_json(issueDaywiseSheet, { defval: 0 });
        }

        const count = this.flatGroupable?.length ?? 0;
        this._setStatus("xlsx", `${count} records loaded`);
        this._renderProjectTable();
        this._switchTab("tab-projects");
    }

    // ── Data helpers ───────────────────────────────────────────────────

    _gameProjects() {
        if (!this.flatGroupable) return [];
        const seen = new Set();
        this.flatGroupable.forEach(r => {
            const name = r['Project Name'] ?? '';
            if (/^9\d{3}/.test(name)) seen.add(name);
        });
        return [...seen].sort();
    }

    _computeProject(project) {
        const rows = this.flatGroupable.filter(r => r['Project Name'] === project);

        // Team size: unique users with >16h total
        const userHours = {};
        rows.forEach(r => {
            const u = r['Log user'];
            if (u) userHours[u] = (userHours[u] || 0) + (+r['Hr. Spent'] || 0);
        });
        const teamSize = Object.values(userHours).filter(h => h > 16).length;

        // Progress: unique tickets (latest status wins)
        const ticketStatus = {};
        rows.forEach(r => {
            const t = r['Ticket No'];
            if (t) ticketStatus[t] = r['Status'];
        });
        const total  = Object.keys(ticketStatus).length;
        const done   = Object.values(ticketStatus).filter(s => DONE_STATUSES.has(s)).length;
        const pct    = total > 0 ? done / total * 100 : 0;

        // Total log work & manday
        const userInfos = workLogHandler.userInfos || {};
        let totalHours = 0, totalManday = 0;
        rows.forEach(r => {
            const h = +r['Hr. Spent'] || 0;
            totalHours += h;
            const cost = userInfos[r['Log user']]?.cost ?? 0;
            totalManday += h * cost;
        });

        // Sprint date ranges (min/max Log Date per sprint)
        const sprintDates = {};
        rows.forEach(r => {
            const sprint = r['Sprint'];
            if (!sprint) return;
            const raw = r['Log Date & Time'];
            if (!raw) return;
            const d = raw instanceof Date ? raw : new Date(raw);
            if (isNaN(d)) return;
            if (!sprintDates[sprint]) sprintDates[sprint] = { min: d, max: d };
            else {
                if (d < sprintDates[sprint].min) sprintDates[sprint].min = d;
                if (d > sprintDates[sprint].max) sprintDates[sprint].max = d;
            }
        });

        return { project, teamSize, pct, done, total, totalHours, totalManday, sprintDates };
    }

    // ── Helpers ────────────────────────────────────────────────────────

    _sortSprints(sprints) {
        return [...sprints].sort((a, b) => {
            const na = parseInt(a) ?? 999;
            const nb = parseInt(b) ?? 999;
            return na !== nb ? na - nb : a.localeCompare(b);
        });
    }

    _fmtSprint(name) {
        return name.replace(/^\d+\s*-\s*/, '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    }

    // ── Render ─────────────────────────────────────────────────────────

    _renderProjectTable() {
        if (!this.flatGroupable) return;

        this._projectsData = this._gameProjects().map(p => this._computeProject(p));
        this._projectsData.sort((a, b) => b.totalHours - a.totalHours);
        const projects = this._projectsData;

        // Collect all sprints across all game projects, sorted
        const allSprintsSet = new Set();
        projects.forEach(p => Object.keys(p.sprintDates).forEach(s => allSprintsSet.add(s)));
        this._allSprints = this._sortSprints([...allSprintsSet]);

        const hasCost = !!workLogHandler.userInfos;

        // ── thead: 2 rows ──────────────────────────────────────────────
        const fixedCols = `
            <th rowspan="2">Project</th>
            <th rowspan="2">Team Size</th>
            <th rowspan="2" style="min-width:200px">Progress</th>
            <th rowspan="2">Total Log Work</th>
            ${hasCost ? '<th rowspan="2">Manday</th>' : ''}`;

        const sprintGroupHeaders = this._allSprints
            .map(s => `<th colspan="2" class="sprint-group-header">${this._fmtSprint(s)}</th>`)
            .join('');

        const sprintSubHeaders = this._allSprints
            .map(() => `<th class="sprint-sub-header">Start</th><th class="sprint-sub-header">End</th>`)
            .join('');

        let html = `<table>
            <thead>
                <tr>${fixedCols}${sprintGroupHeaders}</tr>
                <tr>${sprintSubHeaders}</tr>
            </thead>
            <tbody>`;

        // ── tbody ──────────────────────────────────────────────────────
        projects.forEach(({ project, teamSize, pct, done, total, totalHours, totalManday, sprintDates }) => {
            const barColor = pct >= 90 ? '#16a34a' : pct >= 60 ? '#4f46e5' : pct >= 30 ? '#f59e0b' : '#ef4444';
            const sprintCells = this._allSprints.map(s => {
                const range = sprintDates[s];
                if (!range) return `<td class="zero-cell">—</td><td class="zero-cell">—</td>`;
                return `<td class="center-cell">${fmtDate(range.min)}</td><td class="center-cell">${fmtDate(range.max)}</td>`;
            }).join('');

            html += `<tr>
                <td class="project-name-cell">${project}</td>
                <td class="center-cell">${teamSize}</td>
                <td>
                    <div class="progress-wrap">
                        <div class="progress-bar" style="width:${pct.toFixed(1)}%;background:${barColor}"></div>
                        <span class="progress-label">${pct.toFixed(1)}%&nbsp;<span style="color:#9ca3af">(${done}/${total})</span></span>
                    </div>
                </td>
                <td class="right-cell">${totalHours.toFixed(1)} h</td>
                ${hasCost ? `<td class="right-cell">${totalManday.toFixed(2)}</td>` : ''}
                ${sprintCells}
            </tr>`;
        });

        html += `</tbody></table>`;

        document.getElementById("projectAnalysisTable").innerHTML = html;
        document.getElementById("projectAnalysisSection").style.display = "block";
        document.getElementById("projectEmptyState").style.display = "none";
        document.getElementById("exportProjectsBtn").disabled = false;
    }

    _exportProjectsCSV() {
        if (!this._projectsData || !this._allSprints) return;
        const hasCost = !!workLogHandler.userInfos;
        const q = v => `"${String(v).replace(/"/g, '""')}"`;

        const headers = ['Project', 'Team Size', 'Progress (%)', 'Done Tickets', 'Total Tickets', 'Total Log Work (h)'];
        if (hasCost) headers.push('Manday');
        this._allSprints.forEach(s => {
            const label = this._fmtSprint(s);
            headers.push(`${label} Start`, `${label} End`);
        });

        const rows = this._projectsData.map(({ project, teamSize, pct, done, total, totalHours, totalManday, sprintDates }) => {
            const cols = [project, teamSize, pct.toFixed(2), done, total, totalHours.toFixed(2)];
            if (hasCost) cols.push(totalManday.toFixed(4));
            this._allSprints.forEach(s => {
                const range = sprintDates[s];
                cols.push(range ? fmtDate(range.min) : '—');
                cols.push(range ? fmtDate(range.max) : '—');
            });
            return cols.map(q).join(',');
        });

        const csv = [headers.map(q).join(','), ...rows].join('\n');
        const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'game_projects_overview.csv';
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ── Utils ──────────────────────────────────────────────────────────

    _switchTab(tabId) {
        document.querySelectorAll(".tab-btn").forEach(b =>
            b.classList.toggle("active", b.dataset.tab === tabId)
        );
        document.querySelectorAll(".tab-content").forEach(c =>
            c.classList.toggle("active", c.id === tabId)
        );
    }

    _setStatus(type, message) {
        const badge = document.getElementById(`status-${type}`);
        const text  = document.getElementById(`status-${type}-text`);
        if (badge) badge.classList.add("loaded");
        if (text)  text.textContent = message;
    }
}

export default new XlsxHandler();
