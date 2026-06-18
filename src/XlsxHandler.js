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

        document.getElementById("exportProfilesBtn").addEventListener("click", () => {
            this._exportProfilesCSV();
        });

        // If .jab was loaded after XLSX, trigger cost computation
        document.addEventListener('userInfosReady', () => {
            if (this.flatGroupable) {
                workLogHandler.loadFromXlsx(this.flatGroupable);
            }
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

        // Trigger Cost Report if userInfos already loaded
        if (this.flatGroupable && workLogHandler.userInfos) {
            workLogHandler.loadFromXlsx(this.flatGroupable);
        } else {
            this._switchTab("tab-projects");
        }
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
        if (!this.flatGroupable && !this._projectsData) return;

        if (this.flatGroupable) {
            this._projectsData = this._gameProjects().map(p => this._computeProject(p));
            this._projectsData.sort((a, b) => b.totalHours - a.totalHours);
            // Collect all sprints across all game projects, sorted
            const allSprintsSet = new Set();
            this._projectsData.forEach(p => Object.keys(p.sprintDates).forEach(s => allSprintsSet.add(s)));
            this._allSprints = this._sortSprints([...allSprintsSet]);
        }

        const projects = this._projectsData;
        const hasCost = !!workLogHandler.userInfos ||
            projects.some(p => p.totalManday > 0);

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

        // Only auto-render profiles here when called from live XLSX load (not restore)
        if (!this._restoringFromFirebase) this._renderProjectProfiles();
    }

    // ── Phase profiles (per-project breakdown) ─────────────────────────

    /**
     * Count weekdays (Mon–Fri) between start and end that have no log entry.
     * logDaySet: Set of date strings (d.toDateString())
     */
    _countBreakDays(start, end, logDaySet) {
        let breaks = 0;
        const cur = new Date(start);
        cur.setHours(0, 0, 0, 0);
        const endD = new Date(end);
        endD.setHours(0, 0, 0, 0);
        while (cur <= endD) {
            const dow = cur.getDay();
            if (dow !== 0 && dow !== 6 && !logDaySet.has(cur.toDateString())) {
                breaks++;
            }
            cur.setDate(cur.getDate() + 1);
        }
        return breaks;
    }

    /** Compute per-sprint stats for one project. */
    _computePhases(project) {
        const rows = this.flatGroupable.filter(r => r['Project Name'] === project);
        const phases = {};

        rows.forEach(r => {
            const sprint = r['Sprint'] || '(No Sprint)';
            const raw = r['Log Date & Time'];
            if (!raw) return;
            const d = raw instanceof Date ? raw : new Date(raw);
            if (isNaN(d)) return;

            if (!phases[sprint]) {
                phases[sprint] = { min: d, max: d, users: new Set(), hours: 0, manday: 0, logDays: new Set() };
            } else {
                if (d < phases[sprint].min) phases[sprint].min = d;
                if (d > phases[sprint].max) phases[sprint].max = d;
            }

            phases[sprint].logDays.add(d.toDateString());
            phases[sprint].users.add(r['Log user']);
            const h = +r['Hr. Spent'] || 0;
            phases[sprint].hours += h;
            const userInfo = workLogHandler.userInfos?.[r['Log user']];
            if (userInfo) phases[sprint].manday += h * userInfo.cost;
        });

        return this._sortSprints(Object.keys(phases)).map(s => {
            const p = phases[s];
            // Count distinct weekdays that have at least 1 log
            const loggedDays = [...p.logDays].filter(ds => {
                const dow = new Date(ds).getDay();
                return dow !== 0 && dow !== 6;
            }).length;
            return {
                name: s,
                min: p.min,
                max: p.max,
                people: p.users.size,
                hours: p.hours,
                manday: p.manday,
                loggedDays,
                breakDays: this._countBreakDays(p.min, p.max, p.logDays),
            };
        });
    }

    _renderProjectProfiles(precomputedPhases = null) {
        if (!this._projectsData) return;
        const hasCost = !!workLogHandler.userInfos ||
            this._projectsData.some(p => (precomputedPhases?.[p.project] || []).some(ph => ph.manday > 0));

        // Keep the phases used for rendering so the Export button can reuse them
        // (works for both live XLSX and restored payloads).
        this._phasesData = {};

        let html = '';
        this._projectsData.forEach(({ project }) => {
            const phases = precomputedPhases?.[project] ??
                (this.flatGroupable ? this._computePhases(project) : []);
            if (phases.length === 0) return;
            this._phasesData[project] = phases;

            const phaseRows = phases.map(p => {
                const breakClass = p.breakDays >= 10 ? 'break-high' : p.breakDays >= 5 ? 'break-mid' : '';
                return `<tr>
                    <td class="phase-name-cell">${this._fmtSprint(p.name)}</td>
                    <td class="center-cell">${fmtDate(p.min)}</td>
                    <td class="center-cell">${fmtDate(p.max)}</td>
                    <td class="center-cell">${p.people}</td>
                    <td class="right-cell">${p.hours.toFixed(1)} h</td>
                    ${hasCost ? `<td class="right-cell">${p.manday.toFixed(2)}</td>` : ''}
                    <td class="center-cell">${p.loggedDays}</td>
                    <td class="center-cell ${breakClass}">${p.breakDays}</td>
                </tr>`;
            }).join('');

            html += `
            <div class="profile-card">
                <div class="profile-header">${project}</div>
                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th style="text-align:left">Phase</th>
                                <th>Start</th>
                                <th>End</th>
                                <th>People</th>
                                <th>Hours</th>
                                ${hasCost ? '<th>Mandays</th>' : ''}
                                <th title="Distinct weekdays with at least 1 log entry">Log Days</th>
                                <th title="Weekdays without any log between start–end">Break Days</th>
                            </tr>
                        </thead>
                        <tbody>${phaseRows}</tbody>
                    </table>
                </div>
            </div>`;
        });

        const container = document.getElementById("projectProfiles");
        container.innerHTML = html;
        document.getElementById("projectProfilesSection").style.display = "block";
        document.getElementById("exportProfilesBtn").disabled =
            Object.keys(this._phasesData).length === 0;

        // Signal that project data is ready (LocalDbHandler listens to auto-save)
        document.dispatchEvent(new Event('reportReady'));
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

    _exportProfilesCSV() {
        if (!this._phasesData || Object.keys(this._phasesData).length === 0) return;
        const hasCost = Object.values(this._phasesData)
            .some(arr => arr.some(p => p.manday > 0));
        const q = v => `"${String(v).replace(/"/g, '""')}"`;

        const headers = ['Project', 'Phase', 'Start', 'End', 'People', 'Hours'];
        if (hasCost) headers.push('Mandays');
        headers.push('Log Days', 'Break Days');

        const rows = [];
        Object.entries(this._phasesData).forEach(([project, phases]) => {
            phases.forEach(p => {
                const cols = [
                    project, this._fmtSprint(p.name), fmtDate(p.min), fmtDate(p.max),
                    p.people, p.hours.toFixed(2),
                ];
                if (hasCost) cols.push(p.manday.toFixed(4));
                cols.push(p.loggedDays, p.breakDays);
                rows.push(cols.map(q).join(','));
            });
        });

        const csv  = [headers.map(q).join(','), ...rows].join('\n');
        const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href = url; a.download = 'project_phase_profiles.csv';
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    /** Restore project analysis from Firebase-saved payload (no XLSX needed). */
    restoreFromData(overview, allSprints, phases) {
        this._projectsData = overview;
        this._allSprints   = allSprints?.length
            ? allSprints
            : this._sortSprints([...new Set(overview.flatMap(p => Object.keys(p.sprintDates)))]);
        this._restoringFromFirebase = true;
        this._renderProjectTable();
        this._restoringFromFirebase = false;
        this._renderProjectProfiles(phases);
        document.getElementById("projectAnalysisSection").style.display = "block";
        document.getElementById("projectEmptyState").style.display      = "none";
        document.getElementById("exportProjectsBtn").disabled = false;
        this._switchTab("tab-projects");
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
