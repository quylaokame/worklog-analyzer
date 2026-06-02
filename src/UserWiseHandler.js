/**
 * UserWiseHandler — per-user task log analysis.
 * Requires flatGroupable to be loaded from XLSX.
 */
import xlsxHandler    from './XlsxHandler.js';
import workLogHandler from './WorkLogHandler.js';

// ── Column accessors ───────────────────────────────────────────────────────
const col = {
    project: r => r['Project Name']     || '',
    sprint:  r => r['Sprint']           || '',
    ticket:  r => r['Ticket No']        || r['Issue Key']     || '',
    parent:  r => r['Parent Issue Key'] || r['Parent']        || '',
    user:    r => r['Log user']         || '',
    hours:   r => +r['Hr. Spent']       || 0,
    date:    r => {
        const raw = r['Log Date & Time'];
        if (!raw) return null;
        const d = raw instanceof Date ? raw : new Date(raw);
        return isNaN(d) ? null : d;
    },
};

/**
 * Detect the summary/title column name dynamically from the first XLSX row.
 * Returns the column accessor function r => value.
 */
function detectSummaryAccessor(rows) {
    if (!rows?.length) return () => '';
    const keys = Object.keys(rows[0]);

    // Priority-ordered exact names
    const exact = [
        'Issue Summary', 'Summary', 'Ticket Summary',
        'Issue Title', 'Title', 'Work Description', 'Description',
    ];
    const hit = exact.find(k => keys.includes(k));
    if (hit) return r => r[hit] || '';

    // Fuzzy: first key whose name contains 'summary' or 'title' (case-insensitive)
    const fuzzy = keys.find(k => /summary|title/i.test(k));
    if (fuzzy) return r => r[fuzzy] || '';

    // Nothing found — fall back to empty string (ID will be used as label)
    return () => '';
}

// ── Date helpers ───────────────────────────────────────────────────────────
const fmtDate = d => {
    if (!d || isNaN(d)) return '—';
    return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
};

function countWeekdays(start, end) {
    let n = 0;
    const cur = new Date(start); cur.setHours(0,0,0,0);
    const fin = new Date(end);   fin.setHours(0,0,0,0);
    while (cur <= fin) {
        const d = cur.getDay();
        if (d !== 0 && d !== 6) n++;
        cur.setDate(cur.getDate() + 1);
    }
    return n;
}

function countBreakDays(start, end, logDaySet) {
    let n = 0;
    const cur = new Date(start); cur.setHours(0,0,0,0);
    const fin = new Date(end);   fin.setHours(0,0,0,0);
    while (cur <= fin) {
        const d = cur.getDay();
        if (d !== 0 && d !== 6 && !logDaySet.has(cur.toDateString())) n++;
        cur.setDate(cur.getDate() + 1);
    }
    return n;
}

function weekdayCount(logDaySet) {
    return [...logDaySet].filter(ds => {
        const d = new Date(ds).getDay();
        return d !== 0 && d !== 6;
    }).length;
}

// ── Class ──────────────────────────────────────────────────────────────────
class UserWiseHandler {

    constructor() {
        this._initEvents();
    }

    _initEvents() {
        document.getElementById('userSelect').addEventListener('change', () => this._renderTable());
        document.addEventListener('reportReady', () => this._onDataReady());
    }

    _onDataReady() {
        if (!xlsxHandler.flatGroupable) return;
        this._populateUsers();
    }

    // ── Populate dropdown ──────────────────────────────────────────────────

    _populateUsers() {
        const users = new Set();
        xlsxHandler.flatGroupable.forEach(r => { const u = col.user(r); if (u) users.add(u); });

        const sel  = document.getElementById('userSelect');
        const prev = sel.value;

        sel.innerHTML = '<option value="">— Select User —</option>' +
            [...users].sort().map(u => `<option value="${u}">${u}</option>`).join('');

        if (prev && users.has(prev)) { sel.value = prev; this._renderTable(); }
        else {
            document.getElementById('userWiseEmptyState').style.display = 'block';
            document.getElementById('userWiseContent').style.display    = 'none';
        }
    }

    // ── Data computation ───────────────────────────────────────────────────

    /**
     * Build ticket → summary map from ALL rows so parent summaries are resolvable.
     * Column name is detected once from the first row.
     */
    _buildSummaryMap() {
        const getSummary = detectSummaryAccessor(xlsxHandler.flatGroupable);
        const map = {};
        xlsxHandler.flatGroupable.forEach(r => {
            const t = col.ticket(r);
            const s = getSummary(r);
            if (t && s && !map[t]) map[t] = s;
        });
        return map;
    }

    _getTasksForUser(user) {
        const summaryMap = this._buildSummaryMap();
        const userCost   = workLogHandler.userInfos?.[user]?.cost ?? 0;
        const rows       = xlsxHandler.flatGroupable.filter(r => col.user(r) === user);

        const map = {};
        rows.forEach(r => {
            const ticket = col.ticket(r);
            if (!ticket) return;

            if (!map[ticket]) {
                const parent      = col.parent(r);
                const taskId      = parent || ticket;
                map[ticket] = {
                    project:     col.project(r),
                    sprint:      col.sprint(r),
                    taskId,
                    taskName:    summaryMap[taskId]  || taskId,
                    subTaskId:   parent ? ticket : '',
                    subTaskName: parent ? (summaryMap[ticket] || ticket) : '',
                    hours:       0,
                    manday:      0,
                    logDays:     new Set(),
                    min:         null,
                    max:         null,
                };
            }

            const t = map[ticket];
            const h = col.hours(r);
            t.hours  += h;
            t.manday += h * userCost;

            const d = col.date(r);
            if (d) {
                t.logDays.add(d.toDateString());
                if (!t.min || d < t.min) t.min = d;
                if (!t.max || d > t.max) t.max = d;
            }
        });

        return Object.values(map)
            .map(t => ({
                ...t,
                loggedDays: weekdayCount(t.logDays),
                totalDays:  t.min && t.max ? countWeekdays(t.min, t.max) : 0,
                breakDays:  t.min && t.max ? countBreakDays(t.min, t.max, t.logDays) : 0,
            }))
            .sort((a, b) => {
                if (a.project !== b.project) return a.project.localeCompare(b.project);
                const sa = parseInt(a.sprint) || 999;
                const sb = parseInt(b.sprint) || 999;
                if (sa !== sb) return sa - sb;
                return a.taskId.localeCompare(b.taskId);
            });
    }

    // ── Render ─────────────────────────────────────────────────────────────

    _renderTable() {
        const user = document.getElementById('userSelect').value;

        if (!user || !xlsxHandler.flatGroupable) {
            document.getElementById('userWiseEmptyState').style.display = 'block';
            document.getElementById('userWiseContent').style.display    = 'none';
            return;
        }

        const tasks   = this._getTasksForUser(user);
        const hasCost = !!workLogHandler.userInfos;

        // ── Global stats ───────────────────────────────────────────────────
        let globalMin = null, globalMax = null;
        tasks.forEach(t => {
            if (t.min && (!globalMin || t.min < globalMin)) globalMin = t.min;
            if (t.max && (!globalMax || t.max > globalMax)) globalMax = t.max;
        });

        const totalWorkDays  = globalMin && globalMax ? countWeekdays(globalMin, globalMax) : 0;
        const totalWorkHours = totalWorkDays * 8;
        const totalLogged    = tasks.reduce((s, t) => s + t.hours, 0);
        const totalManday    = tasks.reduce((s, t) => s + t.manday, 0);
        const utilPct        = totalWorkHours > 0
            ? (totalLogged / totalWorkHours * 100).toFixed(1) : '—';

        // ── Period panel ───────────────────────────────────────────────────
        document.getElementById('userWisePeriod').innerHTML = globalMin ? `
            <div class="period-card">
                <div class="period-item">
                    <div class="period-label">Period</div>
                    <div class="period-value">${fmtDate(globalMin)} → ${fmtDate(globalMax)}</div>
                </div>
                <div class="period-item">
                    <div class="period-label">Working Days</div>
                    <div class="period-value">${totalWorkDays} d</div>
                </div>
                <div class="period-item">
                    <div class="period-label">Working Hours Capacity</div>
                    <div class="period-value">${totalWorkHours} h</div>
                </div>
                <div class="period-item">
                    <div class="period-label">Hours Logged</div>
                    <div class="period-value accent">${totalLogged.toFixed(1)} h</div>
                </div>
                <div class="period-item">
                    <div class="period-label">Utilization</div>
                    <div class="period-value accent">${utilPct}%</div>
                </div>
                ${hasCost ? `
                <div class="period-item">
                    <div class="period-label">Total Mandays</div>
                    <div class="period-value accent">${totalManday.toFixed(2)}</div>
                </div>` : ''}
            </div>
        ` : '';

        // ── Summary chips ──────────────────────────────────────────────────
        document.getElementById('userWiseSummary').innerHTML = `
            <span class="summary-chip">${tasks.length} tasks / sub-tasks</span>
            <span class="summary-chip">${totalLogged.toFixed(1)} h logged</span>
            ${hasCost ? `<span class="summary-chip">${totalManday.toFixed(2)} mandays</span>` : ''}
        `;

        // ── Table — grouped by board ───────────────────────────────────────
        const byBoard = {};
        tasks.forEach(t => { (byBoard[t.project] ??= []).push(t); });

        let html = `<table>
            <thead><tr>
                <th style="text-align:left">Project / Board</th>
                <th>Sprint</th>
                <th style="text-align:left">Task</th>
                <th style="text-align:left">Sub Task</th>
                <th>Hours</th>
                <th title="Distinct weekdays with at least 1 log">Log Days</th>
                <th title="Total weekdays between Start and End">Total Days</th>
                ${hasCost ? '<th>Mandays</th>' : ''}
                <th>Start Log</th>
                <th>End Log</th>
                <th title="Weekdays with no log between Start and End">Break Days</th>
            </tr></thead>
            <tbody>`;

        Object.entries(byBoard).forEach(([board, bTasks]) => {
            // ── Board summary row ──────────────────────────────────────────
            const bHours   = bTasks.reduce((s, t) => s + t.hours, 0);
            const bManday  = bTasks.reduce((s, t) => s + t.manday, 0);
            const bLogDays = new Set(bTasks.flatMap(t => [...t.logDays]));
            const bLogWD   = weekdayCount(bLogDays);
            const bAllMins = bTasks.map(t => t.min).filter(Boolean);
            const bAllMaxs = bTasks.map(t => t.max).filter(Boolean);
            const bMin     = bAllMins.length ? new Date(Math.min(...bAllMins)) : null;
            const bMax     = bAllMaxs.length ? new Date(Math.max(...bAllMaxs)) : null;
            const bTotal   = bMin && bMax ? countWeekdays(bMin, bMax) : 0;
            const bBreak   = bMin && bMax ? countBreakDays(bMin, bMax, bLogDays) : 0;
            const bBreakCl = bBreak >= 10 ? 'break-high' : bBreak >= 5 ? 'break-mid' : '';

            html += `<tr class="board-summary-row">
                <td colspan="4">
                    <strong>${board}</strong>
                    <span class="board-count">${bTasks.length} task(s)</span>
                </td>
                <td class="right-cell"><strong>${bHours.toFixed(1)} h</strong></td>
                <td class="center-cell"><strong>${bLogWD}</strong></td>
                <td class="center-cell"><strong>${bTotal}</strong></td>
                ${hasCost ? `<td class="right-cell"><strong>${bManday.toFixed(2)}</strong></td>` : ''}
                <td class="center-cell">${fmtDate(bMin)}</td>
                <td class="center-cell">${fmtDate(bMax)}</td>
                <td class="center-cell ${bBreakCl}">${bBreak}</td>
            </tr>`;

            // ── Detail rows ────────────────────────────────────────────────
            bTasks.forEach(t => {
                const bkCls  = t.breakDays >= 10 ? 'break-high' : t.breakDays >= 5 ? 'break-mid' : '';
                const sprint = t.sprint ? xlsxHandler._fmtSprint(t.sprint) : '—';

                html += `<tr class="task-detail-row">
                    <td></td>
                    <td class="center-cell sprint-cell">${sprint}</td>
                    <td class="task-name-cell" title="${t.taskId}">${t.taskName}</td>
                    <td class="subtask-name-cell" title="${t.subTaskId}">
                        ${t.subTaskName || '<span class="zero-cell">—</span>'}
                    </td>
                    <td class="right-cell">${t.hours.toFixed(1)} h</td>
                    <td class="center-cell">${t.loggedDays}</td>
                    <td class="center-cell">${t.totalDays}</td>
                    ${hasCost ? `<td class="right-cell">${t.manday.toFixed(2)}</td>` : ''}
                    <td class="center-cell">${fmtDate(t.min)}</td>
                    <td class="center-cell">${fmtDate(t.max)}</td>
                    <td class="center-cell ${bkCls}">${t.breakDays}</td>
                </tr>`;
            });
        });

        html += '</tbody></table>';

        document.getElementById('userWiseTableWrap').innerHTML = html;
        document.getElementById('userWiseContent').style.display    = 'block';
        document.getElementById('userWiseEmptyState').style.display = 'none';
    }
}

export default new UserWiseHandler();
