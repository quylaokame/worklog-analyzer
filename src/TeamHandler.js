/**
 * TeamHandler — per-team (ART / ANIM / FE / GD) tabs.
 * For each member it renders a bar chart: one bar per task (hours logged),
 * bars coloured & grouped by board. Members are ordered by level
 * (fresher → senior, i.e. costPerHour ascending), tie-broken by ID.
 *
 * Requires xlsxHandler.flatGroupable + workLogHandler.userInfos for live data,
 * but can also rehydrate from a saved payload (Firebase / IndexedDB restore).
 */
import xlsxHandler    from './XlsxHandler.js';
import workLogHandler from './WorkLogHandler.js';

const ROLES = ['ART', 'ANIM', 'FE', 'GD'];

const BOARD_PALETTE = [
    '#6366f1', '#f59e0b', '#10b981', '#ef4444', '#0ea5e9', '#8b5cf6',
    '#ec4899', '#14b8a6', '#f97316', '#84cc16', '#a855f7', '#06b6d4',
];

/**
 * Level cost reference, expressed per working day (8h). The .jab stores
 * costPerHour, so we compare against costPerHour * 8. FE uses a higher scale
 * than the other teams. Matching is nearest-value to tolerate rounding
 * (e.g. 0.0938 * 8 = 0.7504 ≈ 0.75).
 */
const LEVELS_PER_DAY = {
    DEFAULT: [['Fresher', 0.25], ['Junior', 0.5],  ['Middle', 1.0],  ['Senior', 1.5]],
    FE:      [['Fresher', 0.5],  ['Junior', 0.75], ['Middle', 1.25], ['Senior', 1.75]],
};

/** Map costPerHour → level, role-aware. */
function levelOf(cost, role) {
    const perDay = cost * 8;
    const table  = LEVELS_PER_DAY[role] || LEVELS_PER_DAY.DEFAULT;
    let best = table[0][0], bestDiff = Infinity;
    for (const [level, ref] of table) {
        const diff = Math.abs(perDay - ref);
        if (diff < bestDiff) { bestDiff = diff; best = level; }
    }
    return best;
}

/** Detect the summary/title column from the first XLSX row. */
function detectSummaryAccessor(rows) {
    if (!rows?.length) return () => '';
    const keys = Object.keys(rows[0]);
    const exact = ['Issue Summary', 'Summary', 'Ticket Summary',
                   'Issue Title', 'Title', 'Work Description', 'Description'];
    const hit = exact.find(k => keys.includes(k));
    if (hit) return r => r[hit] || '';
    const fuzzy = keys.find(k => /summary|title/i.test(k));
    return fuzzy ? r => r[fuzzy] || '' : () => '';
}

class TeamHandler {

    constructor() {
        this._data       = null;   // [{ name, id, role, cost, level, total, boards }]
        this._charts     = {};     // role → [Chart]
        this._chartState = {};     // canvasId → { member, colors, hidden, boardOf, chart }
        this._initEvents();
    }

    _initEvents() {
        document.addEventListener('reportReady', () => this._onDataReady());

        document.querySelectorAll('.tab-btn').forEach(btn => {
            const tab = btn.dataset.tab || '';
            if (tab.startsWith('tab-team-')) {
                btn.addEventListener('click', () => this.renderTeam(tab.replace('tab-team-', '')));
            }
        });
    }

    _onDataReady() {
        if (!xlsxHandler.flatGroupable || !workLogHandler.userInfos) return;
        this._data = this._compute();
        const role = this._activeRole();
        if (role) this.renderTeam(role);
    }

    _activeRole() {
        const active = document.querySelector('.tab-content.team-tab.active');
        return active ? active.dataset.role : null;
    }

    // ── Compute ──────────────────────────────────────────────────────────────

    _compute() {
        const rows       = xlsxHandler.flatGroupable;
        const userInfos  = workLogHandler.userInfos;
        const getSummary = detectSummaryAccessor(rows);

        // ticket → summary (resolved from any row)
        const summaryMap = {};
        rows.forEach(r => {
            const t = r['Ticket No'] || r['Issue Key'];
            const s = getSummary(r);
            if (t && s && !summaryMap[t]) summaryMap[t] = s;
        });

        // user → board → ticket → hours
        const agg = {};
        rows.forEach(r => {
            const name = r['Log user'];
            if (!name || !userInfos[name]) return;          // only known team members
            const board  = r['Project Name'] || '(No board)';
            const ticket = r['Ticket No'] || r['Issue Key'];
            if (!ticket) return;
            const hours = +r['Hr. Spent'] || 0;
            ((agg[name] ??= {})[board] ??= {});
            const t = (agg[name][board][ticket] ??= { name: summaryMap[ticket] || ticket, hours: 0 });
            t.hours += hours;
        });

        return Object.keys(agg).map(name => {
            const info = userInfos[name];
            const boards = Object.entries(agg[name]).map(([board, tickets]) => {
                const tasks = Object.values(tickets).sort((a, b) => b.hours - a.hours);
                const total = tasks.reduce((s, t) => s + t.hours, 0);
                return { board, total, tasks };
            }).sort((a, b) => b.total - a.total);
            const total = boards.reduce((s, b) => s + b.total, 0);
            return { name, id: info.ID, role: info.role, cost: info.cost, level: levelOf(info.cost, info.role), total, boards };
        });
    }

    // ── Render ───────────────────────────────────────────────────────────────

    renderTeam(role) {
        const container = document.getElementById(`tab-team-${role}`);
        if (!container) return;

        this._destroyCharts(role);

        if (!this._data) {
            container.innerHTML = this._empty('Load <strong>User Group</strong> + <strong>Worklog XLSX</strong> to see this team.');
            return;
        }

        const members = this._data
            .filter(u => u.role === role)
            .sort((a, b) => (a.cost - b.cost) || (b.id - a.id));   // level asc, then ID high → low

        if (!members.length) {
            container.innerHTML = this._empty('No members with logged work in this team.');
            return;
        }

        const esc = s => String(s).replace(/"/g, '&quot;');

        let html = `<div class="section-title" style="margin:8px 0 16px">${role} Team · ${members.length} member(s)</div>`;
        members.forEach((m, idx) => {
            const canvasId = `teamChart-${role}-${idx}`;
            const colors = this._boardColors(m.boards.map(b => b.board));
            const legend = m.boards.map(b =>
                `<span class="lg-item" data-canvas="${canvasId}" data-board="${esc(b.board)}" title="Click to show/hide">
                    <span class="lg-dot" style="background:${colors[b.board]}"></span>${b.board} (${(b.total / 8).toFixed(2)}d)
                </span>`
            ).join('');
            const taskCount = m.boards.reduce((s, b) => s + b.tasks.length, 0);
            html += `<div class="member-card">
                <div class="member-head">
                    <span class="m-name">${m.name}</span>
                    <span class="level-badge level-${m.level.toLowerCase()}">${m.level}</span>
                    <span class="m-stats">${(m.total / 8).toFixed(2)} d · ${taskCount} task(s) · ${m.boards.length} board(s)</span>
                </div>
                <div class="member-body">
                    <div class="team-legend">${legend}</div>
                    <div class="member-chart-wrap"><canvas id="${canvasId}"></canvas></div>
                </div>
            </div>`;
        });
        container.innerHTML = html;

        this._charts[role] = members.map((m, idx) => this._makeChart(`teamChart-${role}-${idx}`, m));

        // Clickable board legend → toggle that board's bars (ECharts-style)
        container.querySelectorAll('.lg-item').forEach(chip => {
            chip.addEventListener('click', () =>
                this._toggleBoard(chip.dataset.canvas, chip.dataset.board, chip));
        });
    }

    _makeChart(canvasId, m) {
        const colors = this._boardColors(m.boards.map(b => b.board));
        const state  = { member: m, colors, hidden: new Set(), boardOf: [], chart: null };
        this._chartState[canvasId] = state;

        const { labels, data, bg } = this._buildSeries(state);

        const ctx = document.getElementById(canvasId).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'bar',
            data: { labels, datasets: [{ label: 'Hours', data, backgroundColor: bg, borderWidth: 0, borderRadius: 3 }] },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            title: items => `[${state.boardOf[items[0].dataIndex]}] ${items[0].label}`,
                            label: c => ` ${c.parsed.y.toFixed(2)} d`,
                        },
                    },
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: {
                            autoSkip: false, maxRotation: 90, minRotation: 45,
                            font: { size: 9 },
                            callback(v) { const s = this.getLabelForValue(v); return s.length > 18 ? s.slice(0, 17) + '…' : s; },
                        },
                    },
                    y: {
                        beginAtZero: true,
                        grid: { color: '#f0f2f5' },
                        ticks: { font: { size: 11 } },
                        title: { display: true, text: 'Man-days', font: { size: 11 } },
                    },
                },
            },
        });

        state.chart = chart;
        return chart;
    }

    /** Build label/data/color arrays from a chart state, skipping hidden boards. */
    _buildSeries(state) {
        const labels = [], data = [], bg = [];
        state.boardOf = [];
        state.member.boards.forEach(b => {
            if (state.hidden.has(b.board)) return;
            b.tasks.forEach(t => {
                labels.push(t.name);
                data.push(+(t.hours / 8).toFixed(2));   // hours → man-days
                bg.push(state.colors[b.board]);
                state.boardOf.push(b.board);
            });
        });
        return { labels, data, bg };
    }

    /** Toggle one board's bars on/off for a given chart. */
    _toggleBoard(canvasId, board, chipEl) {
        const state = this._chartState[canvasId];
        if (!state) return;

        if (state.hidden.has(board)) state.hidden.delete(board);
        else                          state.hidden.add(board);
        chipEl.classList.toggle('lg-off', state.hidden.has(board));

        const { labels, data, bg } = this._buildSeries(state);
        state.chart.data.labels = labels;
        state.chart.data.datasets[0].data = data;
        state.chart.data.datasets[0].backgroundColor = bg;
        state.chart.update();
    }

    _boardColors(boards) {
        const map = {};
        [...new Set(boards)].forEach((b, i) => (map[b] = BOARD_PALETTE[i % BOARD_PALETTE.length]));
        return map;
    }

    _destroyCharts(role) {
        (this._charts[role] || []).forEach(c => { try { c.destroy(); } catch {} });
        delete this._charts[role];
        // Drop stale per-canvas state for this role
        Object.keys(this._chartState)
            .filter(id => id.startsWith(`teamChart-${role}-`))
            .forEach(id => delete this._chartState[id]);
    }

    _empty(msg) {
        return `<div class="empty-state">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                <path d="M3 3v18h18"/><rect x="7" y="10" width="3" height="8"/><rect x="12" y="6" width="3" height="12"/><rect x="17" y="13" width="3" height="5"/>
            </svg>
            <p>${msg}</p>
        </div>`;
    }

    // ── Persistence ──────────────────────────────────────────────────────────

    buildPayload() {
        if (!this._data) return null;
        return { users: this._data };           // plain, Firebase-safe (no dynamic keys)
    }

    restoreFromData(payload) {
        if (!payload?.users) return;
        this._data = payload.users;
        const role = this._activeRole();
        if (role) this.renderTeam(role);
    }
}

export default new TeamHandler();
