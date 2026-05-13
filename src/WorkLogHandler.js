import { loadJSON } from "./loader.js";

class WorkLogHandler {

    constructor() {
        this.userInfos = null;
        this.roles = [];
        this.projects = null;
        this._listenEvents();
    }

    _listenEvents() {
        document.getElementById("costFile").addEventListener("change", async (e) => {
            const file = e.target.files[0];
            if (!file) return;
            document.getElementById("costFileName").textContent = file.name;
            const data = await loadJSON(file);
            if (!data) return;
            this.userInfos = this.getUsersInfo(data);
            this._setStatus("userGroup", `${Object.keys(this.userInfos).length} users loaded`);
            const studio = document.getElementById("studioInput").value.trim();
            if (studio) document.getElementById("studioName").textContent = studio;
            // Notify XlsxHandler in case XLSX was already loaded
            document.dispatchEvent(new Event('userInfosReady'));
        });

        // CSV listener removed — Cost Report is now computed from XLSX data

        document.getElementById("exportBoardBtn").addEventListener("click", () => {
            if (this.projects) this._download(this._buildBoardCSV(this.projects), "cost_per_board.csv");
        });
        document.getElementById("exportCategoryBtn").addEventListener("click", () => {
            if (this.projects) this._download(this._buildCategoryCSV(this.projects), "cost_by_category.csv");
        });
        document.getElementById("exportProjectBtn").addEventListener("click", () => {
            if (this.projects) this._download(this._buildProjectCSV(this.projects), "cost_per_project.csv");
        });
        document.getElementById("exportAllBtn").addEventListener("click", () => {
            if (this.projects) this.exportToCSV(this.projects);
        });

        document.querySelectorAll(".tab-btn").forEach(btn => {
            btn.addEventListener("click", () => this._switchTab(btn.dataset.tab));
        });
    }

    _switchTab(tabId) {
        document.querySelectorAll(".tab-btn").forEach(b => {
            b.classList.toggle("active", b.dataset.tab === tabId);
        });
        document.querySelectorAll(".tab-content").forEach(c => {
            c.classList.toggle("active", c.id === tabId);
        });
    }

    _setStatus(type, message) {
        const badge = document.getElementById(`status-${type}`);
        const text = document.getElementById(`status-${type}-text`);
        badge.classList.add("loaded");
        text.textContent = message;
    }

    getUsersInfo(data) {
        const users = {};
        const rolesSet = new Set();
        const domain = Object.keys(data.value.groups)[0];
        const groups = data.value.groups[domain];
        groups.forEach(group => {
            group.users.forEach(({ displayName, costPerHour }) => {
                const ID = +displayName.split("-")[0].trim();
                users[displayName] = { ID, name: displayName, role: group.name, cost: +costPerHour };
                rolesSet.add(group.name);
            });
        });
        this.roles = [...rolesSet];
        return users;
    }

    // ── XLSX-based Cost Report ─────────────────────────────────────────

    /**
     * Build the same `projects` map as getWorklogInfo() but directly from
     * XLSX "Flat (Groupable)" rows: { 'Project Name', 'Log user', 'Hr. Spent' }
     */
    getWorklogInfoFromXlsx(rows) {
        const projects = {};
        rows.forEach(r => {
            const boardName = r['Project Name'];
            if (!boardName) return;
            if (!projects[boardName]) {
                const entry = { Total: 0 };
                this.roles.forEach(role => (entry[role] = 0));
                projects[boardName] = entry;
            }
            const user = this.userInfos[r['Log user']];
            if (!user) return;
            const hours = +r['Hr. Spent'] || 0;
            const manday = hours * user.cost;
            projects[boardName][user.role] = (projects[boardName][user.role] || 0) + manday;
            projects[boardName].Total += manday;
        });
        return projects;
    }

    /** Called by XlsxHandler after XLSX is parsed (and userInfos is ready). */
    loadFromXlsx(rows) {
        if (!this.userInfos) return;
        this.projects = this.getWorklogInfoFromXlsx(rows);
        this._setStatus("xlsx", `${Object.keys(this.projects).length} boards · cost computed`);
        this._renderResults(this.projects);
        this._switchTab("tab-results");
    }

    getWorklogInfo(records) {
        const projects = {};

        // Build project list from header keys (records[0] is the sub-header row — skipped via user lookup)
        for (let key in records[0]) {
            if (!key) continue;
            if (key.match(/^_\d/) || key.match(/^__EMPTY/)) continue;
            if (["Group Name", "Cost", "Grand Total", "User Name"].includes(key)) continue;
            const entry = { Total: 0 };
            this.roles.forEach(r => (entry[r] = 0));
            projects[key] = entry;
        }

        records.forEach((record) => {
            const displayName = record["User Name"];
            const user = this.userInfos[displayName];
            if (!user) return;
            const { cost, role } = user;
            for (let key in projects) {
                const hours = +record[key] || 0;
                const manday = hours * cost;
                projects[key][role] += manday;
                projects[key].Total += manday;
            }
        });

        return projects;
    }

    _renderResults(projects) {
        const roles = this.roles;
        const keys = Object.keys(projects);

        // Compute grand totals for summary cards
        const totals = { Total: 0 };
        roles.forEach(r => (totals[r] = 0));
        keys.forEach(k => {
            totals.Total += projects[k].Total;
            roles.forEach(r => (totals[r] += projects[k][r]));
        });

        const cardColors = { ART: "#6366f1", ANIM: "#f59e0b", FE: "#10b981", GD: "#ef4444" };
        const cardsEl = document.getElementById("summaryCards");
        cardsEl.innerHTML = [
            `<div class="summary-card">
                <div class="label">Total Cost</div>
                <div class="value">${totals.Total.toFixed(2)}</div>
                <div class="sub">${keys.length} boards</div>
            </div>`,
            ...roles.map(r => `<div class="summary-card">
                <div class="label" style="color:${cardColors[r] || "#9ca3af"}">${r}</div>
                <div class="value">${totals[r].toFixed(2)}</div>
                <div class="sub">${totals.Total > 0 ? ((totals[r] / totals.Total) * 100).toFixed(1) : 0}% of total</div>
            </div>`)
        ].join("");
        cardsEl.style.display = "grid";

        // Sort: game projects (start with digits) first, then other boards
        const sortedKeys = [...keys].sort((a, b) => {
            const aIsGame = /^\d+/.test(a);
            const bIsGame = /^\d+/.test(b);
            if (aIsGame && !bIsGame) return -1;
            if (!aIsGame && bIsGame) return 1;
            return 0;
        });

        // Build table
        let html = `<table>
            <thead>
                <tr>
                    <th>Board / Project</th>
                    ${roles.map(r => `<th>${r}</th>`).join("")}
                    <th>Total</th>
                </tr>
            </thead>
            <tbody>`;

        sortedKeys.forEach(key => {
            const p = projects[key];
            html += `<tr>
                <td>${key}</td>
                ${roles.map(r => p[r] > 0
                    ? `<td>${p[r].toFixed(2)}</td>`
                    : `<td class="zero-cell">—</td>`
                ).join("")}
                <td class="total-cell">${p.Total.toFixed(2)}</td>
            </tr>`;
        });

        html += `<tr class="tfoot-row">
            <td>Grand Total</td>
            ${roles.map(r => `<td class="total-cell">${totals[r].toFixed(2)}</td>`).join("")}
            <td class="total-cell">${totals.Total.toFixed(2)}</td>
        </tr>`;

        html += `</tbody></table>`;

        const tableEl = document.getElementById("resultsTable");
        tableEl.innerHTML = html;
        tableEl.style.display = "block";

        document.getElementById("emptyState").style.display = "none";
        ["exportBoardBtn", "exportCategoryBtn", "exportProjectBtn", "exportAllBtn"].forEach(id => {
            document.getElementById(id).disabled = false;
        });
        document.getElementById("exportAllWrap").style.display = "flex";

        this._renderCategoryTable(projects);
        this._renderProjectTable(projects);
        this._renderChart(projects);
    }

    _renderChart(projects) {
        const roles = this.roles;
        const ROLE_COLORS = {
            ART:  { line: "#6366f1", fill: "rgba(99,102,241,0.08)" },
            ANIM: { line: "#f59e0b", fill: "rgba(245,158,11,0.08)" },
            FE:   { line: "#10b981", fill: "rgba(16,185,129,0.08)" },
            GD:   { line: "#ef4444", fill: "rgba(239,68,68,0.08)"  },
        };

        const gameKeys = Object.keys(projects)
            .filter(k => /^\d+/.test(k))
            .sort((a, b) => projects[b].Total - projects[a].Total);
        if (gameKeys.length === 0) return;

        const labels = gameKeys.map(k => k);

        const datasets = roles.map(role => {
            const color = ROLE_COLORS[role] || { line: "#9ca3af", fill: "rgba(156,163,175,0.08)" };
            return {
                label: role,
                data: gameKeys.map(k => +projects[k][role].toFixed(4)),
                borderColor: color.line,
                backgroundColor: color.fill,
                borderWidth: 2.5,
                pointRadius: 5,
                pointHoverRadius: 7,
                pointBackgroundColor: color.line,
                tension: 0.35,
                fill: false,
            };
        });

        // Total line — dashed
        datasets.push({
            label: "Total",
            data: gameKeys.map(k => +projects[k].Total.toFixed(4)),
            borderColor: "#1a1a2e",
            backgroundColor: "transparent",
            borderWidth: 2,
            borderDash: [6, 3],
            pointRadius: 4,
            pointHoverRadius: 6,
            pointBackgroundColor: "#1a1a2e",
            tension: 0.35,
            fill: false,
        });

        const ctx = document.getElementById("projectChart").getContext("2d");
        if (this._chart) this._chart.destroy();
        this._chart = new Chart(ctx, {
            type: "line",
            data: { labels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: "index", intersect: false },
                plugins: {
                    legend: {
                        position: "top",
                        labels: { boxWidth: 12, font: { size: 12 } },
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.y.toFixed(2)}`,
                        },
                    },
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 12 } },
                    },
                    y: {
                        beginAtZero: true,
                        grid: { color: "#f0f2f5" },
                        ticks: { font: { size: 11 } },
                    },
                },
            },
        });

        document.getElementById("chartSection").style.display = "block";
    }

    // Game projects → "Game Projects"; everything else → its own board name
    _classifyBoard(key) {
        return /^\d+/.test(key) ? "Game Projects" : key;
    }

    _renderCategoryTable(projects) {
        const roles = this.roles;
        const cats = {};
        const catOrder = [];

        for (let key in projects) {
            const cat = this._classifyBoard(key);
            if (!cats[cat]) {
                cats[cat] = { Total: 0 };
                roles.forEach(r => (cats[cat][r] = 0));
                catOrder.push(cat);
            }
            cats[cat].Total += projects[key].Total;
            roles.forEach(r => (cats[cat][r] += projects[key][r]));
        }

        // "Game Projects" always first
        const gameIdx = catOrder.indexOf("Game Projects");
        if (gameIdx > 0) { catOrder.splice(gameIdx, 1); catOrder.unshift("Game Projects"); }

        const grandTotal = catOrder.reduce((s, c) => s + cats[c].Total, 0);
        if (grandTotal === 0) return;

        const fmt = v => v.toFixed(2);
        const pct = v => grandTotal > 0 ? ((v / grandTotal) * 100).toFixed(1) + "%" : "—";

        let html = `<table>
            <thead>
                <tr>
                    <th>Category</th>
                    ${roles.map(r => `<th>${r}</th>`).join("")}
                    <th>Total</th>
                    <th>%</th>
                </tr>
            </thead>
            <tbody>`;

        catOrder.forEach(cat => {
            const c = cats[cat];
            html += `<tr>
                <td>${cat}</td>
                ${roles.map(r => c[r] > 0 ? `<td>${fmt(c[r])}</td>` : `<td class="zero-cell">—</td>`).join("")}
                <td class="total-cell">${fmt(c.Total)}</td>
                <td class="avg-col">${pct(c.Total)}</td>
            </tr>`;
        });

        const roleTotals = {};
        roles.forEach(r => (roleTotals[r] = catOrder.reduce((s, c) => s + cats[c][r], 0)));

        html += `<tr class="tfoot-row">
            <td>Total</td>
            ${roles.map(r => `<td class="total-cell">${fmt(roleTotals[r])}</td>`).join("")}
            <td class="total-cell">${fmt(grandTotal)}</td>
            <td class="avg-col">100%</td>
        </tr>`;

        html += `</tbody></table>`;

        document.getElementById("categoryTable").innerHTML = html;
        document.getElementById("categorySection").style.display = "block";
    }

    _renderProjectTable(projects) {
        const roles = this.roles;
        const gameKeys = Object.keys(projects)
            .filter(k => /^\d+/.test(k))
            .sort((a, b) => projects[b].Total - projects[a].Total);
        if (gameKeys.length === 0) return;

        // Average row
        const avg = { Total: 0 };
        roles.forEach(r => (avg[r] = 0));
        gameKeys.forEach(k => {
            avg.Total += projects[k].Total;
            roles.forEach(r => (avg[r] += projects[k][r]));
        });
        roles.forEach(r => (avg[r] /= gameKeys.length));
        avg.Total /= gameKeys.length;

        let html = `<table>
            <thead>
                <tr>
                    <th>Project</th>
                    ${roles.map(r => `<th>${r}</th>`).join("")}
                    <th>Total</th>
                </tr>
            </thead>
            <tbody>`;

        gameKeys.forEach(key => {
            const p = projects[key];
            html += `<tr>
                <td>${key}</td>
                ${roles.map(r => p[r] > 0
                    ? `<td>${p[r].toFixed(2)}</td>`
                    : `<td class="zero-cell">—</td>`
                ).join("")}
                <td class="total-cell">${p.Total.toFixed(2)}</td>
            </tr>`;
        });

        html += `<tr class="tfoot-row">
            <td class="avg-col">Average</td>
            ${roles.map(r => `<td class="avg-col">${avg[r].toFixed(2)}</td>`).join("")}
            <td class="avg-col">${avg.Total.toFixed(2)}</td>
        </tr>`;

        html += `</tbody></table>`;

        document.getElementById("projectTable").innerHTML = html;
        document.getElementById("projectSection").style.display = "block";
    }

    // ── CSV helpers ────────────────────────────────────────────────────

    _csvUtils() {
        const fmt = v => (+v).toFixed(4);
        const q   = v => `"${String(v).replace(/"/g, '""')}"`;
        const row = arr => arr.map(q).join(",");
        return { fmt, q, row };
    }

    _buildBoardCSV(projects) {
        const roles = this.roles;
        const { fmt, row } = this._csvUtils();
        const sortedKeys = Object.keys(projects).sort((a, b) => {
            const ag = /^\d+/.test(a), bg = /^\d+/.test(b);
            return ag === bg ? 0 : ag ? -1 : 1;
        });
        const totals = { Total: 0 };
        roles.forEach(r => (totals[r] = 0));
        sortedKeys.forEach(k => {
            totals.Total += projects[k].Total;
            roles.forEach(r => (totals[r] += projects[k][r]));
        });
        return [
            row(["Board / Project", ...roles, "Total"]),
            ...sortedKeys.map(k => row([k, ...roles.map(r => fmt(projects[k][r])), fmt(projects[k].Total)])),
            row(["Grand Total", ...roles.map(r => fmt(totals[r])), fmt(totals.Total)]),
        ].join("\n");
    }

    _buildCategoryCSV(projects) {
        const roles = this.roles;
        const { fmt, row } = this._csvUtils();
        const cats = {}, catOrder = [];
        for (let key in projects) {
            const cat = this._classifyBoard(key);
            if (!cats[cat]) { cats[cat] = { Total: 0 }; roles.forEach(r => (cats[cat][r] = 0)); catOrder.push(cat); }
            cats[cat].Total += projects[key].Total;
            roles.forEach(r => (cats[cat][r] += projects[key][r]));
        }
        const gi = catOrder.indexOf("Game Projects");
        if (gi > 0) { catOrder.splice(gi, 1); catOrder.unshift("Game Projects"); }
        const total = catOrder.reduce((s, c) => s + cats[c].Total, 0);
        const pct = v => total > 0 ? ((v / total) * 100).toFixed(2) + "%" : "0%";
        const roleTotals = {};
        roles.forEach(r => (roleTotals[r] = catOrder.reduce((s, c) => s + cats[c][r], 0)));
        return [
            row(["Category", ...roles, "Total", "%"]),
            ...catOrder.map(c => row([c, ...roles.map(r => fmt(cats[c][r])), fmt(cats[c].Total), pct(cats[c].Total)])),
            row(["Total", ...roles.map(r => fmt(roleTotals[r])), fmt(total), "100%"]),
        ].join("\n");
    }

    _buildProjectCSV(projects) {
        const roles = this.roles;
        const { fmt, row } = this._csvUtils();
        const gameKeys = Object.keys(projects)
            .filter(k => /^\d+/.test(k))
            .sort((a, b) => projects[b].Total - projects[a].Total);
        if (gameKeys.length === 0) return "";
        const avg = { Total: 0 };
        roles.forEach(r => (avg[r] = 0));
        gameKeys.forEach(k => { avg.Total += projects[k].Total; roles.forEach(r => (avg[r] += projects[k][r])); });
        roles.forEach(r => (avg[r] /= gameKeys.length));
        avg.Total /= gameKeys.length;
        return [
            row(["Project", ...roles, "Total"]),
            ...gameKeys.map(k => row([k, ...roles.map(r => fmt(projects[k][r])), fmt(projects[k].Total)])),
            row(["Average", ...roles.map(r => fmt(avg[r])), fmt(avg.Total)]),
        ].join("\n");
    }

    _download(csvContent, filename) {
        const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    exportToCSV(projects) {
        const sections = [
            "=== Cost per Board ===\n"    + this._buildBoardCSV(projects),
            "=== Cost by Category ===\n"  + this._buildCategoryCSV(projects),
            "=== Cost per Game Project ===\n" + this._buildProjectCSV(projects),
        ];
        this._download(sections.join("\n\n"), "cost_report_all.csv");
    }
}

export default new WorkLogHandler();
