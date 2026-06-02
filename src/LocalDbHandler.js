/**
 * LocalDbHandler — auto-save / restore reports using IndexedDB (via Dexie.js).
 * Dexie is loaded as a global via CDN in index.html.
 *
 * DB schema:
 *   reports: { id (auto), studio, date (YYYY-MM-DD), savedAt (ISO), payload }
 */
import { buildPayload, restoreFromPayload } from './ReportSerializer.js';
import xlsxHandler from './XlsxHandler.js';

/** Find the earliest and latest log dates across all projects/sprints in the overview. */
function _computeReportRange(overview) {
    let min = null, max = null;
    (overview || []).forEach(p => {
        Object.values(p.sprintDates || {}).forEach(({ min: a, max: b }) => {
            const dA = new Date(a), dB = new Date(b);
            if (!min || dA < min) min = dA;
            if (!max || dB > max) max = dB;
        });
    });
    if (!min || !max) return null;
    const fmt = d => d.toLocaleString('en-US', { month: 'short', year: 'numeric' });
    return `${fmt(min)} – ${fmt(max)}`;
}

// ── Dexie database definition ────────────────────────────────────────────────
const db = new Dexie('worklog-analyzer');
db.version(1).stores({
    reports: '++id, studio, date',
});

class LocalDbHandler {

    constructor() {
        this._saveTimer = null;
        this._initEvents();
        // On page reload, studio is already restored from localStorage by FirebaseHandler
        // so we can immediately populate the history list
        this._refreshHistory();
    }

    // ── Events ─────────────────────────────────────────────────────────────

    _initEvents() {
        document.getElementById('studioSelect')
            .addEventListener('change', () => this._refreshHistory());

        document.getElementById('localLoadBtn')
            .addEventListener('click', () => this._loadSelected());

        document.getElementById('localDeleteBtn')
            .addEventListener('click', () => this._deleteSelected());

        // Auto-save whenever a report is fully rendered
        document.addEventListener('reportReady', () => this._scheduleAutoSave());
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    get _studio() { return document.getElementById('studioSelect').value; }

    _setStatus(msg) {
        const badge = document.getElementById('status-localdb');
        const text  = document.getElementById('status-localdb-text');
        if (badge) badge.classList.add('loaded');
        if (text)  text.textContent = msg;
    }

    // ── Auto-save ──────────────────────────────────────────────────────────

    /** Debounce: avoid double-save when XLSX + .jab both fire reportReady. */
    _scheduleAutoSave() {
        clearTimeout(this._saveTimer);
        this._saveTimer = setTimeout(() => this._autoSave(), 600);
    }

    async _autoSave() {
        // Only save when we have live XLSX data (not a Firebase restore)
        if (!this._studio || !xlsxHandler.flatGroupable) return;

        const payload = buildPayload(this._studio);
        if (!payload) return;

        const date        = new Date().toISOString().slice(0, 10);
        const studioLabel = document.getElementById('studioSelect').selectedOptions[0]?.text || this._studio;
        const reportRange = _computeReportRange(payload.overview);

        // Upsert: one record per studio + date
        const existing = await db.reports
            .where({ studio: this._studio, date })
            .first();

        const meta = { savedAt: payload.savedAt, studioLabel, reportRange, payload };

        if (existing) {
            await db.reports.update(existing.id, meta);
        } else {
            await db.reports.add({ studio: this._studio, date, ...meta });
        }

        this._setStatus(`Auto-saved · ${date}`);
        await this._refreshHistory();
    }

    // ── History ────────────────────────────────────────────────────────────

    async _refreshHistory() {
        if (!this._studio) {
            document.getElementById('localHistorySection').style.display = 'none';
            return;
        }

        // Newest first
        const records = await db.reports
            .where('studio').equals(this._studio)
            .reverse()
            .sortBy('date');

        const sel = document.getElementById('localDateSelect');

        if (records.length) {
            sel.innerHTML = records.map(r => {
                const time = new Date(r.savedAt).toLocaleString('vi-VN', {
                    hour: '2-digit', minute: '2-digit',
                });
                const studio = r.studioLabel || r.studio;
                const range  = r.reportRange ? ` · ${r.reportRange}` : '';
                return `<option value="${r.id}">${studio}${range} · saved ${r.date} ${time}</option>`;
            }).join('');

            document.getElementById('localHistorySection').style.display = 'flex';
            this._setStatus(`${records.length} report(s) saved locally`);
        } else {
            sel.innerHTML = '<option value="">No local reports</option>';
            document.getElementById('localHistorySection').style.display = 'none';
            this._setStatus('No local history yet');
        }
    }

    // ── Load / Delete ──────────────────────────────────────────────────────

    async _loadSelected() {
        const id = +document.getElementById('localDateSelect').value;
        if (!id) return;

        const btn = document.getElementById('localLoadBtn');
        btn.disabled = true; btn.textContent = 'Loading…';

        try {
            const record = await db.reports.get(id);
            if (!record) throw new Error('Record not found');
            restoreFromPayload(record.payload);
            this._setStatus(`Loaded: ${record.date}`);
        } catch (e) {
            alert(`Load failed: ${e.message}`);
        } finally {
            btn.disabled = false; btn.textContent = '⬇ Load';
        }
    }

    async _deleteSelected() {
        const id = +document.getElementById('localDateSelect').value;
        if (!id) return;
        if (!confirm('Delete this local report?')) return;
        await db.reports.delete(id);
        await this._refreshHistory();
    }
}

export default new LocalDbHandler();
