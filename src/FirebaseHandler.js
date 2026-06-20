import { buildPayload, restoreFromPayload } from './ReportSerializer.js';

// ── Studio list ────────────────────────────────────────────────────────────
const STUDIOS = [
    { label: 'Stream',    slug: 'stream'    },
    { label: 'Kame House', slug: 'kame-house' },
    { label: 'Pandoria',  slug: 'pandoria'  },
    { label: 'Slayteria', slug: 'slayteria' },
    { label: 'Nova',      slug: 'nova'      },
    { label: 'AlphaCrew', slug: 'alphacrew' },
    { label: 'OneForce',  slug: 'oneforce'  },
    { label: 'MiloSpace', slug: 'milospace' },
    { label: 'DeepSea',   slug: 'deepsea'   },
    { label: 'Eighteeen', slug: 'eighteeen' },
    { label: 'Apek',      slug: 'apek'      },
    { label: 'Sky',       slug: 'sky'       },
];

class FirebaseHandler {

    constructor() {
        this.dbUrl = (localStorage.getItem('fbDbUrl') || '').replace(/\/$/, '');
        this._populateSelect();
        this._initEvents();
    }

    // ── Setup ──────────────────────────────────────────────────────────────

    _populateSelect() {
        const sel = document.getElementById('studioSelect');
        STUDIOS.forEach(s => {
            const opt = document.createElement('option');
            opt.value = s.slug;
            opt.textContent = s.label;
            sel.appendChild(opt);
        });
        // Restore previously selected studio
        const saved = localStorage.getItem('fbStudio');
        if (saved) sel.value = saved;
        this._updateHeader();

        // Auto-reconnect and fetch date list if credentials were saved
        if (this.dbUrl && this._studio) {
            this._loadDateList();
            this._setStatus('Reconnected ✓');
        }
    }

    _initEvents() {
        // Studio change → update header + load date list
        document.getElementById('studioSelect').addEventListener('change', () => {
            localStorage.setItem('fbStudio', this._studio);
            this._updateHeader();
            if (this.dbUrl) this._loadDateList();
        });

        // DB URL input
        const urlInput = document.getElementById('firebaseUrl');
        if (this.dbUrl) urlInput.value = this.dbUrl;
        urlInput.addEventListener('change', () => {
            this.dbUrl = urlInput.value.trim().replace(/\/$/, '');
            localStorage.setItem('fbDbUrl', this.dbUrl);
        });

        document.getElementById('firebaseConnectBtn').addEventListener('click', () => this._testConnection());
        document.getElementById('firebaseSaveBtn').addEventListener('click',    () => this._save());
        document.getElementById('firebaseLoadBtn').addEventListener('click',    () => this._load());
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    get _studio() { return document.getElementById('studioSelect').value; }

    _updateHeader() {
        const studio = STUDIOS.find(s => s.slug === this._studio);
        document.getElementById('studioName').textContent =
            studio ? studio.label : 'Worklog Analyzer';
    }

    _setStatus(msg, ok = true) {
        const badge = document.getElementById('status-firebase');
        const text  = document.getElementById('status-firebase-text');
        badge.classList.toggle('loaded', ok);
        badge.classList.toggle('error',  !ok);
        text.textContent = msg;
    }

    // ── Firebase REST ──────────────────────────────────────────────────────

    async _testConnection() {
        if (!this.dbUrl) { alert('Please enter the Firebase Database URL.'); return; }
        const btn = document.getElementById('firebaseConnectBtn');
        btn.disabled = true; btn.textContent = 'Connecting…';
        try {
            const resp = await fetch(`${this.dbUrl}/.json?shallow=true`);
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            this._setStatus('Connected ✓');
            if (this._studio) this._loadDateList();
        } catch (e) {
            this._setStatus(`Error: ${e.message}`, false);
        } finally {
            btn.disabled = false; btn.textContent = 'Connect';
        }
    }

    async _loadDateList() {
        if (!this.dbUrl || !this._studio) return;
        try {
            const resp = await fetch(`${this.dbUrl}/meta/${this._studio}.json`);
            const data = resp.ok ? await resp.json() : null;

            // Newest first (savedAt is ISO, falls back to the key)
            const entries = data ? Object.entries(data) : [];
            entries.sort((a, b) =>
                (b[1].savedAt || b[0]).localeCompare(a[1].savedAt || a[0]));

            const sel = document.getElementById('firebaseDateSelect');
            sel.innerHTML = entries.length
                ? entries.map(([key, m]) => `<option value="${key}">${this._buildLabel(m)}</option>`).join('')
                : '<option value="">No reports saved yet</option>';

            document.getElementById('firebaseLoadSection').style.display =
                entries.length ? 'flex' : 'none';
        } catch (e) {
            console.warn('Firebase list error:', e);
        }
    }

    // ── Save ───────────────────────────────────────────────────────────────

    async _save() {
        if (!this.dbUrl) { alert('Connect to Firebase first.'); return; }
        if (!this._studio) { alert('Select a studio first.'); return; }

        const payload = buildPayload(this._studio);
        if (!payload) { alert('No data to save — load an XLSX file first.'); return; }

        const btn = document.getElementById('firebaseSaveBtn');
        btn.disabled = true; btn.textContent = 'Saving…';

        try {
            // Unique, Firebase-safe key per save (':' and '.' are forbidden in keys)
            const key = payload.savedAt.replace(/[:.]/g, '-');
            const studioLabel = document.getElementById('studioSelect').selectedOptions[0]?.text || this._studio;
            const meta = {
                studioLabel,
                savedAt: payload.savedAt,
                start:   payload.logRange?.start || '',
                end:     payload.logRange?.end   || '',
            };

            // Store the full report and a lightweight meta entry (for the list).
            const put = (path, body) => fetch(`${this.dbUrl}/${path}.json`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });

            const [r1, r2] = await Promise.all([
                put(`worklog/${this._studio}/${key}`, payload),
                put(`meta/${this._studio}/${key}`, meta),
            ]);
            if (!r1.ok) throw new Error(`HTTP ${r1.status}`);
            if (!r2.ok) throw new Error(`HTTP ${r2.status}`);

            this._setStatus(`Saved: ${this._fmtSaved(meta.savedAt)}`);
            this._loadDateList();
        } catch (e) {
            alert(`Save failed: ${e.message}`);
            this._setStatus(`Save failed: ${e.message}`, false);
        } finally {
            btn.disabled = false; btn.textContent = '💾 Save to Firebase';
        }
    }

    // ── Label helpers ────────────────────────────────────────────────────────

    _fmtSaved(iso) {
        const d = new Date(iso);
        if (isNaN(d)) return iso || '';
        const date = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
        const time = d.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });
        return `${date} ${time}`;
    }

    /** "Studio · start → end · saved dd/mm/yyyy hh:mm" */
    _buildLabel(m) {
        const studio = m.studioLabel || this._studio;
        const range  = (m.start || m.end) ? ` · ${m.start || '?'} → ${m.end || '?'}` : '';
        return `${studio}${range} · saved ${this._fmtSaved(m.savedAt)}`;
    }

    // ── Load ───────────────────────────────────────────────────────────────

    async _load() {
        if (!this.dbUrl || !this._studio) return;
        const key = document.getElementById('firebaseDateSelect').value;
        if (!key) return;

        const btn = document.getElementById('firebaseLoadBtn');
        btn.disabled = true; btn.textContent = 'Loading…';

        try {
            const resp = await fetch(
                `${this.dbUrl}/worklog/${this._studio}/${key}.json`
            );
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const payload = await resp.json();
            if (!payload) throw new Error('Empty report');

            restoreFromPayload(payload);
            this._setStatus(`Loaded: ${this._fmtSaved(payload.savedAt)}`);
        } catch (e) {
            alert(`Load failed: ${e.message}`);
            this._setStatus(`Load failed: ${e.message}`, false);
        } finally {
            btn.disabled = false; btn.textContent = '⬇ Load';
        }
    }

}

export default new FirebaseHandler();
