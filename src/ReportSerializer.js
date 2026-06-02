/**
 * ReportSerializer — shared serialize / restore logic used by both
 * FirebaseHandler (remote) and LocalDbHandler (IndexedDB).
 */
import workLogHandler from './WorkLogHandler.js';
import xlsxHandler    from './XlsxHandler.js';

// ── Serialize ────────────────────────────────────────────────────────────────

export function buildPayload(studio) {
    if (!xlsxHandler._projectsData) return null;

    // Convert Date objects → ISO strings for storage
    const overview = xlsxHandler._projectsData.map(p => ({
        ...p,
        sprintDates: Object.fromEntries(
            Object.entries(p.sprintDates).map(([k, v]) => [
                k,
                { min: v.min?.toISOString?.() ?? v.min,
                  max: v.max?.toISOString?.() ?? v.max },
            ])
        ),
    }));

    const phases = {};
    xlsxHandler._projectsData.forEach(({ project }) => {
        phases[project] = xlsxHandler._computePhases(project).map(p => ({
            ...p,
            min: p.min?.toISOString?.() ?? p.min,
            max: p.max?.toISOString?.() ?? p.max,
        }));
    });

    return {
        savedAt:    new Date().toISOString(),
        studio,
        allSprints: xlsxHandler._allSprints || [],
        overview,
        phases,
        costReport: workLogHandler.projects || null,
        roles:      workLogHandler.roles    || [],
    };
}

// ── Restore ──────────────────────────────────────────────────────────────────

export function restoreFromPayload(payload) {
    // Deserialize ISO strings → Date objects
    const overview = (payload.overview || []).map(p => ({
        ...p,
        sprintDates: Object.fromEntries(
            Object.entries(p.sprintDates || {}).map(([k, v]) => [
                k,
                { min: new Date(v.min), max: new Date(v.max) },
            ])
        ),
    }));

    const phases = {};
    Object.entries(payload.phases || {}).forEach(([proj, arr]) => {
        phases[proj] = arr.map(p => ({
            ...p,
            min: new Date(p.min),
            max: new Date(p.max),
        }));
    });

    xlsxHandler.restoreFromData(overview, payload.allSprints || [], phases);

    if (payload.costReport && payload.roles?.length) {
        workLogHandler.restoreFromData(payload.costReport, payload.roles);
    }
}
