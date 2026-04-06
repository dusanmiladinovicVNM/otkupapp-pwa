// Novi fajl: utils/merge.js

/**
 * Generički offline-first merge.
 *
 * @param {Array} local         - records iz IndexedDB
 * @param {Array} server        - records sa servera (već normalizovani u caller-u)
 * @param {Function} normalizeLocal - fn(record) → normalizovan oblik istog shape-a kao server
 * @param {string} [primaryKey]  - keyPath, default 'clientRecordID'
 * @returns {Array} merged records
 *
 * Pravila:
 *  1. Server records su baza
 *  2. Lokalni pending/syncing UVEK prepisuje server
 *  3. Lokalni synced prepisuje server samo ako je updatedAtClient noviji od serverUpdated
 *  4. Lokalni bez server match-a se dodaje
 */
window.mergeOfflineRecords = function (local, server, normalizeLocal, primaryKey) {
    const pk = primaryKey || 'clientRecordID';
    const merged = new Map();

    // 1. Server snapshot kao baza
    (server || []).forEach(r => {
        if (r && r[pk]) merged.set(r[pk], r);
    });

    // 2. Lokalni overlay
    (local || []).forEach(r => {
        if (!r || !r[pk]) return;

        const localNorm = typeof normalizeLocal === 'function' ? normalizeLocal(r) : r;
        const existing = merged.get(localNorm[pk]);

        // Nema na serveru — dodaj
        if (!existing) {
            merged.set(localNorm[pk], localNorm);
            return;
        }

        // Pending/syncing uvek ima prioritet
        if (localNorm.syncStatus === 'pending' || localNorm.syncStatus === 'syncing') {
            merged.set(localNorm[pk], localNorm);
            return;
        }

        // Synced vs synced — noviji pobeđuje
        const localUpdated = localNorm.updatedAtClient || localNorm.createdAtClient || '';
        const serverUpdated = existing.updatedAtServer || existing.updatedAtClient || existing.createdAtClient || '';

        if (localUpdated && serverUpdated && localUpdated > serverUpdated) {
            merged.set(localNorm[pk], localNorm);
        }
    });

    return Array.from(merged.values());
};
