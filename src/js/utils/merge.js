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

window.getRecordDedupeKey = function getRecordDedupeKey(record) {
    if (!record) return '';

    const serverRecordID = String(record.serverRecordID || record.ServerRecordID || '').trim();
    if (serverRecordID) return 'srv:' + serverRecordID;

    const clientRecordID = String(record.clientRecordID || record.ClientRecordID || '').trim();
    if (clientRecordID) return 'cli:' + clientRecordID;

    return '';
};

window.getRecordFreshnessTs = function getRecordFreshnessTs(record) {
    if (!record) return '';

    return String(
        record.updatedAtServer ||
        record.UpdatedAtServer ||
        record.syncedAt ||
        record.SyncedAt ||
        record.updatedAtClient ||
        record.UpdatedAtClient ||
        record.createdAtClient ||
        record.CreatedAtClient ||
        record.ReceivedAt ||
        record.receivedAt ||
        ''
    );
};

window.isLocalPriorityRecord = function isLocalPriorityRecord(record) {
    if (!record) return false;

    const status = String(record.syncStatus || record.SyncStatus || '').toLowerCase();
    const err = String(record.lastSyncError || record.LastSyncError || '').trim();

    return status === 'pending' || status === 'syncing' || !!err;
};

window.pickPreferredRecordForRender = function pickPreferredRecordForRender(existing, candidate) {
    if (!existing) return candidate;
    if (!candidate) return existing;

    const existingPriority = window.isLocalPriorityRecord(existing);
    const candidatePriority = window.isLocalPriorityRecord(candidate);

    if (candidatePriority && !existingPriority) return candidate;
    if (existingPriority && !candidatePriority) return existing;

    const candidateTs = window.getRecordFreshnessTs(candidate);
    const existingTs = window.getRecordFreshnessTs(existing);

    if (candidateTs && existingTs) {
        return candidateTs >= existingTs ? candidate : existing;
    }

    if (candidateTs && !existingTs) return candidate;

    return existing;
};

window.dedupeRecordsForRender = function dedupeRecordsForRender(records, keyFn) {
    const out = new Map();
    const getKey = typeof keyFn === 'function' ? keyFn : window.getRecordDedupeKey;

    (records || []).forEach(record => {
        if (!record) return;

        const key = getKey(record);
        if (!key) {
            // Bez ključa ne dedupujemo, da ne sakrijemo legitimno različite zapise.
            out.set('nokey:' + out.size, record);
            return;
        }

        out.set(
            key,
            window.pickPreferredRecordForRender(out.get(key), record)
        );
    });

    return Array.from(out.values());
};
