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

window.getRecordIdentityAliases = function getRecordIdentityAliases(record) {
    if (!record) return [];

    const serverRecordID = String(record.serverRecordID || record.ServerRecordID || '').trim();
    const clientRecordID = String(record.clientRecordID || record.ClientRecordID || '').trim();

    const aliases = [];

    if (serverRecordID) aliases.push('srv:' + serverRecordID);
    if (clientRecordID) aliases.push('cli:' + clientRecordID);

    return aliases;
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
        record.receivedAt ||
        record.ReceivedAt ||
        ''
    );
};

window.isLocalPriorityRecord = function isLocalPriorityRecord(record) {
    if (!record) return false;

    const status = String(record.syncStatus || record.SyncStatus || '').toLowerCase();
    const error = String(record.lastSyncError || record.LastSyncError || '').trim();

    return status === 'pending' || status === 'syncing' || !!error;
};

window.pickPreferredRecordForRender = function pickPreferredRecordForRender(existing, candidate) {
    if (!existing) return candidate;
    if (!candidate) return existing;

    const existingPriority = window.isLocalPriorityRecord(existing);
    const candidatePriority = window.isLocalPriorityRecord(candidate);

    if (candidatePriority && !existingPriority) return candidate;
    if (existingPriority && !candidatePriority) return existing;

    const existingTs = window.getRecordFreshnessTs(existing);
    const candidateTs = window.getRecordFreshnessTs(candidate);

    if (candidateTs && existingTs) {
        return candidateTs >= existingTs ? candidate : existing;
    }

    if (candidateTs && !existingTs) return candidate;

    return existing;
};

window.dedupeRecordsForRender = function dedupeRecordsForRender(records, aliasesFn) {
    const getAliases = typeof aliasesFn === 'function'
        ? aliasesFn
        : window.getRecordIdentityAliases;

    const recordsByMainKey = new Map();
    const aliasToMainKey = new Map();

    let noKeyCounter = 0;

    (records || []).forEach(record => {
        if (!record) return;

        const aliases = (getAliases(record) || [])
            .map(a => String(a || '').trim())
            .filter(Boolean);

        if (!aliases.length) {
            recordsByMainKey.set('nokey:' + (++noKeyCounter), record);
            return;
        }

        let mainKey = '';

        for (const alias of aliases) {
            if (aliasToMainKey.has(alias)) {
                mainKey = aliasToMainKey.get(alias);
                break;
            }
        }

        if (!mainKey) {
            mainKey = aliases[0];
        }

        const existing = recordsByMainKey.get(mainKey);
        const preferred = window.pickPreferredRecordForRender(existing, record);

        recordsByMainKey.set(mainKey, preferred);

        // Važno: svi poznati aliasi sada pokazuju na isti canonical slot.
        aliases.forEach(alias => {
            aliasToMainKey.set(alias, mainKey);
        });

        // Ako je existing imao alias koji candidate nema, zadrži i njih.
        if (existing) {
            const existingAliases = (getAliases(existing) || [])
                .map(a => String(a || '').trim())
                .filter(Boolean);

            existingAliases.forEach(alias => {
                aliasToMainKey.set(alias, mainKey);
            });
        }
    });

    return Array.from(recordsByMainKey.values());
};
