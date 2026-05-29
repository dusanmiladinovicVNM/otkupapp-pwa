// ============================================================
// VOZAC: STATUS + PROFIL
// ============================================================
const VOZAC_STATUS_KEY = 'vozacStatus';

function initVozacStatus() {
    const heroName = document.getElementById('vozacHeroName');
    const heroId = document.getElementById('vozacHeroId');

    if (heroName) heroName.textContent = CONFIG.ENTITY_NAME || 'Vozač';
    if (heroId) heroId.textContent = CONFIG.ENTITY_ID ? 'ID: ' + CONFIG.ENTITY_ID : '';

    const saved = localStorage.getItem(VOZAC_STATUS_KEY) || 'slobodan';
    applyVozacStatusUI(saved);
}

function applyVozacStatusUI(status) {
    const seg = document.getElementById('vozacStatusSeg');
    if (!seg) return;
    seg.querySelectorAll('.vseg').forEach(btn => {
        btn.classList.toggle('is-active', btn.dataset.status === status);
    });
}

async function setVozacStatus(status) {
    localStorage.setItem(VOZAC_STATUS_KEY, status);
    applyVozacStatusUI(status);

    if (!navigator.onLine) return;

    try {
        await apiPost('updateKamionStatus', {
            vozacID: CONFIG.ENTITY_ID,
            status: status,
            ruta: ''
        });
    } catch (err) {
        console.error('setVozacStatus sync failed:', err);
    }
}

function loadVozacProfil() {
    const nameEl = document.getElementById('vozacProfilName');
    const idEl = document.getElementById('vozacProfilId');

    if (nameEl) nameEl.textContent = CONFIG.ENTITY_NAME || '—';
    if (idEl) idEl.textContent = CONFIG.ENTITY_ID || '—';

    if (typeof generateQRCode === 'function') {
        generateQRCode('vozacQrCanvas', JSON.stringify({
            type: 'VOZ',
            id: CONFIG.ENTITY_ID,
            name: CONFIG.ENTITY_NAME
        }));
    }
}

window.initVozacStatus = initVozacStatus;
window.applyVozacStatusUI = applyVozacStatusUI;
window.setVozacStatus = setVozacStatus;
window.loadVozacProfil = loadVozacProfil;
