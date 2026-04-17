// ============================================================
// AUTH
// ============================================================
function showLoginScreen() {
    const header = document.querySelector('.header');
    const tabBar = document.querySelector('.tab-bar');

    if (header) header.style.display = 'none';
    if (tabBar) tabBar.style.display = 'none';

    ['koopBottomNav', 'otkupBottomNav'].forEach(id => {
        const nav = document.getElementById(id);
        if (nav) nav.classList.remove('visible');
    });
    document.body.classList.remove('has-koop-bottom-nav', 'has-otkup-bottom-nav');


    document.querySelectorAll('.tab-content').forEach(t => t.style.display = 'none');
    document.querySelectorAll('.sub-tab-bar').forEach(t => t.style.display = 'none');

    let container = document.getElementById('loginContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'loginContainer';
        document.body.appendChild(container);
    }

    container.innerHTML = `
        <div class="login-screen">
            <div class="login-shell">

                <div class="login-top-area">
                    <a class="login-top-logo-link" href="index.html" aria-label="AgriX Gazdinstvo">
                        <img
                            src="site/img/AgriX-Logo-Final_Novi.png"
                            alt="AgriX Gazdinstvo"
                            class="login-top-logo"
                        >
                    </a>
                </div>

                <div class="login-card">
                    <div class="login-card-title">Prijava</div>

                    <div class="login-field">
                        <label class="login-field-label">Korisničko ime</label>
                        <input
                            class="login-field-input"
                            type="text"
                            id="loginUsername"
                            autocapitalize="none"
                            autocorrect="off"
                            placeholder="unesite korisničko ime"
                            autocomplete="username"
                        >
                    </div>

                    <div class="login-field">
                        <label class="login-field-label">PIN</label>
                        <input
                            class="login-field-input login-pin-input"
                            type="password"
                            id="loginPin"
                            inputmode="numeric"
                            maxlength="4"
                            placeholder="• • • •"
                            autocomplete="current-password"
                        >
                    </div>

                    <p id="loginError" class="login-error" style="display:none;"></p>

                    <button onclick="doLogin()" id="btnLogin" class="btn-login">
                        Prijavi se
                    </button>
                </div>

                <div class="login-bottom-area">
                    <a class="login-footer-logo-link" href="index.html" aria-label="AgriX">
                        <img
                            src="site/img/AgriX-Logo-Final_Novi.png"
                            alt="AgriX"
                            class="login-footer-logo"
                        >
                    </a>
                </div>
            </div>
        </div>
    `;

    document.getElementById('loginPin').addEventListener('keyup', e => {
        if (e.key === 'Enter') doLogin();
    });
    updateKoopBottomNavVisibility();
    document.getElementById('loginUsername').focus();
}

async function doLogin() {
    const username = (byId('loginUsername').value || '').trim();
    const pin = byId('loginPin').value || '';
    const errorEl = byId('loginError');
    const btnEl = byId('btnLogin');

    if (!username) {
        setText(errorEl, 'Unesite korisničko ime');
        showEl(errorEl, 'block');
        return;
    }

    if (!pin) {
        setText(errorEl, 'Unesite PIN');
        showEl(errorEl, 'block');
        return;
    }

    hideEl(errorEl);
    setText(btnEl, 'Prijavljivanje...');
    btnEl.disabled = true;

    const json = await apiPost('login', { username, pin });

    if (json && json.success) {
        setLs('authToken', json.token);
        setLs('userRole', json.role);
        setLs('otkupacID', json.entityID);
        setLs('entityID', json.entityID);
        setLs('entityName', json.displayName);
        setLs('username', username);
        location.reload();
        return;
    }

    if (json && !json.success) {
        setText(errorEl, json.error || 'Prijava neuspešna');
        showEl(errorEl, 'block');
    } else {
        setText(errorEl, 'Nema internet konekcije');
        showEl(errorEl, 'block');
    }

    setText(btnEl, 'Prijavi se');
    btnEl.disabled = false;
}

function doLogout() {
    removeLs(['userRole', 'otkupacID', 'entityID', 'entityName', 'authToken', 'username']);
    location.reload();
    if (typeof updateBottomNavVisibility === 'function') updateBottomNavVisibility();
}

function applyRoleVisibility() {
    // DODAJ — resetuj header koji je možda sakriven od showLoginScreen()
    const header = document.querySelector('.header');
    if (header) header.style.display = '';
    const role = String(CONFIG.USER_ROLE || '').trim();
    const isMobile = window.matchMedia('(max-width: 900px)').matches;
    const isKooperant = role === 'Kooperant';
    const isOtkupac = role === 'Otkupac';
    const isBottomNavRole = isKooperant || isOtkupac;

    // 1) prvo sakrij sve role-specifične elemente
    document.querySelectorAll('.role-otkupac, .role-kooperant, .role-vozac, .role-management')
        .forEach(el => {
            el.style.display = 'none';
        });

    // 2) helper za aktivnu rolu
    const activeSelector =
        isOtkupac ? '.role-otkupac' :
        isKooperant ? '.role-kooperant' :
        role === 'Vozac' ? '.role-vozac' :
        role === 'Management' ? '.role-management' :
        '';

    if (activeSelector) {
        document.querySelectorAll(activeSelector).forEach(el => {
            // Bottom nav: samo mobile za Kooperant/Otkupac
            if (el.classList.contains('bottom-nav')) {
                el.style.display = (isBottomNavRole && isMobile) ? '' : 'none';
                return;
            }

            // Tab buttons: desktop za Kooperant/Otkupac, uvek za ostale role
            if (el.classList.contains('tab-btn')) {
                if (isBottomNavRole) {
                    el.style.display = isMobile ? 'none' : '';
                } else {
                    el.style.display = '';
                }
                return;
            }

            // Svi ostali role elementi
            el.style.display = '';
        });
    }

    // 3) tab bar kontejner:
    // - mobile: sakriven za Kooperant/Otkupac
    // - desktop: prikazan
    const tabBar = document.querySelector('.tab-bar');
    if (tabBar) {
        if (isBottomNavRole && isMobile) {
            tabBar.style.display = 'none';
        } else {
            tabBar.style.display = 'flex';
        }
    }

    applyHeaderBranding();
}

function applyHeaderBranding() {
    const logoEl = document.getElementById('headerBrandLogo');
    if (!logoEl) return;

    const role = String((CONFIG && CONFIG.USER_ROLE) || '').toLowerCase();

    const isKooperant = role === 'kooperant';

    if (isKooperant) {
        logoEl.src = 'site/img/AgriX-Gazdinstvo-Logo-Final.png';
        logoEl.alt = 'AgriX Gazdinstvo';
    } else {
        logoEl.src = 'site/img/AgriX-Otkup-Logo-Final.png';
        logoEl.alt = 'AgriX Otkup';
    }
}

window.addEventListener('resize', () => {
    applyRoleVisibility();
    if (typeof updateBottomNavVisibility === 'function') updateBottomNavVisibility();
    if (typeof updateBottomNavActive === 'function') updateBottomNavActive();
});
