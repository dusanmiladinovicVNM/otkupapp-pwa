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
                            src="img/AgriX-Logo-Final_Novi.png"
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

                    <button id="btnLogin" class="btn-login">
                        Prijavi se
                    </button>
                </div>

                <div class="login-bottom-area">
                    <a class="login-footer-logo-link" href="index.html" aria-label="AgriX">
                        <img
                            src="img/AgriX-Logo-Final_Novi.png"
                            alt="AgriX"
                            class="login-footer-logo"
                        >
                    </a>
                </div>
            </div>
        </div>
    `;
    
    const loginBtn = document.getElementById('btnLogin');
    if (loginBtn && !loginBtn.dataset.bound) {
        loginBtn.addEventListener('click', doLogin);
        loginBtn.dataset.bound = '1';
    }
    
    document.getElementById('loginPin').addEventListener('keyup', e => {
        if (e.key === 'Enter') doLogin();
    });
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
        const entityID = json.entityID || json.otkupacID || '';

        setLs('authToken', json.token);
        setLs('userRole', json.role);
        setLs('entityID', entityID);
        setLs('entityName', json.displayName);
        setLs('username', username);

        // legacy alias ostaje samo za Otkupac flow
        if (json.role === 'Otkupac') {
            setLs('otkupacID', entityID);
        } else {
            removeLs('otkupacID');
        }

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

    ['koopBottomNav', 'otkupBottomNav', 'mgmtBottomNav'].forEach(id => {
        const nav = document.getElementById(id);
        if (nav) nav.classList.remove('visible');
    });

    document.body.classList.remove('has-koop-bottom-nav', 'has-otkup-bottom-nav', 'has-mgmt-bottom-nav');

    location.reload();
}

function applyRoleVisibility() {
    const role = String((CONFIG && CONFIG.USER_ROLE) || '').trim().toLowerCase();

    document.querySelectorAll('.role-otkupac').forEach(el => {
        el.style.display = (role === 'otkupac') ? '' : 'none';
    });

    document.querySelectorAll('.role-kooperant').forEach(el => {
        el.style.display = (role === 'kooperant') ? '' : 'none';
    });

    document.querySelectorAll('.role-vozac').forEach(el => {
        el.style.display = (role === 'vozac') ? '' : 'none';
    });

    document.querySelectorAll('.role-management').forEach(el => {
        el.style.display = (role === 'management') ? '' : 'none';
    });

    applyHeaderBranding();
}

function applyHeaderBranding() {
    const logoEl = document.getElementById('headerBrandLogo');
    if (!logoEl) return;

    const role = String((CONFIG && CONFIG.USER_ROLE) || '').toLowerCase();

    const isKooperant = role === 'kooperant';

    if (isKooperant) {
        logoEl.src = 'img/AgriX-Gazdinstvo-Logo-Final.png';
        logoEl.alt = 'AgriX Gazdinstvo';
    } else {
        logoEl.src = 'img/AgriX-Otkup-Logo-Final.png';
        logoEl.alt = 'AgriX Otkup';
    }
}

window.addEventListener('resize', () => {
    applyRoleVisibility();

    if (typeof updateRoleNavVisibility === 'function') updateRoleNavVisibility();
    if (typeof updateRoleNavActive === 'function') updateRoleNavActive();
});
