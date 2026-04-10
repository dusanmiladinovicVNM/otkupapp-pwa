// ============================================================
// AUTH
// ============================================================
function showLoginScreen() {
    const header = document.querySelector('.header');
    const tabBar = document.querySelector('.tab-bar');

    if (header) header.style.display = 'none';
    if (tabBar) tabBar.style.display = 'none';

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
            <div class="login-bg-shape1"></div>
            <div class="login-bg-shape2"></div>

            <div class="login-top-area">
                <div class="login-app-icon">
                    <svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-label="OtkupApp logo">
                        <line x1="10" y1="50" x2="32" y2="10" stroke="#c8a84b" stroke-width="6.5" stroke-linecap="round"/>
                        <line x1="32" y1="10" x2="44" y2="50" stroke="#c8a84b" stroke-width="6.5" stroke-linecap="round"/>
                        <line x1="19" y1="34" x2="38" y2="34" stroke="#c8a84b" stroke-width="4.5" stroke-linecap="round"/>
                        <rect x="42" y="8" width="6" height="46" fill="#1e2d14"/>
                        <line x1="49" y1="10" x2="62" y2="50" stroke="#5ea135" stroke-width="6.5" stroke-linecap="round"/>
                        <line x1="62" y1="10" x2="49" y2="50" stroke="#5ea135" stroke-width="5" stroke-linecap="round"/>
                    </svg>
                </div>

                <div class="login-brand-lockup">
                    <div class="login-brand-name">
                        <span class="otkup">Otkup</span><span class="app">App</span>
                    </div>
                    <div class="login-brand-divider"></div>
                    <div class="login-brand-sub">
                        <span>by AgriX</span>
                    </div>
                </div>
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
                        maxlength="6"
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
                <div class="login-agrix-badge">AgriX ekosistem</div>
            </div>
        </div>
    </div>
`;

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
}

function applyRoleVisibility() {
    const role = CONFIG.USER_ROLE;
    document.querySelectorAll('.role-otkupac').forEach(el => el.style.display = (role === 'Otkupac') ? '' : 'none');
    document.querySelectorAll('.role-kooperant').forEach(el => el.style.display = (role === 'Kooperant') ? '' : 'none');
    document.querySelectorAll('.role-vozac').forEach(el => el.style.display = (role === 'Vozac') ? '' : 'none');
    document.querySelectorAll('.role-management').forEach(el => el.style.display = (role === 'Management') ? '' : 'none');
}

