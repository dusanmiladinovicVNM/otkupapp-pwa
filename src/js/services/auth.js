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
        <div class="login-container">
            <div class="login-logo">
                <svg
                  width="110"
                  height="110"
                  viewBox="0 0 512 512"
                  fill="none"
                  xmlns="http://www.w3.org/2000/svg"
                  aria-label="OtkupApp logo"
                >
                  <rect width="512" height="512" rx="120" fill="#1a5e2a"/>
                  <g transform="translate(106,106)">
                    <circle cx="150" cy="150" r="130" stroke="white" stroke-width="16" fill="white" fill-opacity="0.05"/>
                    <line x1="150" y1="40" x2="150" y2="260" stroke="white" stroke-width="8" stroke-linecap="round" opacity="0.2"/>
                    <line x1="40" y1="150" x2="260" y2="150" stroke="white" stroke-width="8" stroke-linecap="round" opacity="0.2"/>

                    <g transform="translate(82,82)">
                      <rect x="0" y="30" width="12" height="25" rx="2" fill="#e8a838"/>
                      <rect x="20" y="15" width="12" height="40" rx="2" fill="#e8a838"/>
                      <rect x="40" y="0" width="12" height="55" rx="2" fill="#e8a838"/>
                      <circle cx="46" cy="10" r="4" fill="white" opacity="0.5"/>
                    </g>

                    <g transform="translate(182,76)">
                      <path d="M22 62C22 62 44 49 44 27C44 5 22 0 22 0C22 0 0 5 0 27C0 49 22 62 22 62Z" fill="#2d8a42"/>
                      <path d="M22 58V6" stroke="white" stroke-width="2.5" stroke-linecap="round" opacity="0.55"/>
                      <path d="M22 44L34 32M22 32L30 24M22 20L27 15" stroke="white" stroke-width="1.5" stroke-linecap="round" opacity="0.3"/>
                    </g>

                    <g transform="translate(70,182)">
                      <path d="M42 10H60V32H42V10Z" fill="white"/>
                      <path d="M60 14L70 22V32H60V14Z" fill="white"/>
                      <rect x="0" y="8" width="38" height="24" rx="2" fill="white"/>
                      <circle cx="10" cy="38" r="5.5" fill="white"/>
                      <circle cx="28" cy="38" r="5.5" fill="white"/>
                      <circle cx="55" cy="38" r="5.5" fill="white"/>
                    </g>

                    <g transform="translate(164,166)">
                      <rect x="0" y="26" width="66" height="40" rx="2" fill="#e8a838"/>
                      <rect x="0" y="30" width="66" height="5" fill="#1a5e2a" fill-opacity="0.1"/>
                      <rect x="0" y="40" width="66" height="5" fill="#1a5e2a" fill-opacity="0.1"/>
                      <rect x="0" y="50" width="66" height="5" fill="#1a5e2a" fill-opacity="0.1"/>
                      <rect x="0" y="26" width="7" height="40" fill="#1a5e2a" fill-opacity="0.15"/>
                      <rect x="59" y="26" width="7" height="40" fill="#1a5e2a" fill-opacity="0.15"/>

                      <g transform="translate(8,-4)">
                        <circle cx="20" cy="20" r="17" fill="#e8a838" stroke="#1a5e2a" stroke-width="2"/>
                        <path d="M20 5C20 5 34 5 34 14" stroke="#2d8a42" stroke-width="4" stroke-linecap="round"/>
                      </g>
                    </g>
                  </g>
                </svg>
            </div>

            <h2 class="login-title">OtkupApp</h2>

            <div class="login-form">
                <div class="form-group">
                    <label>Korisničko ime</label>
                    <input
                        type="text"
                        id="loginUsername"
                        autocapitalize="none"
                        autocorrect="off"
                        placeholder="korisničko ime"
                    >
                </div>

                <div class="form-group">
                    <label>PIN</label>
                    <input
                        type="password"
                        id="loginPin"
                        inputmode="numeric"
                        maxlength="6"
                        placeholder="● ● ● ●"
                    >
                </div>

                <button onclick="doLogin()" id="btnLogin" class="btn-primary">
                    Prijavi se
                </button>

                <p id="loginError" class="login-error" style="display:none;"></p>
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

