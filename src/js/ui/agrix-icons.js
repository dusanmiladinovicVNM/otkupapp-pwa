// ============================================================
// AgriX SVG icons
// Plain JS global helper for otkupapp-pwa.
// SVG-only, currentColor, no emoji.
// ============================================================

window.AgriXIcons = Object.freeze({
    sprout: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20v-8"/><path d="M12 12c0-4 3-6 7-6-1 5-3 7-7 6Z"/><path d="M12 14c0-3-2-5-6-5 1 4 3 6 6 5Z"/></svg>`,
    qr: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><path d="M14 14h3v3h-3zM20 14v7M14 20h3"/></svg>`,
    camera: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 8h4l2-3h6l2 3h4v12H3z"/><circle cx="12" cy="13" r="4"/></svg>`,
    home: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 11 12 4l9 7v9a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1v-9Z"/></svg>`,
    map: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="m3 6 6-2 6 2 6-2v14l-6 2-6-2-6 2V6Z"/><path d="M9 4v14M15 6v14"/></svg>`,
    tractor: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 16h3l1-6h5l3 4h3v2"/><circle cx="7" cy="18" r="2"/><circle cx="17" cy="17" r="3"/></svg>`,
    clipboard: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="5" width="12" height="16" rx="2"/><path d="M9 5a3 3 0 0 1 6 0M9 11h6M9 15h6"/></svg>`,
    doc: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h9l3 3v15H6V3Z"/><path d="M15 3v3h3M9 12h6M9 16h6M9 8h3"/></svg>`,
    info: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 7.5v.5"/></svg>`,
    truck: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h10v10H3zM13 10h5l3 3v3h-8"/><circle cx="7" cy="17" r="2"/><circle cx="17" cy="17" r="2"/></svg>`,
    package: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="m12 3 8 4v10l-8 4-8-4V7l8-4Z"/><path d="m4 7 8 4 8-4M12 11v10"/></svg>`,
    factory: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21V11l6 4V11l6 4V7l6 4v10Z"/><path d="M8 17h2M14 17h2"/></svg>`,
    check: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m5 12 5 5 9-11"/></svg>`,
    x: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6 18 18M18 6 6 18"/></svg>`,
    plus: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.25" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>`,
    "arrow-right": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>`,
    "arrow-left": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5M11 6l-6 6 6 6"/></svg>`,
    chart: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18M6 17v-6M11 17V7M16 17v-4M20 17v-9"/></svg>`,
    coin: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M14.5 9.5c-.8-1-4.5-1.5-4.5 1s5 1 5 3.5c0 2-3.5 2.5-5 1"/><path d="M12 6v2M12 16v2"/></svg>`,
    user: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 4-7 8-7s8 3 8 7"/></svg>`,
    cart: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 4h2l2.5 11.5a2 2 0 0 0 2 1.5h7a2 2 0 0 0 2-1.5L21 8H6"/><circle cx="10" cy="20" r="1.25"/><circle cx="17" cy="20" r="1.25"/></svg>`,
    print: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V3h12v6"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M6 14h12v7H6z"/></svg>`
});

function agIcon(name, size, label, extraClass) {
    var svg = window.AgriXIcons && window.AgriXIcons[name];

    if (!svg) {
        return '';
    }

    var cls = 'i' + (extraClass ? ' ' + String(extraClass) : '');
    var style = size ? ' style="--is:' + escapeHtml(String(size)) + ';"' : '';
    var aria = label
        ? ' role="img" aria-label="' + escapeHtml(label) + '"'
        : ' aria-hidden="true"';

    return '<span class="' + cls + '"' + style + aria + '>' + svg + '</span>';
}

function mountAgriXIcons(root) {
    var scope = root || document;
    var nodes = scope.querySelectorAll('[data-ag-icon]');

    nodes.forEach(function (node) {
        var name = node.getAttribute('data-ag-icon');
        var size = node.getAttribute('data-ag-icon-size') || '';
        var label = node.getAttribute('data-ag-icon-label') || '';
        var svg = window.AgriXIcons && window.AgriXIcons[name];

        if (!svg) return;

        node.classList.add('i');

        if (size) {
            node.style.setProperty('--is', size);
        }

        if (label) {
            node.setAttribute('role', 'img');
            node.setAttribute('aria-label', label);
        } else {
            node.setAttribute('aria-hidden', 'true');
        }

        node.innerHTML = svg;
    });
}

window.agIcon = agIcon;
window.mountAgriXIcons = mountAgriXIcons;

document.addEventListener('DOMContentLoaded', function () {
    mountAgriXIcons(document);
});
