window.qs = function (selector, root) {
    return (root || document).querySelector(selector);
};

window.qsa = function (selector, root) {
    return Array.from((root || document).querySelectorAll(selector));
};

window.byId = function (id) {
    return document.getElementById(id);
};

window.showEl = function (el, displayValue = '') {
    if (!el) return;
    el.classList.remove('is-hidden');
    if (displayValue) el.style.display = displayValue;
    else el.style.display = '';   // reset to CSS default
};

window.hideEl = function (el) {
    if (!el) return;
    el.classList.add('is-hidden');
    el.style.display = '';        // clear any inline override
};

window.setText = function (el, text) {
    if (!el) return;
    el.textContent = text;
};

window.setHtml = function (el, html) {
    if (!el) return;
    el.innerHTML = html;
};

window.addClass = function (el, className) {
    if (!el) return;
    el.classList.add(className);
};

window.removeClass = function (el, className) {
    if (!el) return;
    el.classList.remove(className);
};

window.toggleClass = function (el, className, force) {
    if (!el) return;
    if (typeof force === 'boolean') {
        el.classList.toggle(className, force);
    } else {
        el.classList.toggle(className);
    }
};
