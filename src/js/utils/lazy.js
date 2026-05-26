// utils/lazy.js
window.lazyLoadScript = function (src) {
    window.__lazyLoaded = window.__lazyLoaded || {};

    if (window.__lazyLoaded[src]) return Promise.resolve();

    const existing = document.querySelector('script[src="' + src + '"]');
    if (existing) {
        window.__lazyLoaded[src] = true;
        return Promise.resolve();
    }

    return new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = () => {
            window.__lazyLoaded[src] = true;
            resolve();
        };
        s.onerror = reject;
        document.head.appendChild(s);
    });
};
