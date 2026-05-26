// utils/lazy.js
window.lazyLoadScript = function (src) {
    if (window.__lazyLoaded && window.__lazyLoaded[src]) return Promise.resolve();
    return new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = () => {
            window.__lazyLoaded = window.__lazyLoaded || {};
            window.__lazyLoaded[src] = true;
            resolve();
        };
        s.onerror = reject;
        document.head.appendChild(s);
    });
};
