window.safeAsync = async function (fn, onErrorToast) {
    try {
        return await fn();
    } catch (err) {
        console.error('[safeAsync]', err);

        if (onErrorToast) {
            showToast(onErrorToast, 'error');
        }

        return null;
    }
};
