async function ensureHtml5QrcodeLoaded() {
    if (window.Html5Qrcode) return true;

    if (typeof window.lazyLoadScript !== 'function') {
        showToast('QR loader nije dostupan', 'error');
        return false;
    }

    try {
        await window.lazyLoadScript('/vendor/html5-qrcode.min.js');
        return !!window.Html5Qrcode;
    } catch (err) {
        console.error('html5-qrcode lazy load failed:', err);
        showToast('Ne mogu da učitam QR skener', 'error');
        return false;
    }
}

async function startQRScan() {
    const readerDiv = document.getElementById('qr-reader');
    if (!readerDiv) return;

    const ok = await ensureHtml5QrcodeLoaded();
    if (!ok) return;

    readerDiv.style.display = 'block';

    if (qrScanner) {
        try {
            await qrScanner.stop();
        } catch (e) {}

        try {
            qrScanner.clear();
        } catch (e) {}

        qrScanner = null;
    }

    qrScanner = new Html5Qrcode('qr-reader');

    qrScanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => {
            qrScanner.stop().then(() => {
                readerDiv.style.display = 'none';
                qrScanner = null;
            }).catch(() => {
                readerDiv.style.display = 'none';
                qrScanner = null;
            });

            onQRScanned(decodedText);
        },
        () => {}
    ).catch(err => {
        showToast('Kamera nije dostupna: ' + err, 'error');
        readerDiv.style.display = 'none';
        qrScanner = null;
    });
}

function generateQRCode(canvasId, text) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    const img = new Image();
    img.onload = function () {
        canvas.width = 250;
        canvas.height = 250;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, 250, 250);
        ctx.drawImage(img, 0, 0, 250, 250);
    };
    img.onerror = function () {
        canvas.width = 250;
        canvas.height = 250;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, 250, 250);
        ctx.fillStyle = '#1a5e2a';
        ctx.font = 'bold 16px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(CONFIG.ENTITY_ID, 125, 125);
    };
    img.src = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(text);
}

window.ensureHtml5QrcodeLoaded = ensureHtml5QrcodeLoaded;
window.startQRScan = startQRScan;
window.generateQRCode = generateQRCode;
