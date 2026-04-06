function startQRScan() {
    const readerDiv = document.getElementById('qr-reader');
    if (!readerDiv) return;

    readerDiv.style.display = 'block';

    // Čisti prethodni scanner ako postoji
    if (qrScanner) {
        try {
            qrScanner.stop().catch(() => {});
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
                qrScanner = null; // Oslobodi referencu
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
