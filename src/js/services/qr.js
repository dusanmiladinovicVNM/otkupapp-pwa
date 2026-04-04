function startQRScan() {
    const readerDiv = document.getElementById('qr-reader');
    readerDiv.style.display = 'block';
    if (qrScanner) qrScanner.clear();
    qrScanner = new Html5Qrcode('qr-reader');
    qrScanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => { onQRScanned(decodedText); qrScanner.stop().then(() => { readerDiv.style.display = 'none'; }); },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna: ' + err, 'error'); readerDiv.style.display = 'none'; });
}

function generateQRCode(canvasId, text) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    
    // Koristi QR generisanje iz eksternog CDN-a
    const img = new Image();
    img.onload = function() {
        canvas.width = 250;
        canvas.height = 250;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, 250, 250);
        ctx.drawImage(img, 0, 0, 250, 250);
    };
    img.onerror = function() {
        // Fallback: prikaži tekst ako API ne radi
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
