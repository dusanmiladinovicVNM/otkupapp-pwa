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
