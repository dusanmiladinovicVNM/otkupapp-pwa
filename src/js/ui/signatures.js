function initSignaturePad(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    ctx.scale(scaleX, scaleY);
    ctx.strokeStyle = '#1a1a1a';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    let drawing = false, lastX = 0, lastY = 0;
    function getPos(e) { const r = canvas.getBoundingClientRect(); const touch = e.touches ? e.touches[0] : e; return { x: touch.clientX - r.left, y: touch.clientY - r.top }; }
    function startDraw(e) { e.preventDefault(); drawing = true; const p = getPos(e); lastX = p.x; lastY = p.y; }
    function draw(e) { if (!drawing) return; e.preventDefault(); const p = getPos(e); ctx.beginPath(); ctx.moveTo(lastX, lastY); ctx.lineTo(p.x, p.y); ctx.stroke(); lastX = p.x; lastY = p.y; }
    function stopDraw(e) { if (e) e.preventDefault(); drawing = false; }
    canvas.addEventListener('mousedown', startDraw); canvas.addEventListener('mousemove', draw);
    canvas.addEventListener('mouseup', stopDraw); canvas.addEventListener('mouseleave', stopDraw);
    canvas.addEventListener('touchstart', startDraw, { passive: false });
    canvas.addEventListener('touchmove', draw, { passive: false });
    canvas.addEventListener('touchend', stopDraw, { passive: false });
}

function clearSignature(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return;
    canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height);
}

function getSignatureData(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return '';
    const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    if (!data.some((val, i) => i % 4 === 3 && val > 0)) return '';
    return canvas.toDataURL('image/png');
}
