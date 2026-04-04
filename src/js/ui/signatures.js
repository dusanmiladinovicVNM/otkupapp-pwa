(function () {
    const signaturePads = new Map();

    function getCanvasSize(canvas) {
        const rect = canvas.getBoundingClientRect();
        return {
            cssWidth: Math.max(1, Math.round(rect.width || canvas.clientWidth || 300)),
            cssHeight: Math.max(1, Math.round(rect.height || canvas.clientHeight || 150))
        };
    }

    function setupCanvas(canvas, padState) {
        const dpr = Math.max(window.devicePixelRatio || 1, 1);
        const { cssWidth, cssHeight } = getCanvasSize(canvas);

        canvas.width = Math.round(cssWidth * dpr);
        canvas.height = Math.round(cssHeight * dpr);

        const ctx = canvas.getContext('2d');

        // reset transform before applying DPR scale
        if (typeof ctx.resetTransform === 'function') {
            ctx.resetTransform();
        } else {
            ctx.setTransform(1, 0, 0, 1, 0, 0);
        }

        ctx.scale(dpr, dpr);
        ctx.strokeStyle = '#1a1a1a';
        ctx.lineWidth = 2;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';

        padState.ctx = ctx;
        padState.dpr = dpr;
        padState.cssWidth = cssWidth;
        padState.cssHeight = cssHeight;
    }

    function getPos(canvas, e) {
        const rect = canvas.getBoundingClientRect();
        const point = e.touches ? e.touches[0] : e;

        return {
            x: point.clientX - rect.left,
            y: point.clientY - rect.top
        };
    }

    function bindPad(canvasId) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return null;

        if (signaturePads.has(canvasId)) {
            return signaturePads.get(canvasId);
        }

        const padState = {
            canvas,
            ctx: null,
            drawing: false,
            lastX: 0,
            lastY: 0,
            hasInk: false
        };

        setupCanvas(canvas, padState);

        function startDraw(e) {
            e.preventDefault();
            padState.drawing = true;

            const p = getPos(canvas, e);
            padState.lastX = p.x;
            padState.lastY = p.y;
        }

        function draw(e) {
            if (!padState.drawing) return;
            e.preventDefault();

            const p = getPos(canvas, e);

            padState.ctx.beginPath();
            padState.ctx.moveTo(padState.lastX, padState.lastY);
            padState.ctx.lineTo(p.x, p.y);
            padState.ctx.stroke();

            padState.lastX = p.x;
            padState.lastY = p.y;
            padState.hasInk = true;
        }

        function stopDraw(e) {
            if (e) e.preventDefault();
            padState.drawing = false;
        }

        canvas.addEventListener('mousedown', startDraw);
        canvas.addEventListener('mousemove', draw);
        canvas.addEventListener('mouseup', stopDraw);
        canvas.addEventListener('mouseleave', stopDraw);

        canvas.addEventListener('touchstart', startDraw, { passive: false });
        canvas.addEventListener('touchmove', draw, { passive: false });
        canvas.addEventListener('touchend', stopDraw, { passive: false });
        canvas.addEventListener('touchcancel', stopDraw, { passive: false });

        signaturePads.set(canvasId, padState);
        return padState;
    }

    window.initSignaturePad = function (canvasId) {
        const pad = bindPad(canvasId);
        if (!pad) return;

        // ako se layout promenio od poslednjeg otvaranja modala, re-setup canvas
        const { cssWidth, cssHeight } = getCanvasSize(pad.canvas);
        if (cssWidth !== pad.cssWidth || cssHeight !== pad.cssHeight) {
            const hadInk = pad.hasInk;
            setupCanvas(pad.canvas, pad);
            if (hadInk) {
                // ovde ne pokušavamo da restauriramo postojeći potpis
                // da ne komplikujemo stabilization fazu
                pad.hasInk = false;
            }
        }
    };

    window.clearSignature = function (canvasId) {
        const pad = bindPad(canvasId);
        if (!pad) return;

        if (typeof pad.ctx.resetTransform === 'function') {
            pad.ctx.resetTransform();
        } else {
            pad.ctx.setTransform(1, 0, 0, 1, 0, 0);
        }

        pad.ctx.clearRect(0, 0, pad.canvas.width, pad.canvas.height);

        pad.ctx.scale(pad.dpr, pad.dpr);
        pad.ctx.strokeStyle = '#1a1a1a';
        pad.ctx.lineWidth = 2;
        pad.ctx.lineCap = 'round';
        pad.ctx.lineJoin = 'round';

        pad.hasInk = false;
        pad.drawing = false;
    };

    window.getSignatureData = function (canvasId) {
        const pad = bindPad(canvasId);
        if (!pad) return '';
        if (!pad.hasInk) return '';

        return pad.canvas.toDataURL('image/png');
    };
})();
