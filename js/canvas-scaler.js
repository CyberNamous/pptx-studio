/* ==========================================================
   canvas-scaler.js
   Locks the slide stage to a rigid 960x540 (16:9) base and
   scales it with a GPU-accelerated transform: scale() matrix
   to fit any viewport (ultra-wide desktop → narrow phone).

   Public API (global): window.CanvasScaler
     init({ viewport, frame, stage, padding?, maxScale? })
     refresh()
     getScale()
     destroy()
   Emits: window event "canvasscale" → { detail: { scale } }
   ========================================================== */
(function (global) {
  'use strict';

  const BASE_W = 960;
  const BASE_H = 540;
  const MIN_SCALE = 0.1;

  let viewportEl = null;
  let frameEl = null;
  let stageEl = null;
  let padding = 8;
  let maxScale = 4;
  let scale = 1;
  let resizeObserver = null;
  let rafId = 0;

  /** Pure fit computation: largest scale that keeps 16:9 inside the box. */
  function computeScale(availW, availH) {
    const w = Math.max(0, availW - padding * 2);
    const h = Math.max(0, availH - padding * 2);
    const s = Math.min(w / BASE_W, h / BASE_H);
    if (!Number.isFinite(s)) return 1;
    return Math.min(maxScale, Math.max(MIN_SCALE, s));
  }

  function apply() {
    rafId = 0;
    if (!viewportEl || !stageEl || !frameEl) return;

    const rect = viewportEl.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return; // hidden: keep last scale

    const next = Math.round(computeScale(rect.width, rect.height) * 10000) / 10000;

    // Frame reserves the *scaled* footprint so flex centering stays exact.
    frameEl.style.width = (BASE_W * next).toFixed(2) + 'px';
    frameEl.style.height = (BASE_H * next).toFixed(2) + 'px';

    // translate3d forces a compositor layer → hardware-accelerated.
    stageEl.style.transform = 'translate3d(0,0,0) scale(' + next + ')';

    if (next !== scale) {
      scale = next;
      global.dispatchEvent(new CustomEvent('canvasscale', { detail: { scale } }));
    }
  }

  /** Coalesce bursts of resize events into one frame. */
  function schedule() {
    if (rafId) return;
    rafId = global.requestAnimationFrame(apply);
  }

  function init(opts) {
    destroy();
    viewportEl = opts.viewport;
    frameEl = opts.frame;
    stageEl = opts.stage;
    if (typeof opts.padding === 'number') padding = opts.padding;
    if (typeof opts.maxScale === 'number') maxScale = opts.maxScale;
    if (!viewportEl || !frameEl || !stageEl) {
      throw new Error('CanvasScaler.init requires viewport, frame and stage elements.');
    }

    if ('ResizeObserver' in global) {
      resizeObserver = new ResizeObserver(schedule);
      resizeObserver.observe(viewportEl);
    }
    // Fallbacks / mobile browser chrome (URL bar show/hide, rotation).
    global.addEventListener('resize', schedule);
    global.addEventListener('orientationchange', schedule);
    if (global.visualViewport) global.visualViewport.addEventListener('resize', schedule);

    apply();
  }

  function destroy() {
    if (resizeObserver) resizeObserver.disconnect();
    resizeObserver = null;
    global.removeEventListener('resize', schedule);
    global.removeEventListener('orientationchange', schedule);
    if (global.visualViewport) global.visualViewport.removeEventListener('resize', schedule);
    if (rafId) global.cancelAnimationFrame(rafId);
    rafId = 0;
    viewportEl = frameEl = stageEl = null;
  }

  global.CanvasScaler = {
    BASE_W,
    BASE_H,
    init,
    destroy,
    refresh: apply,
    getScale: () => scale,
    computeScale,
  };
})(window);
