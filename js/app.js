/* ==========================================================
   app.js: Alpine.js controller
   - Slide model (960x540 px space; export converts px → inches @96dpi)
   - Upload intercept: scan first, never open unscanned files
   - Export pipeline: PptxGenJS → arraybuffer → JSZip → .zip download
   ========================================================== */
(function () {
  'use strict';

  const PX_PER_IN = 96;
  const PT_TO_PX = 96 / 72; // canvas font px per typographic point

  let seq = 0;
  const uid = (p) => `${p || 'id'}_${Date.now().toString(36)}_${(seq++).toString(36)}`;
  const num = (v, fb) => (v !== '' && v !== null && Number.isFinite(+v) ? +v : fb);

  const makeText = (o) => Object.assign({
    id: uid('el'), type: 'text', text: 'New text',
    x: 80, y: 80, w: 400, h: 80,
    fontSize: 24, bold: false, color: '#222222', align: 'left',
  }, o || {});

  const makeSlide = (elements) => ({
    id: uid('sl'), background: '#FFFFFF', elements: elements || [],
  });

  document.addEventListener('alpine:init', () => {
    Alpine.data('pptApp', () => ({
      /* ---------- state ---------- */
      slides: [],
      current: 0,
      selectedId: null,

      showRepairModal: false,
      pendingFileBuffer: null,
      pendingFileName: '',
      repairReason: '',
      repairError: '',

      busy: false,
      toast: null,
      _toastTimer: null,
      _drag: null,

      /* ---------- lifecycle ---------- */
      init() {
        this.newDeck();
        this.$nextTick(() => {
          CanvasScaler.init({
            viewport: this.$refs.viewport,
            frame: this.$refs.frame,
            stage: this.$refs.stage,
            padding: 8,
          });
        });
      },

      /* ---------- derived ---------- */
      get currentSlide() { return this.slides[this.current] || null; },

      // Renderer block: while the repair modal is open, nothing is drawn.
      get renderableElements() {
        if (this.showRepairModal || !this.currentSlide) return [];
        return this.currentSlide.elements;
      },

      get selected() {
        if (!this.currentSlide) return null;
        return this.currentSlide.elements.find((e) => e.id === this.selectedId) || null;
      },

      slidePreview(s) {
        const t = s.elements.find((e) => e.text && e.text.trim());
        return t ? t.text : '(empty)';
      },

      elStyle(el) {
        return [
          `left:${num(el.x, 0)}px`,
          `top:${num(el.y, 0)}px`,
          `width:${num(el.w, 100)}px`,
          `height:${num(el.h, 40)}px`,
          `font-size:${num(el.fontSize, 18) * PT_TO_PX}px`,
          `font-weight:${el.bold ? 700 : 400}`,
          `color:${el.color}`,
          `text-align:${el.align}`,
        ].join(';');
      },

      /* ---------- deck editing ---------- */
      newDeck() {
        this.slides = [makeSlide([
          makeText({ text: 'Click to edit title', x: 60, y: 160, w: 840, h: 100, fontSize: 40, bold: true, align: 'center' }),
          makeText({ text: 'Subtitle goes here', x: 60, y: 280, w: 840, h: 60, fontSize: 22, align: 'center', color: '#555555' }),
        ])];
        this.current = 0;
        this.selectedId = null;
      },

      addSlide() {
        this.slides.push(makeSlide([
          makeText({ text: 'Slide title', x: 60, y: 40, w: 840, h: 70, fontSize: 32, bold: true }),
        ]));
        this.selectSlide(this.slides.length - 1);
      },

      deleteSlide() {
        if (this.slides.length < 2) return;
        this.slides.splice(this.current, 1);
        this.current = Math.min(this.current, this.slides.length - 1);
        this.selectedId = null;
      },

      selectSlide(i) {
        this.current = i;
        this.selectedId = null;
      },

      addText() {
        if (!this.currentSlide) return;
        const el = makeText();
        this.currentSlide.elements.push(el);
        this.selectedId = el.id;
      },

      deleteElement() {
        if (!this.currentSlide || !this.selected) return;
        const i = this.currentSlide.elements.findIndex((e) => e.id === this.selectedId);
        if (i > -1) this.currentSlide.elements.splice(i, 1);
        this.selectedId = null;
      },

      /* ---------- drag (scale-aware) ---------- */
      startDrag(ev, el) {
        this.selectedId = el.id;
        this._drag = {
          el,
          startX: ev.clientX, startY: ev.clientY,
          origX: num(el.x, 0), origY: num(el.y, 0),
        };
        const move = (e) => {
          const d = this._drag;
          if (!d) return;
          const s = CanvasScaler.getScale() || 1; // screen px → canvas px
          const maxX = CanvasScaler.BASE_W - 20;
          const maxY = CanvasScaler.BASE_H - 20;
          d.el.x = Math.round(Math.min(maxX, Math.max(-d.el.w + 20, d.origX + (e.clientX - d.startX) / s)));
          d.el.y = Math.round(Math.min(maxY, Math.max(-d.el.h + 20, d.origY + (e.clientY - d.startY) / s)));
        };
        const up = () => {
          this._drag = null;
          window.removeEventListener('pointermove', move);
          window.removeEventListener('pointerup', up);
          window.removeEventListener('pointercancel', up);
        };
        window.addEventListener('pointermove', move);
        window.addEventListener('pointerup', up);
        window.addEventListener('pointercancel', up);
      },

      /* ---------- upload: scan → (open | intercept) ---------- */
      async onFileChosen(ev) {
        const input = ev.target;
        const file = input.files && input.files[0];
        input.value = ''; // allow re-selecting the same file
        if (file) await this.ingestFile(file);
      },

      async ingestFile(file) {
        this.busy = true;
        try {
          const buffer = await file.arrayBuffer();
          try {
            const { slides } = await PptxRepair.scanBuffer(buffer);
            this.applySlides(slides);
            this.notify(`Opened ${file.name} (${slides.length} slides)`);
          } catch (err) {
            if (err instanceof PptxRepair.PptxScanError) {
              // HALT: stash raw bytes, lock workspace, show modal.
              this.pendingFileBuffer = buffer;
              this.pendingFileName = file.name;
              this.repairReason = err.message;
              this.repairError = '';
              this.showRepairModal = true;
              this.$nextTick(() => this.$refs.repairBtn && this.$refs.repairBtn.focus());
            } else {
              throw err;
            }
          }
        } catch (err) {
          console.error(err);
          this.notify('Could not read file: ' + (err.message || err), 'error');
        } finally {
          this.busy = false;
        }
      },

      applySlides(slides) {
        this.slides = slides.length ? slides : [makeSlide()];
        this.current = 0;
        this.selectedId = null;
        this.$nextTick(() => CanvasScaler.refresh());
      },

      /* ---------- repair modal actions ---------- */
      async forceRepair() {
        if (!this.pendingFileBuffer) return;
        this.busy = true;
        this.repairError = '';
        try {
          const { slides, stats } = await PptxRepair.forceRepair(this.pendingFileBuffer);
          if (!stats.slidesFound || !stats.textBlocks) {
            this.repairError = 'No recoverable slide text was found in this file.';
            return; // modal stays open so the user can cancel
          }
          this.applySlides(slides);
          this.showRepairModal = false;
          this.pendingFileBuffer = null;
          this.notify(`Recovered ${stats.textBlocks} text block(s) across ${stats.slidesFound} slide(s)`);
        } catch (err) {
          console.error(err);
          this.repairError = 'Repair failed: ' + (err.message || err);
        } finally {
          this.busy = false;
        }
      },

      cancelRepair() {
        if (this.busy) return;
        this.showRepairModal = false;
        this.pendingFileBuffer = null;
        this.pendingFileName = '';
        this.repairError = '';
        this.$nextTick(() => CanvasScaler.refresh());
      },

      /* ---------- export: PptxGenJS → arraybuffer → JSZip → .zip ---------- */
      async exportZip() {
        if (this.busy || !this.slides.length) return;
        if (typeof PptxGenJS === 'undefined' || typeof JSZip === 'undefined') {
          this.notify('Libraries failed to load from CDN. Check your connection.', 'error');
          return;
        }
        this.busy = true;
        try {
          const pres = new PptxGenJS();
          pres.layout = 'LAYOUT_16x9'; // 10in x 5.625in = 960x540 px @ 96dpi

          for (const s of this.slides) {
            const slide = pres.addSlide();
            slide.background = { color: (s.background || '#FFFFFF').replace('#', '') };
            for (const el of s.elements) {
              if (el.type !== 'text' || !String(el.text || '').trim()) continue;
              slide.addText(String(el.text), {
                x: num(el.x, 0) / PX_PER_IN,
                y: num(el.y, 0) / PX_PER_IN,
                w: Math.max(0.2, num(el.w, 100) / PX_PER_IN),
                h: Math.max(0.2, num(el.h, 40) / PX_PER_IN),
                fontFace: 'Calibri',
                fontSize: num(el.fontSize, 18),
                bold: !!el.bold,
                color: (el.color || '#222222').replace('#', ''),
                align: el.align || 'left',
                valign: 'top',
                margin: 0.04,
                fit: 'none',
              });
            }
          }

          // NOT writeFile(): get the compiled deck as an in-memory buffer.
          const pptxBuffer = await pres.write('arraybuffer');

          const zip = new JSZip();
          zip.file('presentation.pptx', pptxBuffer);
          zip.file('readme.txt', [
            'PPTX Studio: Export Manifest',
            '============================',
            `Exported:     ${new Date().toISOString()}`,
            `Total slides: ${this.slides.length}`,
            '',
            'Contents:',
            '  - presentation.pptx',
            '  - readme.txt',
            '',
          ].join('\n'));

          const blob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
          });

          this.download(blob, 'exported-package.zip');
          this.notify(`Exported ${this.slides.length} slide(s) → exported-package.zip`);
        } catch (err) {
          console.error(err);
          this.notify('Export failed: ' + (err.message || err), 'error');
        } finally {
          this.busy = false;
        }
      },

      download(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 10000);
      },

      /* ---------- toast ---------- */
      notify(msg, kind) {
        this.toast = { msg, kind: kind || 'ok' };
        clearTimeout(this._toastTimer);
        this._toastTimer = setTimeout(() => { this.toast = null; }, 4000);
      },
    }));
  });
})();
