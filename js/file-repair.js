/* ==========================================================
   file-repair.js
   1) scanBuffer(): strict integrity scan using JSZip (CRC32 on).
      Any ZIP / checksum / XML-syntax failure → PptxScanError.
   2) forceRepair(): fallback recovery.
        Tier A: JSZip with CRC checks disabled, per-entry try/catch.
        Tier B: manual walk of local ZIP headers (PK\x03\x04) with
                native DecompressionStream('deflate-raw'), tolerant
                of truncated streams. Zero dependencies.
        Then regex-scrape <a:t> text from raw slide XML.

   Global: window.PptxRepair
   ========================================================== */
(function (global) {
  'use strict';

  const SLIDE_RE = /^ppt\/slides\/slide(\d+)\.xml$/i;
  const WANTED_RE = /^ppt\/(?:slides\/slide\d+|presentation)\.xml$/i;
  const DEFAULT_SLD_W = 12192000; // EMU (13.333in widescreen)
  const DEFAULT_SLD_H = 6858000;
  const CANVAS_W = 960;
  const CANVAS_H = 540;

  class PptxScanError extends Error {
    constructor(message, code) {
      super(message);
      this.name = 'PptxScanError';
      this.code = code || 'SCAN_FAILED';
    }
  }

  const rid = () => 'e' + Math.random().toString(36).slice(2, 10);
  const utf8 = new TextDecoder('utf-8', { fatal: false });

  /* ---------------- Precompiled regular expressions ---------------- */
  // Match <a:t> / <a:t xml:space="preserve"> but NOT <a:tab>, <a:tbl>, <a:tc> …
  // Alternation also catches <a:br/> so line breaks survive.
  const RE_TOKEN = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>|<a:br\b[^>]*\/?>/g;
  // <a:p> / <a:p ...> but NOT <a:pPr>, <a:pPr/>
  const RE_PARA = /<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g;
  // Split on shape *openings* so truncated files still yield their last shape.
  const RE_SHAPE_OPEN = /<p:sp(?:\s[^>]*)?>/g;
  const RE_OFF = /<a:off\b[^>]*>/;
  const RE_EXT = /<a:ext\b[^>]*\bcx="\d+"[^>]*>/;
  const RE_SZ = /<a:rPr\b[^>]*\bsz="(\d+)"/;
  const RE_BOLD = /<a:rPr\b[^>]*\bb="1"/;
  const RE_SLDSZ = /<p:sldSz\b[^>]*>/;
  const RE_ENTITY = /&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi;

  const attr = (tag, name) => {
    if (!tag) return null;
    const m = new RegExp('\\b' + name + '="(-?\\d+)"').exec(tag);
    return m ? parseInt(m[1], 10) : null;
  };

  function decodeEntities(s) {
    return s.replace(RE_ENTITY, (_, e) => {
      const k = e.toLowerCase();
      if (k === 'lt') return '<';
      if (k === 'gt') return '>';
      if (k === 'amp') return '&';
      if (k === 'quot') return '"';
      if (k === 'apos') return "'";
      const code = k[1] === 'x' ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10);
      try { return String.fromCodePoint(code); } catch (_) { return ''; }
    });
  }

  /** Extract plain text (paragraphs joined by \n) from a chunk of raw XML. */
  function textOfChunk(chunk) {
    const paras = chunk.match(RE_PARA) || [chunk]; // truncated tail → scan raw
    const lines = [];
    for (const p of paras) {
      let line = '';
      RE_TOKEN.lastIndex = 0;
      let m;
      while ((m = RE_TOKEN.exec(p)) !== null) {
        line += m[1] === undefined ? '\n' : decodeEntities(m[1]);
      }
      if (line.trim()) lines.push(line);
    }
    return lines.join('\n').trim();
  }

  function slideDimsFromXml(presXml) {
    const tag = presXml ? RE_SLDSZ.exec(presXml) : null;
    return {
      cx: (tag && attr(tag[0], 'cx')) || DEFAULT_SLD_W,
      cy: (tag && attr(tag[0], 'cy')) || DEFAULT_SLD_H,
    };
  }

  /**
   * Regex-scrape one slide's raw XML into an editable slide object.
   * Pure string work: safe on malformed / truncated XML.
   */
  function scrapeSlideXml(xml, dims) {
    const d = dims || { cx: DEFAULT_SLD_W, cy: DEFAULT_SLD_H };
    const sx = CANVAS_W / d.cx;
    const sy = CANVAS_H / d.cy;
    const elements = [];

    const chunks = xml.split(RE_SHAPE_OPEN).slice(1);
    chunks.forEach((chunk) => {
      const text = textOfChunk(chunk);
      if (!text) return;

      const off = RE_OFF.exec(chunk);
      const ext = RE_EXT.exec(chunk);
      const ox = off ? attr(off[0], 'x') : null;
      const oy = off ? attr(off[0], 'y') : null;
      const cx = ext ? attr(ext[0], 'cx') : null;
      const cy = ext ? attr(ext[0], 'cy') : null;
      const idx = elements.length;

      const szM = RE_SZ.exec(chunk);
      const hasGeom = ox !== null && oy !== null && cx && cy;

      elements.push({
        id: rid(),
        type: 'text',
        text,
        x: hasGeom ? Math.round(ox * sx) : 60,
        y: hasGeom ? Math.round(oy * sy) : 40 + idx * 120,
        w: hasGeom ? Math.max(40, Math.round(cx * sx)) : 840,
        h: hasGeom ? Math.max(30, Math.round(cy * sy)) : 100,
        fontSize: szM ? Math.max(6, Math.round(parseInt(szM[1], 10) / 100)) : (idx === 0 ? 32 : 20),
        bold: RE_BOLD.test(chunk) || (!szM && idx === 0),
        color: '#222222',
        align: 'left',
      });
    });

    // Fallback: no <p:sp> text (tables, damaged tree): scan the whole document.
    if (!elements.length) {
      const text = textOfChunk(xml);
      if (text) {
        elements.push({
          id: rid(), type: 'text', text,
          x: 60, y: 40, w: 840, h: 460, fontSize: 20, bold: false, color: '#222222', align: 'left',
        });
      }
    }

    return { id: rid(), background: '#FFFFFF', elements };
  }

  /* ---------------- 1) Strict scan ---------------- */

  async function scanBuffer(buffer) {
    if (typeof JSZip === 'undefined') throw new Error('JSZip failed to load from CDN.');

    const head = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
    if (head.length < 4 || head[0] !== 0x50 || head[1] !== 0x4b) {
      throw new PptxScanError('Not a ZIP container (missing PK signature).', 'NOT_ZIP');
    }

    let zip;
    try {
      zip = await JSZip.loadAsync(buffer, { checkCRC32: true });
    } catch (err) {
      throw new PptxScanError('ZIP error: ' + (err && err.message ? err.message : err), 'ZIP_CORRUPT');
    }

    if (!zip.file('[Content_Types].xml')) {
      throw new PptxScanError('Missing [Content_Types].xml: not a valid Office package.', 'NOT_PPTX');
    }

    const slideNames = Object.keys(zip.files)
      .filter((n) => SLIDE_RE.test(n))
      .sort((a, b) => parseInt(SLIDE_RE.exec(a)[1], 10) - parseInt(SLIDE_RE.exec(b)[1], 10));
    if (!slideNames.length) {
      throw new PptxScanError('No slides found at ppt/slides/slide[x].xml.', 'NO_SLIDES');
    }

    const parser = new DOMParser();
    const checkXml = (name, xml) => {
      const doc = parser.parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) {
        throw new PptxScanError('XML syntax error in ' + name + '.', 'XML_SYNTAX');
      }
    };

    let presXml = '';
    try {
      // async() reads entry data → this is where CRC32 mismatches surface.
      const ct = await zip.file('[Content_Types].xml').async('string');
      checkXml('[Content_Types].xml', ct);

      const presFile = zip.file('ppt/presentation.xml');
      if (presFile) {
        presXml = await presFile.async('string');
        checkXml('ppt/presentation.xml', presXml);
      }

      const dims = slideDimsFromXml(presXml);
      const slides = [];
      for (const name of slideNames) {
        const xml = await zip.file(name).async('string');
        checkXml(name, xml);
        slides.push(scrapeSlideXml(xml, dims));
      }
      return { slides };
    } catch (err) {
      if (err instanceof PptxScanError) throw err;
      throw new PptxScanError('Archive read error: ' + (err && err.message ? err.message : err), 'ZIP_CORRUPT');
    }
  }

  /* ---------------- 2) Force repair ---------------- */

  /** Inflate raw DEFLATE bytes, keeping whatever decoded before any error. */
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') return null;
    const ds = new DecompressionStream('deflate-raw');
    const writer = ds.writable.getWriter();
    writer.write(bytes).catch(() => {});
    writer.close().catch(() => {});
    const reader = ds.readable.getReader();
    const chunks = [];
    let total = 0;
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
        total += value.length;
      }
    } catch (_) { /* truncated / trailing junk: keep partial output */ }
    const out = new Uint8Array(total);
    let o = 0;
    for (const c of chunks) { out.set(c, o); o += c.length; }
    return out;
  }

  function findSig(u8, from) {
    for (let i = from; i + 3 < u8.length; i++) {
      if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x03 && u8[i + 3] === 0x04) return i;
    }
    return -1;
  }

  /** Tier B: ignore the (possibly destroyed) central directory entirely. */
  async function walkLocalHeaders(buffer) {
    const u8 = new Uint8Array(buffer);
    const dv = new DataView(buffer);
    const found = {};
    let i = findSig(u8, 0);

    while (i !== -1 && i + 30 <= u8.length) {
      const flags = dv.getUint16(i + 6, true);
      const method = dv.getUint16(i + 8, true);
      const compSize = dv.getUint32(i + 18, true);
      const nameLen = dv.getUint16(i + 26, true);
      const extraLen = dv.getUint16(i + 28, true);
      const nameEnd = i + 30 + nameLen;
      if (nameEnd > u8.length) break;

      const name = utf8.decode(u8.subarray(i + 30, nameEnd));
      const dataStart = nameEnd + extraLen;
      const sizeKnown = !(flags & 0x08) && compSize > 0 && dataStart + compSize <= u8.length;
      let dataEnd;
      if (sizeKnown) {
        dataEnd = dataStart + compSize;
      } else {
        const nxt = findSig(u8, dataStart);
        dataEnd = nxt === -1 ? u8.length : nxt;
      }

      if (WANTED_RE.test(name)) {
        const raw = u8.subarray(dataStart, dataEnd);
        try {
          const data = method === 0 ? raw : method === 8 ? await inflateRaw(raw) : null;
          if (data && data.length) found[name] = data;
        } catch (_) { /* skip entry */ }
      }

      const next = findSig(u8, Math.max(dataEnd, i + 4));
      i = next;
    }
    return found;
  }

  async function forceRepair(buffer) {
    const stats = { method: 'none', entriesRecovered: 0, slidesFound: 0, textBlocks: 0 };
    let entries = {};

    // Tier A: JSZip without CRC enforcement, entry-by-entry.
    try {
      const zip = await JSZip.loadAsync(buffer, { checkCRC32: false });
      for (const name of Object.keys(zip.files)) {
        if (!WANTED_RE.test(name)) continue;
        try { entries[name] = await zip.files[name].async('uint8array'); } catch (_) { /* skip */ }
      }
      if (Object.keys(entries).some((n) => SLIDE_RE.test(n))) stats.method = 'jszip-lenient';
    } catch (_) { /* archive structure destroyed → Tier B */ }

    // Tier B: raw local-header walk.
    if (!Object.keys(entries).some((n) => SLIDE_RE.test(n))) {
      entries = await walkLocalHeaders(buffer);
      if (Object.keys(entries).length) stats.method = 'local-header-walk';
    }

    stats.entriesRecovered = Object.keys(entries).length;

    const presBytes = entries['ppt/presentation.xml'];
    const dims = slideDimsFromXml(presBytes ? utf8.decode(presBytes) : '');

    const slideNames = Object.keys(entries)
      .filter((n) => SLIDE_RE.test(n))
      .sort((a, b) => parseInt(SLIDE_RE.exec(a)[1], 10) - parseInt(SLIDE_RE.exec(b)[1], 10));

    const slides = slideNames.map((name) => scrapeSlideXml(utf8.decode(entries[name]), dims));
    stats.slidesFound = slides.length;
    stats.textBlocks = slides.reduce((n, s) => n + s.elements.length, 0);

    return { slides, stats };
  }

  global.PptxRepair = { PptxScanError, scanBuffer, forceRepair, scrapeSlideXml };
})(window);
