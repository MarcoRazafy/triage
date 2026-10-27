/*
 * xlsx-lite.js
 * Lecture et modification de classeurs Excel (.xlsx) sans aucune dépendance.
 *
 * - Ouvre l'archive ZIP avec les API natives du navigateur (DecompressionStream).
 * - Lit les feuilles et les chaînes partagées.
 * - Ajoute ou remplace une colonne en modifiant directement le XML de la feuille :
 *   toute la mise en forme d'origine est conservée (couleurs, filtres, volets figés,
 *   largeurs de colonnes, liens…).
 *
 * Fonctionne dans un navigateur récent, y compris en ouvrant la page en local
 * (file://), et dans Node 18 ou plus.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.XlsxLite = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const utf8 = new TextDecoder('utf-8');
  const utf8enc = new TextEncoder();

  /* ------------------------------------------------------------------ */
  /* Outils                                                              */
  /* ------------------------------------------------------------------ */

  /** Rend la main au navigateur pour qu'il puisse rafraîchir l'affichage. */
  const tick = () => new Promise((resolve) => {
    if (typeof MessageChannel === 'function') {
      const ch = new MessageChannel();
      ch.port1.onmessage = () => { ch.port1.close(); resolve(); };
      ch.port2.postMessage(0);
    } else {
      setTimeout(resolve, 0);
    }
  });

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  async function pipe(bytes, transform) {
    const stream = new Blob([bytes]).stream().pipeThrough(transform);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  const inflateRaw = (b) => pipe(b, new DecompressionStream('deflate-raw'));
  const deflateRaw = (b) => pipe(b, new CompressionStream('deflate-raw'));

  function colToNum(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + ((letters.charCodeAt(i) & ~32) - 64);
    return n;
  }

  function numToCol(n) {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  const ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

  function decodeEntities(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (m, g) => {
      if (g[0] !== '#') return ENTITIES[g];
      const code = g[1] === 'x' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
      return String.fromCodePoint(code);
    });
  }

  /** Texte Excel : entités XML + séquences d'échappement _xHHHH_. */
  function decodeXString(s) {
    s = decodeEntities(s);
    if (s.indexOf('_x') < 0) return s;
    return s.replace(/_x([0-9a-fA-F]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  }

  function escapeXml(s) {
    return String(s)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      .replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }

  function parseAttrs(s) {
    const out = {};
    const re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let m;
    while ((m = re.exec(s))) out[m[1]] = decodeEntities(m[2] !== undefined ? m[2] : m[3]);
    return out;
  }

  function dosDateTime(d) {
    return {
      time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
      date: ((Math.max(1980, d.getFullYear()) - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
    };
  }

  /* ------------------------------------------------------------------ */
  /* ZIP                                                                 */
  /* ------------------------------------------------------------------ */

  function readZip(input) {
    const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const fail = () => { throw new Error("Ce fichier n'est pas un classeur Excel .xlsx valide."); };
    if (u8.length < 22) fail();

    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 22 - 65535); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) fail();

    let count = dv.getUint16(eocd + 10, true);
    let cdOffset = dv.getUint32(eocd + 16, true);
    if (count === 0xFFFF || cdOffset === 0xFFFFFFFF) {
      const loc = eocd - 20;
      if (loc >= 0 && dv.getUint32(loc, true) === 0x07064b50) {
        const z64 = Number(dv.getBigUint64(loc + 8, true));
        if (dv.getUint32(z64, true) === 0x06064b50) {
          count = Number(dv.getBigUint64(z64 + 32, true));
          cdOffset = Number(dv.getBigUint64(z64 + 48, true));
        }
      }
    }

    const entries = [];
    let p = cdOffset;
    for (let i = 0; i < count; i++) {
      if (p + 46 > u8.length || dv.getUint32(p, true) !== 0x02014b50) fail();
      const method = dv.getUint16(p + 10, true);
      const time = dv.getUint16(p + 12, true);
      const date = dv.getUint16(p + 14, true);
      const crc = dv.getUint32(p + 16, true);
      let compSize = dv.getUint32(p + 20, true);
      let size = dv.getUint32(p + 24, true);
      const nLen = dv.getUint16(p + 28, true);
      const xLen = dv.getUint16(p + 30, true);
      const cLen = dv.getUint16(p + 32, true);
      let local = dv.getUint32(p + 42, true);
      const name = utf8.decode(u8.subarray(p + 46, p + 46 + nLen));

      let x = p + 46 + nLen;
      const xEnd = x + xLen;
      while (x + 4 <= xEnd) {
        const id = dv.getUint16(x, true);
        const len = dv.getUint16(x + 2, true);
        if (id === 0x0001) {
          let q = x + 4;
          if (size === 0xFFFFFFFF) { size = Number(dv.getBigUint64(q, true)); q += 8; }
          if (compSize === 0xFFFFFFFF) { compSize = Number(dv.getBigUint64(q, true)); q += 8; }
          if (local === 0xFFFFFFFF) { local = Number(dv.getBigUint64(q, true)); q += 8; }
        }
        x += 4 + len;
      }

      if (dv.getUint32(local, true) !== 0x04034b50) fail();
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      entries.push({ name, method, time, date, crc, compSize, size, raw: u8.subarray(start, start + compSize) });
      p += 46 + nLen + xLen + cLen;
    }

    const byName = new Map();
    const byLower = new Map();
    for (const e of entries) { byName.set(e.name, e); byLower.set(e.name.toLowerCase(), e); }
    const find = (n) => byName.get(n) || byLower.get(String(n).toLowerCase()) || null;

    return {
      entries,
      find,
      async bytes(n) {
        const e = find(n);
        if (!e) return null;
        if (e.method === 0) return e.raw;
        if (e.method === 8) return inflateRaw(e.raw);
        throw new Error('Méthode de compression non gérée (' + e.method + ') pour ' + n);
      },
      async text(n) {
        const b = await this.bytes(n);
        return b ? utf8.decode(b) : null;
      },
    };
  }

  /** Reconstruit l'archive : les parties inchangées sont recopiées telles quelles. */
  async function writeZip(zip, replacements) {
    const parts = [];
    const central = [];
    let offset = 0;
    const now = dosDateTime(new Date());

    for (const e of zip.entries) {
      let { method, crc, compSize, size, raw, time, date } = e;
      if (replacements.has(e.name)) {
        const data = replacements.get(e.name);
        raw = await deflateRaw(data);
        method = 8;
        crc = crc32(data);
        size = data.length;
        compSize = raw.length;
        time = now.time;
        date = now.date;
      }
      if (offset > 0xFFFFFFF0 || size > 0xFFFFFFF0 || compSize > 0xFFFFFFF0) {
        throw new Error('Fichier trop volumineux (plus de 4 Go).');
      }
      const nameBytes = utf8enc.encode(e.name);
      const flags = /[^\x20-\x7e]/.test(e.name) ? 0x0800 : 0;

      const lh = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, flags, true);
      lv.setUint16(8, method, true);
      lv.setUint16(10, time, true);
      lv.setUint16(12, date, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, compSize, true);
      lv.setUint32(22, size, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);
      parts.push(lh, raw);

      const ch = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, flags, true);
      cv.setUint16(10, method, true);
      cv.setUint16(12, time, true);
      cv.setUint16(14, date, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, compSize, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint32(42, offset, true);
      ch.set(nameBytes, 46);
      central.push(ch);

      offset += lh.length + raw.length;
    }

    const n = zip.entries.length;
    if (n > 0xFFFF) throw new Error("Trop d'éléments dans l'archive.");
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new Uint8Array(22);
    const ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, n, true);
    ev.setUint16(10, n, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);
    return new Blob([...parts, ...central, end], { type: XLSX_MIME });
  }

  /* ------------------------------------------------------------------ */
  /* Classeur                                                            */
  /* ------------------------------------------------------------------ */

  function resolvePath(baseDir, target) {
    if (target.startsWith('/')) return target.slice(1);
    const out = [];
    for (const seg of (baseDir + target).split('/')) {
      if (seg === '..') out.pop();
      else if (seg !== '.' && seg !== '') out.push(seg);
    }
    return out.join('/');
  }

  function relsPathFor(part) {
    const i = part.lastIndexOf('/');
    return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels';
  }

  function parseRels(xml) {
    const map = {};
    if (!xml) return map;
    const re = /<(?:\w+:)?Relationship\b([^>]*)>/g;
    let m;
    while ((m = re.exec(xml))) {
      const a = parseAttrs(m[1]);
      map[a.Id] = { target: a.Target || '', type: a.Type || '' };
    }
    return map;
  }

  function textOfRuns(inner) {
    inner = inner.replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, '');
    const re = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g;
    let s = '';
    let m;
    while ((m = re.exec(inner))) s += m[1];
    return decodeXString(s);
  }

  function parseSharedStrings(xml) {
    const out = [];
    const re = /<(?:\w+:)?si\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?si>)/g;
    let m;
    while ((m = re.exec(xml))) out.push(m[1] ? textOfRuns(m[1]) : '');
    return out;
  }

  /**
   * Ouvre un classeur.
   * @param {ArrayBuffer|Uint8Array} buffer contenu du fichier .xlsx
   * @returns {{zip, wbPath, wbXml, sheets: Array<{index,name,state,path}>, sst: string[]}}
   */
  async function openWorkbook(buffer) {
    const zip = readZip(buffer);
    const rootRels = parseRels(await zip.text('_rels/.rels'));
    let wbPath = 'xl/workbook.xml';
    for (const id in rootRels) {
      if (/\/officeDocument$/.test(rootRels[id].type)) wbPath = resolvePath('', rootRels[id].target);
    }
    const wbEntry = zip.find(wbPath);
    if (!wbEntry) throw new Error("Classeur introuvable dans le fichier : est-ce bien un fichier .xlsx ?");
    wbPath = wbEntry.name;
    const wbXml = await zip.text(wbPath);
    const wbDir = wbPath.slice(0, wbPath.lastIndexOf('/') + 1);
    const wbRels = parseRels(await zip.text(relsPathFor(wbPath)));

    const sheets = [];
    const re = /<(?:\w+:)?sheet\b([^>]*?)\/?>/g;
    let m;
    let index = 0;
    while ((m = re.exec(wbXml))) {
      const a = parseAttrs(m[1]);
      const ridKey = Object.keys(a).find((k) => k === 'id' || /:id$/.test(k));
      const rel = ridKey ? wbRels[a[ridKey]] : null;
      const entry = rel ? zip.find(resolvePath(wbDir, rel.target)) : null;
      if (entry && /\/worksheet$/.test(rel.type)) {
        sheets.push({ index, name: a.name || 'Feuille ' + (index + 1), state: a.state || 'visible', path: entry.name });
      }
      index++;
    }
    if (!sheets.length) throw new Error('Aucune feuille de calcul trouvée dans ce classeur.');

    let sst = [];
    let stylesPath = null;
    for (const id in wbRels) {
      if (/\/sharedStrings$/.test(wbRels[id].type)) {
        const xml = await zip.text(resolvePath(wbDir, wbRels[id].target));
        if (xml) sst = parseSharedStrings(xml);
      }
      if (/\/styles$/.test(wbRels[id].type)) stylesPath = resolvePath(wbDir, wbRels[id].target);
    }
    return { zip, wbPath, wbXml, sheets, sst, stylesPath };
  }

  /** Pour chaque style de cellule (index s="…"), vrai s'il a un fond coloré. */
  async function styleFills(book) {
    if (book._fills) return book._fills;
    const fills = [];
    const xml = book.stylesPath ? await book.zip.text(book.stylesPath) : null;
    if (xml) {
      const m = /<(?:\w+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/.exec(xml);
      if (m) {
        const re = /<(?:\w+:)?xf\b([^>]*?)\/?>/g;
        let x;
        while ((x = re.exec(m[1]))) {
          const a = parseAttrs(x[1]);
          fills.push(+(a.fillId || 0) >= 2);
        }
      }
    }
    book._fills = fills;
    return fills;
  }

  async function sheetXml(book, sheet) {
    if (sheet._xml == null) sheet._xml = await book.zip.text(sheet.path);
    return sheet._xml;
  }

  /* ------------------------------------------------------------------ */
  /* Feuille                                                             */
  /* ------------------------------------------------------------------ */

  const ROW_SRC = '<(?:\\w+:)?row\\b([^>]*?)(\\/>|>([\\s\\S]*?)<\\/(?:\\w+:)?row>)';
  const CELL_SRC = '<(?:\\w+:)?c\\b([^>]*?)(\\/>|>([\\s\\S]*?)<\\/(?:\\w+:)?c>)';
  const ROW_NUM = /\br="(\d+)"/;
  const CELL_REF = /\br="([A-Za-z]+)\d*"/;
  const S_ATTR = /\bs="(\d+)"/;
  const T_ATTR = /\bt="(\w+)"/;
  const V_ELEM = /<(?:\w+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?v>/;
  const SPANS_ATTR = /\s+spans="[^"]*"/;

  function sheetBounds(xml) {
    const open = /<((?:\w+:)?)sheetData\b[^>]*?(\/?)>/.exec(xml);
    if (!open) throw new Error('Feuille illisible : zone de données absente.');
    const prefix = open[1];
    const start = open.index + open[0].length;
    if (open[2] === '/') return { prefix, openIdx: open.index, start, end: start, closeEnd: start, selfClosing: true };
    const closeTag = '</' + prefix + 'sheetData>';
    const end = xml.lastIndexOf(closeTag);
    if (end < start) throw new Error('Feuille illisible : zone de données incomplète.');
    return { prefix, openIdx: open.index, start, end, closeEnd: end + closeTag.length, selfClosing: false };
  }

  function cellValue(attrs, inner, sst) {
    const tm = T_ATTR.exec(attrs);
    const t = tm ? tm[1] : 'n';
    if (t === 'inlineStr') return textOfRuns(inner);
    const vm = V_ELEM.exec(inner);
    if (!vm) return null;
    const v = vm[1];
    if (t === 's') {
      const s = sst[parseInt(v, 10)];
      return s == null ? '' : s;
    }
    if (t === 'str') return decodeXString(v);
    if (t === 'b') return v === '1';
    if (t === 'e' || t === 'd') return decodeEntities(v);
    const num = Number(v);
    return Number.isNaN(num) ? decodeEntities(v) : num;
  }

  /** Valeurs d'une ligne : tableau indexé par numéro de colonne (1 = A). */
  function parseRowCells(inner, sst) {
    const cells = [];
    const re = new RegExp(CELL_SRC, 'g');
    let m;
    let prevC = 0;
    while ((m = re.exec(inner))) {
      const rr = CELL_REF.exec(m[1]);
      const c = rr ? colToNum(rr[1]) : prevC + 1;
      prevC = c;
      const v = cellValue(m[1], m[3] || '', sst);
      if (v !== null && v !== '') cells[c] = v;
    }
    return cells;
  }

  /**
   * Lit les lignes d'une feuille.
   * @returns {Promise<Array<{r:number, cells:Array}>>}
   */
  async function readRows(book, sheet, opts) {
    const maxRows = (opts && opts.maxRows) || Infinity;
    const xml = await sheetXml(book, sheet);
    const b = sheetBounds(xml);
    const rows = [];
    if (b.selfClosing) return rows;
    const re = new RegExp(ROW_SRC, 'g');
    re.lastIndex = b.start;
    let m;
    let prevR = 0;
    while ((m = re.exec(xml)) && m.index < b.end) {
      const rm = ROW_NUM.exec(m[1]);
      const r = rm ? +rm[1] : prevR + 1;
      prevR = r;
      if (r > maxRows) break;
      rows.push({ r, cells: parseRowCells(m[3] || '', book.sst) });
      if (opts && opts.onProgress && rows.length % 3000 === 0) {
        opts.onProgress(m.index / xml.length);
        await tick();
      }
    }
    return rows;
  }

  /** Numéro de la dernière ligne de la feuille (0 si inconnu). */
  async function lastRowNumber(book, sheet) {
    const xml = await sheetXml(book, sheet);
    const b = sheetBounds(xml);
    if (b.selfClosing) return 0;
    const from = Math.max(b.start, b.end - 400000);
    const re = /<(?:\w+:)?row\b[^>]*?\br="(\d+)"/g;
    re.lastIndex = from;
    let m;
    let last = 0;
    while ((m = re.exec(xml)) && m.index < b.end) last = +m[1];
    return last;
  }

  /** Plus grand numéro de colonne utilisé dans la feuille. */
  function maxColumn(xml, b) {
    if (b.selfClosing) return 0;
    const re = /<(?:\w+:)?c\b[^>]*?\br="([A-Za-z]+)\d+"/g;
    re.lastIndex = b.start;
    let m;
    let max = 0;
    let maxLetters = '';
    while ((m = re.exec(xml)) && m.index < b.end) {
      const l = m[1];
      if (l.length > maxLetters.length || (l.length === maxLetters.length && l > maxLetters)) {
        maxLetters = l;
        max = colToNum(l);
      }
    }
    return max;
  }

  function scanRow(inner, sst, want, targetCol, wantStyle) {
    const values = [];
    let hit = null;
    let insertAt = -1;
    let lastStyle = null;
    let filled = false;
    let prevC = 0;
    const re = new RegExp(CELL_SRC, 'g');
    let m;
    while ((m = re.exec(inner))) {
      const a = m[1];
      const rr = CELL_REF.exec(a);
      const c = rr ? colToNum(rr[1]) : prevC + 1;
      prevC = c;
      if (c === targetCol) { hit = [m.index, m.index + m[0].length]; continue; }
      if (c > targetCol && insertAt < 0) insertAt = m.index;
      const hasContent = !!m[3];
      if (hasContent) filled = true;
      if (wantStyle && c < targetCol && hasContent) {
        const sm = S_ATTR.exec(a);
        lastStyle = sm ? sm[1] : null;
      }
      if (want.has(c)) values[c] = cellValue(a, m[3] || '', sst);
    }
    if (insertAt < 0) insertAt = inner.length;
    return { values, hit, insertAt, lastStyle, filled };
  }

  function makeCell(p, ref, text, style) {
    const sp = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    const s = style != null ? ' s="' + style + '"' : '';
    return '<' + p + 'c r="' + ref + '"' + s + ' t="inlineStr"><' + p + 'is><' + p + 't' + sp + '>' +
      escapeXml(text) + '</' + p + 't></' + p + 'is></' + p + 'c>';
  }

  function widenDimension(head, col) {
    return head.replace(/(<(?:\w+:)?dimension\b[^>]*?\bref=")([^"]*)(")/, (m, a, ref, z) => {
      const mm = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);
      if (!mm) return m;
      const c2 = mm[3] || mm[1];
      const r2 = mm[4] || mm[2];
      const end = Math.max(colToNum(c2), col);
      return a + mm[1] + mm[2] + ':' + numToCol(end) + r2 + z;
    });
  }

  function ensureColWidth(head, p, col, width) {
    const def = '<' + p + 'col min="' + col + '" max="' + col + '" width="' + width + '" customWidth="1"/>';
    const m = /<((?:\w+:)?)cols>([\s\S]*?)<\/(?:\w+:)?cols>/.exec(head);
    if (!m) return head + '<' + p + 'cols>' + def + '</' + p + 'cols>';
    const inner = m[2];
    const re = /<(?:\w+:)?col\b([^>]*?)\/?>/g;
    let c;
    let insertAt = inner.length;
    while ((c = re.exec(inner))) {
      const a = parseAttrs(c[1]);
      const mn = +a.min;
      const mx = +a.max;
      if (col >= mn && col <= mx) return head;
      if (mn > col) { insertAt = c.index; break; }
    }
    const newInner = inner.slice(0, insertAt) + def + inner.slice(insertAt);
    return head.slice(0, m.index) + '<' + m[1] + 'cols>' + newInner + '</' + m[1] + 'cols>' + head.slice(m.index + m[0].length);
  }

  function widenAutoFilter(part, headerRow, col) {
    const res = { xml: part, changed: false, oldEnd: null, newEnd: null };
    res.xml = part.replace(/(<(?:\w+:)?autoFilter\b[^>]*?\bref=")([A-Z]+)(\d+):([A-Z]+)(\d+)"/, (m, pre, c1, r1, c2, r2) => {
      if (+r1 !== headerRow || colToNum(c2) !== col - 1) return m;
      res.changed = true;
      res.oldEnd = c2;
      res.newEnd = numToCol(col);
      return pre + c1 + r1 + ':' + res.newEnd + r2 + '"';
    });
    return res;
  }

  function widenFilterDatabase(wbXml, sheetIndex, oldEnd, newEnd) {
    return wbXml.replace(/<((?:\w+:)?)definedName\b([^>]*)>([^<]*)<\/(?:\w+:)?definedName>/g, (m, p, attrs, body) => {
      const a = parseAttrs(attrs);
      if (a.name !== '_xlnm._FilterDatabase' || String(a.localSheetId) !== String(sheetIndex)) return m;
      const nb = body.replace(/:\$([A-Z]+)\$(\d+)\s*$/, (mm, c, r) => (c === oldEnd ? ':$' + newEnd + '$' + r : mm));
      return m.replace('>' + body + '<', '>' + nb + '<');
    });
  }

  /**
   * Ajoute (ou remplace) une colonne dans une feuille et reconstruit le classeur.
   *
   * @param book  classeur ouvert avec openWorkbook
   * @param sheet élément de book.sheets
   * @param opts  {
   *   headerRow   n° de la ligne d'en-tête (1 = première ligne),
   *   headerText  texte de l'en-tête de la colonne,
   *   targetCol   n° de colonne à écrire ; 0 = première colonne libre à droite,
   *   readCols    n° des colonnes à lire pour calculer la valeur,
   *   valueFor(rowNumber, values, filled) texte à écrire, ou null pour laisser vide,
   *   onProgress(fraction, phase) facultatif
   * }
   * @returns {Promise<{blob: Blob, targetCol: number}>}
   */
  async function addColumn(book, sheet, opts) {
    const xml = await sheetXml(book, sheet);
    const b = sheetBounds(xml);
    const p = b.prefix;
    const progress = opts.onProgress || (() => {});
    const want = new Set(opts.readCols || []);
    const targetCol = opts.targetCol || maxColumn(xml, b) + 1;
    const L = numToCol(targetCol);

    const pieces = [];
    pieces.push(ensureColWidth(widenDimension(xml.slice(0, b.openIdx), targetCol), p, targetCol, opts.width || 16));
    let cursor = b.openIdx;

    if (!b.selfClosing) {
      const re = new RegExp(ROW_SRC, 'g');
      re.lastIndex = b.start;
      let m;
      let prevR = 0;
      let n = 0;
      while ((m = re.exec(xml)) && m.index < b.end) {
        const rm = ROW_NUM.exec(m[1]);
        const r = rm ? +rm[1] : prevR + 1;
        prevR = r;
        if (r < opts.headerRow) continue;

        const inner = m[3] || '';
        const isHeader = r === opts.headerRow;
        const scan = scanRow(inner, book.sst, want, targetCol, isHeader);
        let cell = null;
        if (isHeader) {
          cell = makeCell(p, L + r, opts.headerText, scan.lastStyle);
        } else {
          const v = opts.valueFor(r, scan.values, scan.filled);
          if (v != null && v !== '') cell = makeCell(p, L + r, String(v), opts.dataStyle);
        }

        if (cell !== null || scan.hit) {
          const newInner = scan.hit
            ? inner.slice(0, scan.hit[0]) + (cell || '') + inner.slice(scan.hit[1])
            : inner.slice(0, scan.insertAt) + cell + inner.slice(scan.insertAt);
          const attrs = m[1].replace(SPANS_ATTR, '').replace(/\s+$/, '');
          pieces.push(xml.slice(cursor, m.index), '<' + p + 'row' + attrs + '>' + newInner + '</' + p + 'row>');
          cursor = m.index + m[0].length;
        }

        if (++n % 3000 === 0) {
          progress(m.index / xml.length, 'rows');
          await tick();
        }
      }
    }

    const filter = widenAutoFilter(xml.slice(b.closeEnd), opts.headerRow, targetCol);
    pieces.push(xml.slice(cursor, b.closeEnd), filter.xml);

    progress(1, 'zip');
    await tick();
    const sheetBytes = new Uint8Array(await new Blob(pieces).arrayBuffer());
    const replacements = new Map([[sheet.path, sheetBytes]]);
    if (filter.changed) {
      const wb2 = widenFilterDatabase(book.wbXml, sheet.index, filter.oldEnd, filter.newEnd);
      if (wb2 !== book.wbXml) replacements.set(book.wbPath, utf8enc.encode(wb2));
    }
    const blob = await writeZip(book.zip, replacements);
    return { blob, targetCol };
  }

  /** Cellules fusionnées de la feuille : [{c1, r1, c2, r2}]. */
  async function readMerges(book, sheet) {
    const xml = await sheetXml(book, sheet);
    const b = sheetBounds(xml);
    const out = [];
    const re = /<(?:\w+:)?mergeCell\b[^>]*?\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g;
    re.lastIndex = b.closeEnd;
    let m;
    while ((m = re.exec(xml))) out.push({ c1: colToNum(m[1]), r1: +m[2], c2: colToNum(m[3]), r2: +m[4] });
    return out;
  }

  /** Nombre de lignes contenant des données après la ligne d'en-tête. */
  async function countDataRows(book, sheet, headerRow) {
    const xml = await sheetXml(book, sheet);
    const b = sheetBounds(xml);
    if (b.selfClosing) return 0;
    const re = new RegExp(ROW_SRC, 'g');
    re.lastIndex = b.start;
    let m;
    let prevR = 0;
    let n = 0;
    while ((m = re.exec(xml)) && m.index < b.end) {
      const rm = ROW_NUM.exec(m[1]);
      const r = rm ? +rm[1] : prevR + 1;
      prevR = r;
      if (r > headerRow && m[3] && /<(?:\w+:)?(?:v|is)[\s>]/.test(m[3])) n++;
    }
    return n;
  }

  /**
   * Remplace toutes les lignes situées sous la ligne d'en-tête par de nouvelles lignes.
   * Chaque ligne est un tableau indexé par numéro de colonne (1 = A) : nombre, texte ou vide.
   * Les formats sont repris de la première ligne existante sous l'en-tête,
   * à défaut du style de colonne. Renvoie le nouveau XML, sans toucher au classeur chargé.
   */
  async function fillSheetXml(book, sheet, headerRow, rows) {
    const xml = await sheetXml(book, sheet);
    const b = sheetBounds(xml);
    const p = b.prefix;
    const styles = [];
    const explicit = [];
    const fills = await styleFills(book);

    // Style de colonne, sauf s'il a un fond coloré : une colonne grisée dans le modèle vide
    // ne doit pas griser les données écrites.
    const colsM = /<(?:\w+:)?cols>([\s\S]*?)<\/(?:\w+:)?cols>/.exec(xml.slice(0, b.openIdx));
    if (colsM) {
      const re = /<(?:\w+:)?col\b([^>]*?)\/?>/g;
      let c;
      while ((c = re.exec(colsM[1]))) {
        const a = parseAttrs(c[1]);
        if (a.style == null || fills[+a.style]) continue;
        for (let k = +a.min; k <= Math.min(+a.max, 512); k++) styles[k] = a.style;
      }
    }

    const kept = [];
    let replaced = 0;
    let styleRowSeen = false;
    if (!b.selfClosing) {
      const re = new RegExp(ROW_SRC, 'g');
      re.lastIndex = b.start;
      let m;
      let prevR = 0;
      while ((m = re.exec(xml)) && m.index < b.end) {
        const rm = ROW_NUM.exec(m[1]);
        const r = rm ? +rm[1] : prevR + 1;
        prevR = r;
        if (r <= headerRow) { kept.push(m[0]); continue; }
        const inner = m[3] || '';
        if (!styleRowSeen) {
          styleRowSeen = true;
          const cre = new RegExp(CELL_SRC, 'g');
          let c;
          let prevC = 0;
          while ((c = cre.exec(inner))) {
            const rr = CELL_REF.exec(c[1]);
            const col = rr ? colToNum(rr[1]) : prevC + 1;
            prevC = col;
            const sm = S_ATTR.exec(c[1]);
            if (sm && !fills[+sm[1]]) { styles[col] = sm[1]; explicit[col] = sm[1]; }
          }
        }
        if (/<(?:\w+:)?(?:v|is)[\s>]/.test(inner)) replaced++;
      }
    }

    // Colonne sans style utilisable : on reprend celui de la colonne voisine la plus proche,
    // pour que les cellules vides gardent les mêmes bordures que le reste de la ligne.
    const width = rows.reduce((m, row) => Math.max(m, row.length - 1), 0);
    for (let c = 1; c <= width; c++) {
      if (styles[c] != null) continue;
      for (let d = 1; d <= width && styles[c] == null; d++) {
        if (explicit[c - d] != null) styles[c] = explicit[c - d];
        else if (explicit[c + d] != null) styles[c] = explicit[c + d];
      }
    }

    const out = [];
    let maxCol = 0;
    for (let i = 0; i < rows.length; i++) {
      const r = headerRow + 1 + i;
      const row = rows[i];
      let cells = '';
      for (let c = 1; c <= width; c++) {
        const v = row[c];
        if (v == null || v === '') {
          if (styles[c] != null) cells += '<' + p + 'c r="' + numToCol(c) + r + '" s="' + styles[c] + '"/>';
          continue;
        }
        const ref = numToCol(c) + r;
        if (typeof v === 'number') {
          if (!Number.isFinite(v)) continue;
          const s = styles[c] != null ? ' s="' + styles[c] + '"' : '';
          cells += '<' + p + 'c r="' + ref + '"' + s + '><' + p + 'v>' + v + '</' + p + 'v></' + p + 'c>';
        } else {
          cells += makeCell(p, ref, String(v), styles[c]);
        }
        if (c > maxCol) maxCol = c;
      }
      out.push('<' + p + 'row r="' + r + '">' + cells + '</' + p + 'row>');
    }

    const sheetData = '<' + p + 'sheetData>' + kept.join('') + out.join('') + '</' + p + 'sheetData>';
    const before = xml.slice(0, b.openIdx).replace(/(<(?:\w+:)?dimension\b[^>]*?\bref=")([^"]*)(")/, (m0, a, ref, z) => {
      const mm = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);
      const endCol = Math.max(maxCol, mm ? colToNum(mm[3] || mm[1]) : 1);
      return a + 'A1:' + numToCol(endCol) + Math.max(headerRow + rows.length, headerRow) + z;
    });
    return { xml: before + sheetData + xml.slice(b.closeEnd), replaced };
  }

  /** Reconstruit le classeur avec des parties remplacées : Map chemin → texte XML. */
  // (les styles sont mis en cache dans book._fills par styleFills)
  async function writeWorkbook(book, parts) {
    const repl = new Map();
    for (const [path, content] of parts) {
      const entry = book.zip.find(path);
      repl.set(entry ? entry.name : path, typeof content === 'string' ? utf8enc.encode(content) : content);
    }
    return writeZip(book.zip, repl);
  }

  return {
    XLSX_MIME,
    readMerges,
    countDataRows,
    fillSheetXml,
    writeWorkbook,
    openWorkbook,
    readRows,
    lastRowNumber,
    addColumn,
    colToNum,
    numToCol,
    crc32,
    readZip,
    writeZip,
    tick,
  };
});
