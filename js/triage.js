/*
 * triage.js
 * Logique métier : lecture des référentiels, comparaison des EAN
 * et ajout de la colonne « produit » dans le fichier brut.
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const X = isNode ? require('./xlsx-lite.js') : root.XlsxLite;
  const api = factory(X);
  if (isNode) module.exports = api;
  else root.Triage = api;
})(typeof self !== 'undefined' ? self : this, function (X) {
  'use strict';

  /** Les catégories suivies pour le client, dans l'ordre d'affichage. */
  const CATEGORIES = ['Fromage LS', 'CORPS GRAS', 'LAIT', 'CREME'];
  const NOT_FOUND = 'Non trouvé';
  const NO_EAN = 'Sans EAN';

  /** Minuscules, sans accents, sans « _ » ni « - », espaces simplifiés. */
  function fold(s) {
    return String(s == null ? '' : s)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase()
      .replace(/[_\-]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /**
   * Ramène un EAN à une forme comparable : chiffres seuls, sans zéros en tête.
   * « 0000090027819 », 90027819 et « 90027819 » donnent tous « 90027819 ».
   */
  function normalizeEan(v) {
    if (v == null || typeof v === 'boolean') return '';
    let s;
    if (typeof v === 'number') {
      if (!Number.isFinite(v) || v <= 0) return '';
      s = Math.round(v).toString();
    } else {
      s = String(v).trim();
      if (/^\d+[.,]0+$/.test(s)) s = s.replace(/[.,]0+$/, '');
    }
    return s.replace(/\D/g, '').replace(/^0+/, '');
  }

  function displayEan(v) {
    if (v == null) return '';
    if (typeof v === 'number') return Number.isFinite(v) ? Math.round(v).toString() : '';
    return String(v).trim();
  }

  /** Devine la catégorie à partir d'un texte (segmentation ou nom de fichier). */
  function categoryFromText(s) {
    const f = fold(s);
    if (!f) return null;
    if (/fromage/.test(f)) return 'Fromage LS';
    if (/corps gras|beurre|margarine|matiere grasse|graisse/.test(f)) return 'CORPS GRAS';
    if (/creme/.test(f)) return 'CREME';
    if (/\blait\b|boisson vegetale/.test(f)) return 'LAIT';
    return null;
  }

  const EAN_HEADER = /^(ean|ean ?13|ean ?8|code ean|ean code|code ean ?13|gencod|gencode|code barres?|codebarres?)$/;
  const LABEL_HEADERS = [
    /^libelle (de l )?article$/,
    /^designation( article)?$/,
    /^libelle reference$/,
    /^libelle produit$/,
    /^produits$/,
    /^libelle$/,
  ];

  /** Cherche la ligne d'en-tête : la première qui contient une colonne « EAN ». */
  function findHeader(rows, maxScan) {
    const limit = maxScan || 40;
    for (const row of rows) {
      if (row.r > limit) break;
      for (let c = 1; c < row.cells.length; c++) {
        const v = row.cells[c];
        if (typeof v === 'string' && EAN_HEADER.test(fold(v))) {
          return { headerRow: row.r, eanCol: c, headers: row.cells.slice() };
        }
      }
    }
    return null;
  }

  function findColumn(headers, patterns, exclude) {
    for (const re of patterns) {
      for (let c = 1; c < headers.length; c++) {
        const v = headers[c];
        if (typeof v !== 'string') continue;
        const f = fold(v);
        if (exclude && f === exclude) continue;
        if (re.test(f)) return c;
      }
    }
    return 0;
  }

  function findLabelColumn(headers, outName) {
    return findColumn(headers, LABEL_HEADERS, fold(outName || 'produit'));
  }

  /** Colonne de sortie déjà présente (fichier déjà trié une première fois). */
  function findOutputColumn(headers, outName) {
    const target = fold(outName || 'produit');
    for (let c = 1; c < headers.length; c++) {
      if (typeof headers[c] === 'string' && fold(headers[c]) === target) return c;
    }
    return 0;
  }

  /* ------------------------------------------------------------------ */
  /* Référentiels                                                        */
  /* ------------------------------------------------------------------ */

  function buildReference(name, sheetsRows) {
    const fileCat = categoryFromText(name);
    const items = new Map();
    const autoCats = new Map();
    let duplicates = 0;
    let emptyEan = 0;
    let sheetsUsed = 0;

    for (const rows of sheetsRows) {
      const h = findHeader(rows);
      if (!h) continue;
      sheetsUsed++;
      const seg0Col = findColumn(h.headers, [/^segmentation ?0$/]);
      const labelCol = findLabelColumn(h.headers);
      for (const row of rows) {
        if (row.r <= h.headerRow) continue;
        const ean = normalizeEan(row.cells[h.eanCol]);
        if (!ean) {
          if (row.cells.some((v) => v != null && v !== '')) emptyEan++;
          continue;
        }
        if (items.has(ean)) { duplicates++; continue; }
        const seg0 = seg0Col ? row.cells[seg0Col] : null;
        const cat = categoryFromText(seg0) || fileCat || (seg0 ? String(seg0).trim().toUpperCase() : 'NON CLASSÉ');
        items.set(ean, { cat, label: labelCol ? String(row.cells[labelCol] || '') : '' });
        autoCats.set(cat, (autoCats.get(cat) || 0) + 1);
      }
    }
    if (!sheetsUsed) throw new Error('Aucune colonne « ean » trouvée dans ce fichier.');
    return { name, items, count: items.size, autoCats, duplicates, emptyEan, override: '' };
  }

  /** Lit un fichier référentiel (objet File du navigateur ou équivalent). */
  async function loadReference(file) {
    const book = await X.openWorkbook(await file.arrayBuffer());
    const all = [];
    for (const sheet of book.sheets) {
      if (sheet.state !== 'visible' && book.sheets.length > 1) continue;
      all.push(await X.readRows(book, sheet));
      sheet._xml = null;
    }
    return buildReference(file.name, all);
  }

  /** Catégorie détectée pour un référentiel (pour l'affichage). */
  function autoLabel(ref) {
    const keys = [...ref.autoCats.keys()];
    if (keys.length === 1) return keys[0];
    if (!keys.length) return '—';
    return 'plusieurs';
  }

  /**
   * Fusionne les référentiels en un seul index EAN → catégorie.
   * En cas de conflit, le premier référentiel de la liste l'emporte.
   */
  function buildIndex(refs) {
    const index = new Map();
    const conflicts = [];
    for (const ref of refs) {
      for (const [ean, it] of ref.items) {
        const cat = ref.override || it.cat;
        const prev = index.get(ean);
        if (prev) {
          if (prev.cat !== cat) conflicts.push({ ean, kept: prev, dropped: { cat, source: ref.name } });
          continue;
        }
        index.set(ean, { cat, source: ref.name, label: it.label });
      }
    }
    return { index, conflicts };
  }

  /* ------------------------------------------------------------------ */
  /* Triage du fichier brut                                              */
  /* ------------------------------------------------------------------ */

  /**
   * @param book   classeur brut ouvert avec XlsxLite.openWorkbook
   * @param sheet  feuille à traiter
   * @param cfg    { headerRow, eanCol, labelCol, outName, notFoundText, targetCol }
   * @param index  Map EAN → { cat, source }
   */
  async function runTriage(book, sheet, cfg, index, onProgress) {
    const stats = { rows: 0, byCat: new Map(), notFound: 0, noEan: 0 };
    const samples = new Map();
    const SAMPLE_MAX = 300;
    const push = (key, item) => {
      let list = samples.get(key);
      if (!list) samples.set(key, (list = []));
      if (list.length < SAMPLE_MAX) list.push(item);
    };
    const notFoundText = cfg.notFoundText ? String(cfg.notFoundText) : null;
    const readCols = [cfg.eanCol];
    if (cfg.labelCol) readCols.push(cfg.labelCol);

    const res = await X.addColumn(book, sheet, {
      headerRow: cfg.headerRow,
      headerText: cfg.outName || 'produit',
      targetCol: cfg.targetCol || 0,
      readCols,
      onProgress,
      valueFor(r, values, filled) {
        if (!filled) return null;
        stats.rows++;
        const raw = values[cfg.eanCol];
        const label = cfg.labelCol && values[cfg.labelCol] != null ? String(values[cfg.labelCol]).trim() : '';
        const ean = normalizeEan(raw);
        if (!ean) {
          stats.noEan++;
          push(NO_EAN, { r, ean: displayEan(raw), label, cat: NO_EAN });
          return notFoundText;
        }
        const hit = index.get(ean);
        if (!hit) {
          stats.notFound++;
          push(NOT_FOUND, { r, ean: displayEan(raw), label, cat: NOT_FOUND });
          return notFoundText;
        }
        stats.byCat.set(hit.cat, (stats.byCat.get(hit.cat) || 0) + 1);
        push(hit.cat, { r, ean: displayEan(raw), label, cat: hit.cat });
        return hit.cat;
      },
    });
    return { blob: res.blob, targetCol: res.targetCol, stats, samples };
  }

  return {
    CATEGORIES,
    NOT_FOUND,
    NO_EAN,
    fold,
    normalizeEan,
    categoryFromText,
    findHeader,
    findLabelColumn,
    findOutputColumn,
    buildReference,
    loadReference,
    autoLabel,
    buildIndex,
    runTriage,
  };
});
