/*
 * modele.js
 * Logique du remplissage du modèle mensuel à partir d'un fichier trié :
 * correspondance des colonnes, détection du mois, règles de calcul,
 * regroupement des lignes et écriture dans les onglets 01 à 12.
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const X = isNode ? require('./xlsx-lite.js') : root.XlsxLite;
  const T = isNode ? require('./triage.js') : root.Triage;
  const api = factory(X, T);
  if (isNode) module.exports = api;
  else root.Modele = api;
})(typeof self !== 'undefined' ? self : this, function (X, T) {
  'use strict';

  const MONTHS = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'];
  const MONTH_LABELS = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
  const MONTH_ABBR = { janv: 1, fev: 2, fevr: 2, avr: 4, juil: 7, sept: 9, oct: 10, nov: 11, dec: 12 };
  const pad2 = (n) => String(n).padStart(2, '0');
  const fold = T.fold;

  /* ------------------------------------------------------------------ */
  /* Mois et année                                                       */
  /* ------------------------------------------------------------------ */

  function monthFromText(s) {
    const f = fold(s);
    if (!f) return 0;
    for (let i = 0; i < 12; i++) {
      if (new RegExp('(^|[^a-z])' + MONTHS[i] + '([^a-z]|$)').test(f)) return i + 1;
    }
    for (const k in MONTH_ABBR) {
      if (new RegExp('(^|[^a-z])' + k + '\\.?([^a-z]|$)').test(f)) return MONTH_ABBR[k];
    }
    return 0;
  }

  /** « FRAIS STATS AVRIL 2026 » donne 4, « 202609 - LACTALIS » donne 9. */
  function monthFromFileName(name) {
    const base = String(name || '').replace(/\.[^.]+$/, '');
    const m = monthFromText(base);
    if (m) return m;
    const ym = /(?:^|\D)20\d{2}[-_ .]?(0[1-9]|1[0-2])(?:\D|$)/.exec(base);
    return ym ? +ym[1] : 0;
  }

  function yearFromText(s) {
    s = String(s || '');
    let m = /(?:^|\D)(20\d{2})(?:0[1-9]|1[0-2])(?:\D|$)/.exec(s);
    if (m) return +m[1];
    m = /(?:^|\D)(20\d{2})(?:\D|$)/.exec(s);
    return m ? +m[1] : 0;
  }

  /** Mois d'une cellule : 7, « 07 », date Excel, « juillet-26 », « 2026-07-01 », « 07/2026 »… */
  function monthFromValue(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') {
      if (Number.isInteger(v) && v >= 1 && v <= 12) return v;
      if (v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)).getUTCMonth() + 1;
      if (Number.isInteger(v) && v >= 190001 && v <= 210012 && v % 100 >= 1 && v % 100 <= 12) return v % 100;
      return 0;
    }
    const s = String(v).trim();
    let m = /^(\d{4})[-/.](\d{1,2})(?:[-/.]\d{1,2})?(?:[ T].*)?$/.exec(s);
    if (m) return +m[2] >= 1 && +m[2] <= 12 ? +m[2] : 0;
    m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(s);
    if (m) return +m[2] >= 1 && +m[2] <= 12 ? +m[2] : 0;
    m = /^(\d{1,2})[-/.](\d{4})$/.exec(s);
    if (m) return +m[1] >= 1 && +m[1] <= 12 ? +m[1] : 0;
    if (/^\d{1,2}$/.test(s)) return +s >= 1 && +s <= 12 ? +s : 0;
    if (/^\d{6}$/.test(s)) return monthFromValue(+s);
    return monthFromText(s);
  }

  /* ------------------------------------------------------------------ */
  /* Valeurs                                                             */
  /* ------------------------------------------------------------------ */

  function toNumber(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (v == null || typeof v === 'boolean') return null;
    const s = String(v).replace(/[\s  €]/g, '').replace(',', '.');
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : null;
  }

  /** Traduit une valeur en O ou N : PROMO, O, Oui, 1, montant positif… donnent O. */
  function toFlag(v) {
    if (v == null || v === '') return 'N';
    if (typeof v === 'number') return v > 0 ? 'O' : 'N';
    if (typeof v === 'boolean') return v ? 'O' : 'N';
    const f = fold(v);
    if (/^(o|oui|y|yes|promo|promotion|en promo|p|vrai|true|x)$/.test(f)) return 'O';
    if (/^(n|non|no|fdr|hors promo|normal|np|faux|false)$/.test(f)) return 'N';
    const n = toNumber(v);
    return n != null && n > 0 ? 'O' : 'N';
  }

  /** EAN écrit en nombre pour que le format 13 chiffres du modèle s'applique. */
  function eanOut(v) {
    if (v == null) return null;
    if (typeof v === 'number') return Number.isFinite(v) ? Math.round(v) : null;
    const s = String(v).trim();
    if (!s) return null;
    return /^\d{1,15}$/.test(s) ? Number(s) : s;
  }

  const round6 = (x) => Math.round(x * 1e6) / 1e6;

  /* ------------------------------------------------------------------ */
  /* Description des colonnes                                            */
  /* ------------------------------------------------------------------ */

  /** Rôle d'une colonne du modèle d'après son titre. */
  function templateRole(header) {
    const f = fold(header);
    if (/^(code )?ean/.test(f)) return 'ean';
    if (/promo/.test(f) && /vendu|ca\b|montant/.test(f)) return 'promoVendue';
    if (/promo/.test(f)) return 'promo';
    if (/site|magasin/.test(f)) return 'site';
    if (/rayon/.test(f)) return 'rayon';
    if (/famille/.test(f)) return 'famille';
    if (/fournisseur/.test(f)) return 'fournisseur';
    if (/article|designation|produit/.test(f)) return 'article';
    if (/^(qte|quantite)/.test(f)) return 'qte';
    if (/montant|^ca\b|chiffre/.test(f)) return 'montant';
    return 'autre';
  }

  /**
   * Colonnes du fichier source : lettre, titre, sur-titre (ligne au-dessus, cellules fusionnées
   * comprises). Le sur-titre distingue par exemple « Total » des quantités et « Total » du CA.
   */
  function describeColumns(rows, headerRow, merges) {
    const header = (rows.find((r) => r.r === headerRow) || { cells: [] }).cells;
    const above = (rows.find((r) => r.r === headerRow - 1) || { cells: [] }).cells;
    let n = header.length - 1;
    for (const row of rows) if (row.r > headerRow && row.cells.length - 1 > n) n = row.cells.length - 1;
    const cols = [];
    for (let c = 1; c <= n; c++) {
      let sup = above[c];
      if (sup == null) {
        const mg = (merges || []).find((g) => g.r1 === headerRow - 1 && c >= g.c1 && c <= g.c2);
        if (mg) sup = above[mg.c1];
      }
      const h = header[c] == null ? '' : String(header[c]).trim();
      const s = sup == null ? '' : String(sup).trim();
      cols.push({ col: c, letter: X.numToCol(c), header: h, sup: s, fh: fold(h), fs: fold(s) });
    }
    return cols;
  }

  /** Fichier avec un mois par colonne (sur-titres janvier, février…) : groupes de colonnes par mois. */
  function detectMonthGroups(cols) {
    const groups = new Map();
    for (const c of cols) {
      const m = c.fs ? monthFromText(c.sup) : 0;
      if (!m) continue;
      if (!groups.has(m)) groups.set(m, []);
      groups.get(m).push(c.col);
    }
    if (groups.size < 2) return null;
    const list = [...groups.entries()].sort((a, b) => a[0] - b[0]).map(([month, cs]) => ({ month, cols: cs }));
    const first = list[0];
    return {
      groups: list,
      offsets: first.cols.map((c, i) => ({ offset: i, header: cols[c - 1].header })),
    };
  }

  const ROLE_PATTERNS = {
    site: [/^libelle site$/, /^libelle magasin$/, /^nom magasin$/, /^magasins?$/, /^site$/, /^libelle mag\.?$/, /^mag\.?$/, /^code site$/],
    rayon: [/^libelle rayon$/, /^rayons?$/, /^libelle ray\.?$/],
    famille: [/^libelle famille$/, /^libelle fam\.?$/, /^familles?$/],
    fournisseur: [/^libelle fournisseur$/, /^fournisseurs?$/, /^libelle fourn\.?$/, /^nom fournisseur$/],
    article: [/^libelle (de l )?article$/, /^designation( article)?$/, /^libelle reference$/, /^libelle produit$/, /^produits$/],
    qte: [/^(quantite|quantites|qte)( vendues?)?( total)?$/],
    montant: [/^ca ttc( total)?$/, /^montant( vente| ttc)?$/, /^ca$/],
    promoVendue: [/^ca ttc promo$/, /^montant promo$/],
    promo: [/^promo$/, /^promotion$/, /^en promo$/],
  };

  function findByPatterns(cols, patterns, exclude) {
    for (const useSup of [false, true]) {
      for (const re of patterns) {
        for (const c of cols) {
          if (exclude && exclude.has(c.col)) continue;
          const label = useSup ? (c.fs ? c.fs + ' ' + c.fh : '') : c.fh;
          if (label && re.test(label)) return c.col;
        }
      }
    }
    return 0;
  }

  /**
   * Analyse un fichier source déjà lu.
   * @returns {{headerRow, eanCol, catCol, cols, wide, fileMonth, signature}}
   */
  function analyzeSource(fileName, rows, merges, outName) {
    const found = T.findHeader(rows);
    const headerRow = found ? found.headerRow : (rows[0] ? rows[0].r : 1);
    return analyzeWithHeader(fileName, rows, merges, headerRow, outName);
  }

  function analyzeWithHeader(fileName, rows, merges, headerRow, outName) {
    const cols = describeColumns(rows, headerRow, merges);
    const header = [];
    for (const c of cols) header[c.col] = c.header;
    const h = T.findHeader([{ r: 1, cells: header }]);
    const catCol = T.findOutputColumn(header, outName || 'produit');
    const wide = detectMonthGroups(cols);
    return {
      fileName,
      headerRow,
      cols,
      eanCol: h ? h.eanCol : 0,
      catCol,
      wide,
      fileMonth: monthFromFileName(fileName),
      fileYear: yearFromText(fileName),
      signature: cols.map((c) => c.fh).join('|'),
    };
  }

  /** Valeurs distinctes de la colonne catégorie avec leur nombre de lignes. */
  function categoryCounts(rows, headerRow, catCol) {
    const counts = new Map();
    let none = 0;
    for (const row of rows) {
      if (row.r <= headerRow || !row.cells.length) continue;
      const v = catCol ? row.cells[catCol] : null;
      const k = v == null ? '' : String(v).trim();
      if (!k) { none++; continue; }
      counts.set(k, (counts.get(k) || 0) + 1);
    }
    return { counts, none };
  }

  /* ------------------------------------------------------------------ */
  /* Suggestions de correspondance                                       */
  /* ------------------------------------------------------------------ */

  /**
   * Règle de remplissage d'une colonne du modèle :
   *   {type:'empty'} | {type:'fixed', value} | {type:'category'} | {type:'col', col}
   *   {type:'flag', col}       O si la valeur est PROMO, O, Oui ou un nombre positif, sinon N
   *   {type:'monthly', offset} n-ième colonne du mois, pour les fichiers à un mois par colonne
   * La colonne promo (O/N) du modèle n'est jamais remplie : elle est saisie à la main dans Excel.
   */
  function suggestMapping(tplCols, src) {
    const map = {};
    const used = new Set([src.catCol].filter(Boolean));
    const wideCols = new Set();
    if (src.wide) for (const g of src.wide.groups) for (const c of g.cols) wideCols.add(c);
    const toMonthly = (col) => {
      if (!src.wide || !wideCols.has(col)) return null;
      for (const g of src.wide.groups) {
        const i = g.cols.indexOf(col);
        if (i >= 0) return { type: 'monthly', offset: i };
      }
      return null;
    };

    for (const t of tplCols) {
      let spec = { type: 'empty' };
      switch (t.role) {
        case 'ean':
          if (src.eanCol) spec = { type: 'col', col: src.eanCol };
          break;
        case 'famille':
          if (src.catCol) spec = { type: 'col', col: src.catCol };
          else {
            const c = findByPatterns(src.cols, ROLE_PATTERNS.famille, used);
            if (c) spec = { type: 'col', col: c };
          }
          break;
        case 'qte':
        case 'montant': {
          const c = findByPatterns(src.cols, ROLE_PATTERNS[t.role], used);
          if (c) spec = toMonthly(c) || { type: 'col', col: c };
          break;
        }
        case 'promo':
          // Toujours laissée vide : l'utilisateur la saisit à la main dans Excel.
          break;
        case 'promoVendue': {
          const c = findByPatterns(src.cols, ROLE_PATTERNS.promoVendue, used);
          if (c) spec = { type: 'col', col: c };
          break;
        }
        default:
          if (ROLE_PATTERNS[t.role]) {
            const c = findByPatterns(src.cols, ROLE_PATTERNS[t.role], used);
            if (c) spec = { type: 'col', col: c };
          }
      }
      map[t.col] = spec;
    }
    return map;
  }

  function suggestMonth(src) {
    const c = src.cols.find((x) => /^(mois|month|periode|mois annee|annee mois)$/.test(x.fh));
    if (c) return { type: 'col', col: c.col };
    if (src.wide) return { type: 'wide' };
    if (src.fileMonth) return { type: 'fixed', month: src.fileMonth, fromFile: true };
    return { type: 'fixed', month: 0 };
  }

  /* ------------------------------------------------------------------ */
  /* Construction des lignes du modèle                                   */
  /* ------------------------------------------------------------------ */

  /**
   * @param src  résultat de analyzeSource
   * @param rows lignes lues du fichier source
   * @param cfg  { tplCols, mapping, month, keep: Set|null, merge: bool, limit? }
   * @returns {{ byMonth: Map<number, Array>, stats }}
   */
  function buildOutput(src, rows, cfg) {
    const tpl = cfg.tplCols;
    const width = tpl.reduce((m, t) => Math.max(m, t.col), 0) + 1;
    const roleCol = {};
    for (const t of tpl) if (!(t.role in roleCol)) roleCol[t.role] = t.col;
    const numeric = new Set(tpl.filter((t) => ['qte', 'montant', 'promoVendue'].includes(t.role)).map((t) => t.col));
    const monthSpec = cfg.month;
    const groups = monthSpec.type === 'wide' && src.wide ? src.wide.groups : null;

    const byMonth = new Map();
    const index = new Map();
    const stats = { read: 0, kept: 0, skippedCategory: 0, noMonth: 0, written: 0, merged: 0, months: new Map() };

    const emit = (month, out) => {
      if (!byMonth.has(month)) byMonth.set(month, []);
      const list = byMonth.get(month);
      if (cfg.merge) {
        const ean = roleCol.ean ? out[roleCol.ean] : null;
        const art = roleCol.article ? out[roleCol.article] : null;
        const site = roleCol.site ? out[roleCol.site] : null;
        const key = month + '|' + (ean != null && ean !== '' ? 'e' + ean : 'a' + (art || '')) + '|' + (site == null ? '' : site);
        const prev = index.get(key);
        if (prev) {
          stats.merged++;
          for (let c = 1; c < width; c++) {
            const v = out[c];
            if (numeric.has(c)) {
              if (v != null) prev[c] = round6((prev[c] || 0) + v);
            } else if (c === roleCol.promo) {
              if (v === 'O') prev[c] = 'O';
            } else if ((prev[c] == null || prev[c] === '') && v != null && v !== '') {
              prev[c] = v;
            }
          }
          return;
        }
        index.set(key, out);
      }
      list.push(out);
    };

    const evaluate = (cells, cat, group) => {
      const out = new Array(width).fill(null);
      for (const t of tpl) {
        const sp = cfg.mapping[t.col] || { type: 'empty' };
        let v = null;
        switch (sp.type) {
          case 'col': v = cells[sp.col]; break;
          case 'monthly': v = group ? cells[group.cols[sp.offset]] : null; break;
          case 'fixed': v = sp.value == null || sp.value === '' ? null : sp.value; break;
          case 'category': v = cat; break;
          case 'flag': v = toFlag(cells[sp.col]); break;
          default: v = null;
        }
        if (v != null && numeric.has(t.col)) v = toNumber(v);
        else if (t.role === 'ean') v = eanOut(v);
        else if (typeof v === 'string') v = v.trim();
        out[t.col] = v === '' ? null : v;
      }
      if (roleCol.promo) out[roleCol.promo] = null;
      return out;
    };

    for (const row of rows) {
      if (row.r <= src.headerRow || !row.cells.length) continue;
      if (cfg.limit && stats.written >= cfg.limit) break;
      stats.read++;
      const catRaw = src.catCol ? row.cells[src.catCol] : null;
      const cat = catRaw == null ? '' : String(catRaw).trim();
      if (cfg.keep && !cfg.keep.has(cat)) { stats.skippedCategory++; continue; }
      stats.kept++;

      if (groups) {
        for (const g of groups) {
          const out = evaluate(row.cells, cat || null, g);
          const q = roleCol.qte ? out[roleCol.qte] : null;
          const mt = roleCol.montant ? out[roleCol.montant] : null;
          if (!q && !mt) continue;
          stats.written++;
          emit(g.month, out);
        }
        continue;
      }

      let month = 0;
      if (monthSpec.type === 'col') month = monthFromValue(row.cells[monthSpec.col]);
      else if (monthSpec.type === 'fixed') month = monthSpec.month || 0;
      if (!month) { stats.noMonth++; continue; }
      stats.written++;
      emit(month, evaluate(row.cells, cat || null, null));
    }

    for (const [m, list] of byMonth) stats.months.set(m, list.length);
    return { byMonth, stats };
  }

  /** Valeurs distinctes de la colonne promo et leur traduction en O ou N. */
  function flagPreview(rows, headerRow, col, max) {
    const seen = new Map();
    for (const row of rows) {
      if (row.r <= headerRow) continue;
      const v = row.cells[col];
      if (typeof v === 'number') { seen.set('nombre', 'O si supérieur à 0'); continue; }
      const k = v == null ? '(vide)' : String(v).trim() || '(vide)';
      if (!seen.has(k)) seen.set(k, toFlag(v));
      if (seen.size >= (max || 8)) break;
    }
    return seen;
  }

  /* ------------------------------------------------------------------ */
  /* Modèle                                                              */
  /* ------------------------------------------------------------------ */

  function sheetForMonth(book, m) {
    return book.sheets.find((s) => {
      const f = fold(s.name);
      return f === pad2(m) || f === String(m) || (!/^\d+$/.test(f) && monthFromText(s.name) === m);
    }) || null;
  }

  /** Lit le modèle : colonnes, ligne d'en-tête et contenu actuel de chaque onglet mensuel. */
  async function analyzeTemplate(book) {
    const first = sheetForMonth(book, 1) || book.sheets[0];
    const rows = await X.readRows(book, first, { maxRows: 20 });
    const h = T.findHeader(rows);
    if (!h) throw new Error('Aucune colonne « EAN » trouvée dans le modèle.');
    const tplCols = [];
    for (let c = 1; c < h.headers.length; c++) {
      const v = h.headers[c];
      if (v == null || String(v).trim() === '') continue;
      tplCols.push({ col: c, header: String(v), role: templateRole(v) });
    }
    const months = [];
    for (let m = 1; m <= 12; m++) {
      const sheet = sheetForMonth(book, m);
      months.push({ month: m, sheet, rows: sheet ? await X.countDataRows(book, sheet, h.headerRow) : 0 });
    }
    return { headerRow: h.headerRow, tplCols, months };
  }

  /** Écrit les lignes de chaque mois dans son onglet et renvoie le nouveau classeur. */
  async function fillTemplate(book, tplInfo, byMonth, onProgress) {
    const parts = new Map();
    const report = [];
    const months = [...byMonth.keys()].sort((a, b) => a - b);
    for (let i = 0; i < months.length; i++) {
      const m = months[i];
      const sheet = sheetForMonth(book, m);
      if (!sheet) throw new Error("Le modèle n'a pas d'onglet pour le mois " + pad2(m) + '.');
      if (onProgress) onProgress(i / months.length, sheet.name);
      const res = await X.fillSheetXml(book, sheet, tplInfo.headerRow, byMonth.get(m));
      parts.set(sheet.path, res.xml);
      report.push({ month: m, sheet: sheet.name, rows: byMonth.get(m).length, replaced: res.replaced });
      await X.tick();
    }
    if (onProgress) onProgress(1, '');
    const blob = await X.writeWorkbook(book, parts);
    return { blob, report };
  }

  function outputFileName(enseigne, year) {
    const e = String(enseigne || '').trim().replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() || 'enseigne';
    return e + '_' + (year || new Date().getFullYear()) + '.xlsx';
  }

  /** Nom d'enseigne et année tirés d'un fichier déjà rempli, comme « LECLERC_2026.xlsx ». */
  function parseOutputName(name) {
    const m = /^(.+?)[_ -](20\d{2})\.xls[xm]$/i.exec(String(name || ''));
    if (!m) return null;
    if (/^nom[_ ]?enseigne/i.test(m[1])) return null;
    return { enseigne: m[1].trim(), year: +m[2] };
  }

  return {
    MONTH_LABELS,
    pad2,
    monthFromText,
    monthFromFileName,
    monthFromValue,
    yearFromText,
    toFlag,
    toNumber,
    templateRole,
    describeColumns,
    analyzeSource,
    analyzeWithHeader,
    categoryCounts,
    suggestMapping,
    suggestMonth,
    buildOutput,
    flagPreview,
    sheetForMonth,
    analyzeTemplate,
    fillTemplate,
    outputFileName,
    parseOutputName,
  };
});
