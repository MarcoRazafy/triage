/*
 * app-modele.js
 * Écran « Remplir le modèle » : fichier trié, modèle, correspondance des colonnes,
 * remplissage des onglets mensuels et téléchargement.
 */
(function () {
  'use strict';

  const X = window.XlsxLite;
  const T = window.Triage;
  const M = window.Modele;
  const { nf, catClass, h, show, formatCell, formatSize, checkExtension, setupDrop, setProgress, download } = window.UI;
  const $ = (sel) => document.querySelector(sel);
  const MEMO_KEY = 'triage.modele.v2';
  const PREVIEW_MAX = 15;

  const st = {
    src: null,      // { file, book, sheetIdx, sheet, rows, merges, info, counts, none, dataRows }
    tpl: null,      // { file, book, info }
    mapping: null,  // n° de colonne du modèle → règle
    month: null,    // règle du mois
    keep: null,     // catégories à garder
    monthsFound: null,
    monthsBad: 0,
    memoApplied: false,
    busy: false,
    result: null,
  };

  const el = {
    dropSrc: $('#m-drop-src'), inSrc: $('#m-in-src'), srcError: $('#m-src-error'), srcPanel: $('#m-src-panel'),
    srcName: $('#m-src-name'), srcMeta: $('#m-src-meta'), fSheetWrap: $('#m-f-sheet-wrap'), fSheet: $('#m-f-sheet'),
    fHeader: $('#m-f-header'), fCat: $('#m-f-cat'), cats: $('#m-cats'), catsNote: $('#m-cats-note'),
    dropTpl: $('#m-drop-tpl'), inTpl: $('#m-in-tpl'), tplError: $('#m-tpl-error'), tplPanel: $('#m-tpl-panel'),
    tplName: $('#m-tpl-name'), tplMeta: $('#m-tpl-meta'), months: $('#m-months'),
    fEnseigne: $('#m-f-enseigne'), fYear: $('#m-f-year'), outName: $('#m-outname'),
    mapEmpty: $('#m-map-empty'), mapPanel: $('#m-map-panel'), fMonth: $('#m-f-month'), monthNote: $('#m-month-note'),
    mapBody: $('#m-map tbody'), flagNote: $('#m-flag-note'), fMerge: $('#m-f-merge'), memoNote: $('#m-memo-note'),
    btnRun: $('#m-btn-run'), runHint: $('#m-run-hint'), progress: $('#m-progress'), runError: $('#m-run-error'),
    result: $('#m-result'), summary: $('#m-summary'), report: $('#m-report tbody'),
    btnDl: $('#m-btn-dl'), btnNext: $('#m-btn-next'), dlName: $('#m-dl-name'),
    previewTitle: $('#m-preview-title'), preview: $('#m-preview'),
  };

  /* ------------------------------------------------------------------ */
  /* Mémoire des réglages, par format de fichier                         */
  /* ------------------------------------------------------------------ */

  function memoAll() {
    try { return JSON.parse(localStorage.getItem(MEMO_KEY) || '{}') || {}; } catch (e) { return {}; }
  }
  function memoGet(sig) { return memoAll()[sig] || null; }
  function memoSet(sig, data) {
    try {
      const all = memoAll();
      all[sig] = data;
      localStorage.setItem(MEMO_KEY, JSON.stringify(all));
    } catch (e) { /* stockage indisponible : on continue sans mémoire */ }
  }

  function saveMemo() {
    const mapping = {};
    for (const t of st.tpl.info.tplCols) mapping[T.fold(t.header)] = st.mapping[t.col];
    let month = null;
    if (st.month.type !== 'fixed') month = st.month;
    else if (st.month.fromFile) month = 'file';
    memoSet(st.src.info.signature, {
      mapping,
      month,
      merge: el.fMerge.checked,
      keep: st.keep ? [...st.keep] : null,
      enseigne: el.fEnseigne.value.trim(),
      saved: new Date().toISOString(),
    });
  }

  /* ------------------------------------------------------------------ */
  /* Règles et valeurs des listes déroulantes                            */
  /* ------------------------------------------------------------------ */

  /** Valeur de la liste pour une règle : toujours une colonne du fichier, ou vide. */
  function specToValue(sp) {
    const info = st.src.info;
    if (!sp) return '';
    if (sp.type === 'col' || sp.type === 'flag') return 'col:' + sp.col;
    if (sp.type === 'monthly') return info.wide ? 'col:' + info.wide.groups[0].cols[sp.offset] : '';
    if (sp.type === 'category') return info.catCol ? 'col:' + info.catCol : '';
    return '';
  }

  /**
   * Règle correspondant à une colonne choisie. La traduction en O ou N pour la promo
   * et la lecture mois par mois pour les fichiers à un mois par colonne sont automatiques.
   */
  function valueToSpec(value, t) {
    if (!value) return { type: 'empty' };
    const col = +String(value).split(':')[1];
    const info = st.src.info;
    if (t.role === 'promo') return { type: 'flag', col };
    if (st.month && st.month.type === 'wide' && info.wide) {
      for (const g of info.wide.groups) {
        const i = g.cols.indexOf(col);
        if (i >= 0) return { type: 'monthly', offset: i };
      }
    }
    return { type: 'col', col };
  }

  /** Remet chaque règle sous sa forme « colonne », par exemple après un changement de mois. */
  function normalizeMapping() {
    for (const t of st.tpl.info.tplCols) st.mapping[t.col] = valueToSpec(specToValue(st.mapping[t.col]), t);
  }

  function monthKey(ms) {
    if (!ms) return '';
    if (ms.type === 'col') return 'col:' + ms.col;
    if (ms.type === 'wide') return 'wide';
    if (ms.fromFile) return 'file';
    return ms.month ? 'fixed:' + ms.month : '';
  }

  function parseMonthKey(k) {
    if (k === 'wide') return { type: 'wide' };
    if (k === 'file') return { type: 'fixed', month: st.src.info.fileMonth, fromFile: true };
    const [type, n] = k.split(':');
    if (type === 'col') return { type: 'col', col: +n };
    return { type: 'fixed', month: +n || 0 };
  }

  function specValid(sp, info) {
    if (!sp || !sp.type) return false;
    if (sp.type === 'col' || sp.type === 'flag') return sp.col >= 1 && sp.col <= info.cols.length;
    if (sp.type === 'monthly') return !!info.wide && sp.offset < info.wide.offsets.length;
    if (sp.type === 'category') return !!info.catCol;
    return sp.type === 'empty' || sp.type === 'ifpromo';
  }

  function monthValid(ms, info) {
    if (!ms) return false;
    if (ms.type === 'col') return ms.col >= 1 && ms.col <= info.cols.length;
    if (ms.type === 'wide') return !!info.wide;
    return ms.type === 'fixed' && ms.month >= 1 && ms.month <= 12;
  }

  /** Libellé d'une colonne : lettre, titre, et sur-titre quand il est utile pour la distinguer. */
  function colLabel(c) {
    const info = st.src.info;
    const dup = info.cols.some((o) => o !== c && o.fh && o.fh === c.fh);
    const sup = c.sup && (dup || M.monthFromText(c.sup)) ? ' · ' + c.sup : '';
    return c.letter + ' · ' + (c.header || '(sans titre)') + sup;
  }

  function columnOptions(current) {
    return [
      h('option', { value: '', selected: current === '' }, '(aucune)'),
      ...st.src.info.cols.map((c) => h('option', { value: 'col:' + c.col, selected: current === 'col:' + c.col }, colLabel(c))),
    ];
  }

  /* ------------------------------------------------------------------ */
  /* Chargement des fichiers                                             */
  /* ------------------------------------------------------------------ */

  function clearResult() {
    st.result = null;
    show(el.result, false);
    show(el.runError, false);
  }

  function showError(box, msg) {
    box.textContent = msg || '';
    show(box, !!msg);
  }

  async function loadSource(files) {
    const file = files[0];
    showError(el.srcError, '');
    clearResult();
    const err = checkExtension(file.name);
    if (err) { showError(el.srcError, err); return; }
    st.src = null;
    st.mapping = null;
    renderAll();
    const label = el.dropSrc.querySelector('.drop-text');
    el.dropSrc.classList.add('loading');
    label.textContent = 'Lecture de « ' + file.name + ' »…';
    try {
      const book = await X.openWorkbook(await file.arrayBuffer());
      const first = book.sheets.findIndex((s) => s.state === 'visible');
      st.src = { file, book };
      await selectSourceSheet(first >= 0 ? first : 0, (f) => {
        label.textContent = 'Lecture de « ' + file.name + ' »… ' + Math.round(f * 100) + ' %';
      });
    } catch (e) {
      st.src = null;
      showError(el.srcError, 'Impossible de lire « ' + file.name + ' » : ' + (e.message || e));
    } finally {
      el.dropSrc.classList.remove('loading');
    }
    renderAll();
  }

  async function selectSourceSheet(i, onProgress) {
    const s = st.src;
    s.sheetIdx = i;
    s.sheet = s.book.sheets[i];
    s.rows = await X.readRows(s.book, s.sheet, { onProgress });
    s.merges = await X.readMerges(s.book, s.sheet);
    s.sheet._xml = null;
    s.info = M.analyzeSource(s.file.name, s.rows, s.merges, 'produit');
    afterSourceAnalysis(true);
  }

  function afterSourceAnalysis(resetKeep) {
    const s = st.src;
    const cc = M.categoryCounts(s.rows, s.info.headerRow, s.info.catCol);
    s.counts = cc.counts;
    s.none = cc.none;
    s.dataRows = s.rows.reduce((n, r) => n + (r.r > s.info.headerRow && r.cells.length ? 1 : 0), 0);
    if (resetKeep || !st.keep) {
      const cats = [...cc.counts.keys()];
      const memo = memoGet(s.info.signature);
      const fromMemo = memo && memo.keep ? memo.keep.filter((c) => cc.counts.has(c)) : [];
      const lactalis = cats.filter((c) => T.CATEGORIES.includes(c));
      st.keep = new Set(fromMemo.length ? fromMemo : lactalis.length ? lactalis : cats);
    }
    if (!s.info.catCol) st.keep = null;
    const memo = memoGet(s.info.signature);
    if (memo && memo.enseigne && !el.fEnseigne.value.trim()) el.fEnseigne.value = memo.enseigne;
    if (!+el.fYear.value && s.info.fileYear) el.fYear.value = String(s.info.fileYear);
    st.mapping = null;
    ensureMapping();
  }

  async function loadTemplate(files) {
    const file = files[0];
    showError(el.tplError, '');
    clearResult();
    const err = checkExtension(file.name);
    if (err) { showError(el.tplError, err); return; }
    if (st.tpl) for (const s of st.tpl.book.sheets) s._xml = null;
    st.tpl = null;
    st.mapping = null;
    renderAll();
    el.dropTpl.classList.add('loading');
    el.dropTpl.querySelector('.drop-text').textContent = 'Lecture de « ' + file.name + ' »…';
    try {
      const book = await X.openWorkbook(await file.arrayBuffer());
      const info = await M.analyzeTemplate(book);
      if (!info.months.some((m) => m.sheet)) throw new Error('aucun onglet 01 à 12 trouvé.');
      st.tpl = { file, book, info };
      const parsed = M.parseOutputName(file.name);
      if (parsed) {
        el.fEnseigne.value = parsed.enseigne;
        el.fYear.value = String(parsed.year);
      }
      if (!+el.fYear.value) el.fYear.value = String((st.src && st.src.info.fileYear) || new Date().getFullYear());
      ensureMapping();
    } catch (e) {
      st.tpl = null;
      showError(el.tplError, 'Impossible d\'utiliser « ' + file.name + ' » comme modèle : ' + (e.message || e));
    } finally {
      el.dropTpl.classList.remove('loading');
    }
    renderAll();
  }

  /** Correspondance proposée, complétée par les choix mémorisés pour ce format. */
  function ensureMapping() {
    if (!st.src || !st.tpl || st.mapping) return;
    const info = st.src.info;
    const tplCols = st.tpl.info.tplCols;
    const mapping = M.suggestMapping(tplCols, info);
    let month = M.suggestMonth(info);
    let merge = true;
    st.memoApplied = false;
    const memo = memoGet(info.signature);
    if (memo) {
      for (const t of tplCols) {
        const sp = memo.mapping && memo.mapping[T.fold(t.header)];
        if (specValid(sp, info)) mapping[t.col] = Object.assign({}, sp);
      }
      const montantT = tplCols.find((t) => t.role === 'montant');
      for (const t of tplCols) {
        if (mapping[t.col] && mapping[t.col].type === 'ifpromo') {
          mapping[t.col] = montantT ? Object.assign({}, mapping[montantT.col]) : { type: 'empty' };
        }
      }
      const ms = memo.month === 'file'
        ? (info.fileMonth ? { type: 'fixed', month: info.fileMonth, fromFile: true } : null)
        : memo.month;
      if (monthValid(ms, info)) month = ms;
      if (typeof memo.merge === 'boolean') merge = memo.merge;
      st.memoApplied = true;
    }
    st.mapping = mapping;
    st.month = month;
    el.fMerge.checked = merge;
    normalizeMapping();
  }

  /* ------------------------------------------------------------------ */
  /* Calculs pour l'affichage                                            */
  /* ------------------------------------------------------------------ */

  function buildCfg(limit) {
    return {
      tplCols: st.tpl.info.tplCols,
      mapping: st.mapping,
      month: st.month,
      keep: st.src.info.catCol ? st.keep : null,
      merge: el.fMerge.checked,
      limit: limit || 0,
    };
  }

  function computeMonthsFound() {
    const info = st.src.info;
    const ms = st.month;
    const found = new Map();
    let bad = 0;
    if (ms && ms.type === 'wide' && info.wide) {
      for (const g of info.wide.groups) found.set(g.month, 1);
    } else if (ms && ms.type === 'fixed') {
      if (ms.month) found.set(ms.month, 1);
    } else if (ms && ms.type === 'col') {
      for (const row of st.src.rows) {
        if (row.r <= info.headerRow || !row.cells.length) continue;
        if (info.catCol && st.keep) {
          const c = row.cells[info.catCol];
          if (!st.keep.has(c == null ? '' : String(c).trim())) continue;
        }
        const m = M.monthFromValue(row.cells[ms.col]);
        if (m) found.set(m, (found.get(m) || 0) + 1);
        else bad++;
      }
    }
    st.monthsFound = found;
    st.monthsBad = bad;
  }

  function previewRow() {
    const { byMonth } = M.buildOutput(st.src.info, st.src.rows, buildCfg(40));
    const first = [...byMonth.keys()].sort((a, b) => a - b)[0];
    return first ? byMonth.get(first)[0] : null;
  }

  function exampleText(row, t) {
    if (!row) return '—';
    const v = row[t.col];
    return v == null || v === '' ? '—' : formatCell(v);
  }

  function updateExamples() {
    const row = previewRow();
    for (const t of st.tpl.info.tplCols) {
      const td = el.mapBody.querySelector('td[data-col="' + t.col + '"]');
      if (td) td.textContent = exampleText(row, t);
    }
  }

  /* ------------------------------------------------------------------ */
  /* Affichage                                                           */
  /* ------------------------------------------------------------------ */

  function renderAll() {
    renderSource();
    renderMapping();
    renderTemplate();
    updateRun();
  }

  function setDropText(zone, loaded, emptyText, changeText) {
    zone.classList.toggle('compact', loaded);
    zone.querySelector('.drop-text').replaceChildren(
      h('strong', null, loaded ? changeText : emptyText),
      loaded ? ' : glisser un autre fichier ou cliquer' : ' ou cliquer pour le choisir');
  }

  function renderSource() {
    const s = st.src;
    setDropText(el.dropSrc, !!s, 'Glisser le fichier trié ici', 'Changer de fichier trié');
    show(el.srcPanel, !!s);
    if (!s) return;
    el.srcName.textContent = s.file.name;
    el.srcMeta.textContent = [
      nf.format(s.dataRows) + ' lignes de données',
      formatSize(s.file.size),
      'feuille « ' + s.sheet.name + ' »',
      'en-tête ligne ' + s.info.headerRow,
    ].join(' · ');
    show(el.fSheetWrap, s.book.sheets.length > 1);
    el.fSheet.replaceChildren(...s.book.sheets.map((sh, i) => h('option', { value: String(i), selected: i === s.sheetIdx }, sh.name)));
    el.fHeader.value = String(s.info.headerRow);
    el.fCat.replaceChildren(
      h('option', { value: '0', selected: !s.info.catCol }, '(aucune)'),
      ...s.info.cols.map((c) => h('option', { value: String(c.col), selected: c.col === s.info.catCol }, colLabel(c))));

    const order = (c) => { const i = T.CATEGORIES.indexOf(c); return i < 0 ? 99 : i; };
    const cats = [...s.counts.entries()].sort((a, b) => order(a[0]) - order(b[0]) || b[1] - a[1]);
    el.cats.replaceChildren(...cats.map(([cat, n]) => h('label', { class: 'check-chip' },
      h('input', {
        type: 'checkbox',
        checked: !!(st.keep && st.keep.has(cat)),
        onchange: (e) => {
          if (e.target.checked) st.keep.add(cat); else st.keep.delete(cat);
          clearResult();
          renderMapping();
          renderTemplate();
          updateRun();
        },
      }),
      h('span', { class: 'dot ' + catClass(cat) }),
      h('span', null, cat),
      h('span', { class: 'count' }, nf.format(n)))));
    if (!s.info.catCol) {
      el.catsNote.textContent = 'Pas de colonne catégorie : toutes les lignes seront écrites. Passez d\'abord par « Trier » pour ne garder que les catégories suivies.';
    } else {
      el.catsNote.textContent = s.none ? nf.format(s.none) + ' lignes sans catégorie seront ignorées.' : '';
    }
  }

  function renderTemplate() {
    const t = st.tpl;
    setDropText(el.dropTpl, !!t, 'Glisser le modèle ici', 'Changer de modèle');
    show(el.tplPanel, !!t);
    if (!t) return;
    el.tplName.textContent = t.file.name;
    const present = t.info.months.filter((m) => m.sheet);
    const filled = present.filter((m) => m.rows > 0);
    el.tplMeta.textContent = [
      present.length + ' onglets mensuels',
      filled.length ? filled.length + ' déjà rempli' + (filled.length > 1 ? 's' : '') : 'aucun onglet rempli',
      t.info.tplCols.length + ' colonnes',
    ].join(' · ');
    const target = st.monthsFound || new Map();
    el.months.replaceChildren(...t.info.months.map((m) => {
      let state = 'vide';
      if (!m.sheet) state = 'absent';
      else if (m.rows) state = nf.format(m.rows);
      const cls = ['month'];
      if (m.rows) cls.push('filled');
      if (target.has(m.month)) cls.push('target');
      if (!m.sheet) cls.push('missing');
      const tip = M.MONTH_LABELS[m.month - 1] + ' : ' + (!m.sheet ? 'onglet absent' : m.rows ? nf.format(m.rows) + ' lignes' : 'vide') +
        (target.has(m.month) ? ', sera rempli' : '');
      return h('div', { class: cls.join(' '), title: tip },
        h('span', { class: 'm-num' }, M.pad2(m.month)),
        h('span', { class: 'm-state' }, state));
    }));
    el.outName.textContent = M.outputFileName(el.fEnseigne.value, +el.fYear.value);
  }

  function renderMapping() {
    const ready = !!(st.src && st.tpl && st.mapping);
    show(el.mapEmpty, !ready);
    show(el.mapPanel, ready);
    if (!ready) { st.monthsFound = null; return; }
    const info = st.src.info;

    const cur = monthKey(st.month);
    const detected = [];
    if (info.wide) detected.push(['wide', 'Un mois par colonne : ' + info.wide.groups.map((g) => M.MONTH_LABELS[g.month - 1].toLowerCase()).join(', ')]);
    if (info.fileMonth) detected.push(['file', 'Nom du fichier : ' + M.MONTH_LABELS[info.fileMonth - 1] + ', onglet ' + M.pad2(info.fileMonth)]);
    const opt = ([v, l]) => h('option', { value: v, selected: v === cur }, l);
    el.fMonth.replaceChildren(
      cur === '' ? h('option', { value: '', selected: true, disabled: true }, 'Choisir le mois…') : null,
      detected.length ? h('optgroup', { label: 'Détecté' }, detected.map(opt)) : null,
      h('optgroup', { label: 'Colonne du fichier trié' }, info.cols.map((c) => opt(['col:' + c.col, colLabel(c)]))),
      h('optgroup', { label: 'Mois choisi' }, M.MONTH_LABELS.map((l, i) => opt(['fixed:' + (i + 1), M.pad2(i + 1) + ' · ' + l]))));

    computeMonthsFound();
    const months = [...st.monthsFound.keys()].sort((a, b) => a - b).map(M.pad2);
    let note = months.length
      ? (months.length > 1 ? 'Onglets concernés : ' : 'Onglet concerné : ') + months.join(', ') + '.'
      : 'Aucun mois lisible avec ce réglage.';
    if (st.monthsBad) note += ' Mois illisible sur ' + nf.format(st.monthsBad) + ' lignes : elles seront ignorées.';
    el.monthNote.textContent = note;

    const row = previewRow();
    el.mapBody.replaceChildren(...st.tpl.info.tplCols.map((t) => {
      const select = h('select', {
        'aria-label': 'Colonne pour ' + t.header,
        onchange: (e) => { st.mapping[t.col] = valueToSpec(e.target.value, t); onMappingChanged(); },
      }, columnOptions(specToValue(st.mapping[t.col])));
      return h('tr', null,
        h('td', { class: 'tpl-col' }, t.header),
        h('td', { class: 'src-cell' }, h('div', { class: 'src-wrap' }, select)),
        h('td', { class: 'example', 'data-col': String(t.col) }, exampleText(row, t)));
    }));

    const notes = [];
    const tplCols = st.tpl.info.tplCols;
    const promoT = tplCols.find((t) => t.role === 'promo');
    const psp = promoT && st.mapping[promoT.col];
    if (psp && psp.type === 'flag') {
      const fp = M.flagPreview(st.src.rows, info.headerRow, psp.col, 8);
      notes.push(promoT.header + ' : ' + [...fp].map(([k, v]) => k + ' donne ' + v).join(', ') + '.');
    }
    const montantT = tplCols.find((t) => t.role === 'montant');
    const pvT = tplCols.find((t) => t.role === 'promoVendue');
    if (montantT && pvT && M.sameSource(st.mapping[pvT.col], st.mapping[montantT.col])) {
      notes.push(pvT.header + ' lit la même colonne que ' + montantT.header + ' : le montant est compté seulement sur les lignes en promo.');
    }
    if (st.month && st.month.type === 'wide' && info.wide) {
      notes.push('Un mois par colonne : la colonne choisie pour ' + M.MONTH_LABELS[info.wide.groups[0].month - 1].toLowerCase() + ' est lue au même endroit pour chaque mois.');
    }
    el.flagNote.textContent = notes.join(' ');
    el.memoNote.textContent = st.memoApplied
      ? 'Réglages repris de la dernière fois pour ce format de fichier.'
      : 'Ces réglages seront mémorisés pour ce format de fichier après le remplissage.';
  }

  function onMappingChanged() {
    clearResult();
    renderMapping();
    renderTemplate();
    updateRun();
  }

  function updateRun() {
    let hint = '';
    let ready = false;
    if (st.busy) hint = '';
    else if (!st.src) hint = 'Chargez le fichier trié.';
    else if (!st.tpl) hint = 'Chargez le modèle.';
    else if (st.src.info.catCol && (!st.keep || !st.keep.size)) hint = 'Cochez au moins une catégorie.';
    else if (!monthValid(st.month, st.src.info)) hint = 'Choisissez le mois.';
    else if (!st.monthsFound || !st.monthsFound.size) hint = 'Aucun mois lisible : vérifiez le réglage du mois.';
    else {
      ready = true;
      const ms = [...st.monthsFound.keys()].sort((a, b) => a - b);
      hint = (ms.length > 1 ? 'Onglets ' : 'Onglet ') + ms.map(M.pad2).join(', ') + ' à remplir.';
      const over = st.tpl.info.months.filter((m) => st.monthsFound.has(m.month) && m.rows > 0).map((m) => M.pad2(m.month));
      if (over.length) hint += ' ' + over.join(', ') + (over.length > 1 ? ' contiennent' : ' contient') + ' déjà des lignes : elles seront remplacées.';
      const missing = ms.filter((m) => !st.tpl.info.months[m - 1].sheet);
      if (missing.length) { ready = false; hint = 'Le modèle n\'a pas d\'onglet ' + missing.map(M.pad2).join(', ') + '.'; }
    }
    el.btnRun.disabled = !ready || st.busy;
    el.runHint.textContent = hint;
  }

  function currentOutName() {
    return M.outputFileName(el.fEnseigne.value, +el.fYear.value);
  }

  /* ------------------------------------------------------------------ */
  /* Remplissage                                                         */
  /* ------------------------------------------------------------------ */

  async function run() {
    if (st.busy) return;
    st.busy = true;
    clearResult();
    updateRun();
    setProgress(el.progress, 0.05, 'Préparation des lignes…');
    await X.tick();
    try {
      const out = M.buildOutput(st.src.info, st.src.rows, buildCfg(0));
      if (!out.byMonth.size) throw new Error('aucune ligne à écrire. Vérifiez les catégories cochées et le mois.');
      const res = await M.fillTemplate(st.tpl.book, st.tpl.info, out.byMonth, (f, name) => {
        setProgress(el.progress, 0.1 + f * 0.8, name ? 'Écriture de l\'onglet ' + name + '…' : 'Création du fichier Excel…');
      });
      st.result = { blob: res.blob, report: res.report, stats: out.stats, byMonth: out.byMonth };
      saveMemo();
      show(el.progress, false);
      renderResult();
    } catch (e) {
      show(el.progress, false);
      showError(el.runError, 'Erreur pendant le remplissage : ' + (e.message || e));
    } finally {
      st.busy = false;
      updateRun();
    }
  }

  function renderResult() {
    const r = st.result;
    const total = r.report.reduce((a, x) => a + x.rows, 0);
    el.summary.replaceChildren(
      h('strong', null, nf.format(total) + ' lignes écrites'),
      ' dans ' + r.report.length + (r.report.length > 1 ? ' onglets.' : ' onglet.'),
      r.stats.merged ? ' ' + nf.format(r.stats.merged) + ' lignes ont été additionnées à une ligne du même produit, magasin et mois.' : '',
      r.stats.noMonth ? ' ' + nf.format(r.stats.noMonth) + ' lignes ont été ignorées faute de mois lisible.' : '');
    el.report.replaceChildren(...r.report.map((x) => h('tr', null,
      h('td', null, x.sheet),
      h('td', null, M.MONTH_LABELS[x.month - 1]),
      h('td', { class: 'num' }, nf.format(x.rows)),
      h('td', { class: 'num' }, x.replaced ? nf.format(x.replaced) : '—'))));
    el.dlName.textContent = currentOutName() + ' · ' + formatSize(r.blob.size);

    const first = r.report[0];
    const rows = r.byMonth.get(first.month).slice(0, PREVIEW_MAX);
    el.previewTitle.textContent = 'Aperçu de l\'onglet ' + first.sheet + ' : ' + rows.length + ' premières lignes';
    const cols = st.tpl.info.tplCols;
    el.preview.replaceChildren(
      h('thead', null, h('tr', null, cols.map((t) => h('th', null, t.header)))),
      h('tbody', null, rows.map((row) => h('tr', null, cols.map((t) => {
        const v = row[t.col];
        return h('td', { class: typeof v === 'number' && t.role !== 'ean' ? 'num-val' : null }, formatCell(v));
      })))));
    show(el.result, true);
  }

  /* ------------------------------------------------------------------ */
  /* Branchements                                                        */
  /* ------------------------------------------------------------------ */

  function init() {
    setupDrop(el.dropSrc, el.inSrc, loadSource);
    setupDrop(el.dropTpl, el.inTpl, loadTemplate);

    el.fSheet.addEventListener('change', async () => {
      clearResult();
      await selectSourceSheet(+el.fSheet.value);
      renderAll();
    });
    el.fHeader.addEventListener('change', () => {
      const s = st.src;
      if (!s) return;
      const r = Math.min(60, Math.max(1, parseInt(el.fHeader.value, 10) || 1));
      clearResult();
      s.info = M.analyzeWithHeader(s.file.name, s.rows, s.merges, r, 'produit');
      afterSourceAnalysis(true);
      renderAll();
    });
    el.fCat.addEventListener('change', () => {
      clearResult();
      st.src.info.catCol = +el.fCat.value;
      afterSourceAnalysis(true);
      renderAll();
    });
    el.fMonth.addEventListener('change', () => {
      st.month = parseMonthKey(el.fMonth.value);
      normalizeMapping();
      onMappingChanged();
    });
    el.fMerge.addEventListener('change', () => { clearResult(); updateExamples(); });
    el.fEnseigne.addEventListener('input', () => {
      el.outName.textContent = currentOutName();
      if (st.result) el.dlName.textContent = currentOutName() + ' · ' + formatSize(st.result.blob.size);
    });
    el.fYear.addEventListener('input', () => {
      el.outName.textContent = currentOutName();
      if (st.result) el.dlName.textContent = currentOutName() + ' · ' + formatSize(st.result.blob.size);
    });
    el.btnRun.addEventListener('click', run);
    el.btnDl.addEventListener('click', () => { if (st.result) download(st.result.blob, currentOutName()); });
    el.btnNext.addEventListener('click', async () => {
      if (!st.result) return;
      const file = new File([st.result.blob], currentOutName());
      await loadTemplate([file]);
      el.dropSrc.scrollIntoView({ behavior: 'smooth', block: 'center' });
      el.dropSrc.focus({ preventScroll: true });
    });

    renderAll();
  }

  init();

  window.ModeleApp = {
    useSourceFile(file) { loadSource([file]); },
  };
})();
