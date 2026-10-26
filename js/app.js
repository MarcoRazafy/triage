/*
 * app.js
 * Interface : chargement des fichiers, réglages, lancement du triage,
 * aperçu des résultats et téléchargement.
 */
(function () {
  'use strict';

  const X = window.XlsxLite;
  const T = window.Triage;
  const $ = (sel) => document.querySelector(sel);
  const { nf, pct, secs, catClass, h, show, formatCell, formatSize, checkExtension, setupDrop } = window.UI;
  const PREVIEW_ROWS = 60;

  const state = {
    refs: [],
    index: new Map(),
    conflicts: [],
    raw: null,
    result: null,
    busy: false,
    filter: 'all',
  };

  const el = {
    support: $('#support'),
    dropRef: $('#drop-ref'), inRef: $('#in-ref'), refList: $('#ref-list'), refNotice: $('#ref-notice'),
    dropRaw: $('#drop-raw'), inRaw: $('#in-raw'), rawError: $('#raw-error'), rawPanel: $('#raw-panel'),
    rawName: $('#raw-name'), rawMeta: $('#raw-meta'),
    fSheetWrap: $('#f-sheet-wrap'), fSheet: $('#f-sheet'), fHeader: $('#f-header'),
    fEan: $('#f-ean'), fLabel: $('#f-label'), fOut: $('#f-out'), fNf: $('#f-nf'),
    headPreview: $('#head-preview'),
    btnRun: $('#btn-run'), runHint: $('#run-hint'),
    progress: $('#progress'), progressBar: $('#progress .bar span'), progressLabel: $('#progress .label'),
    runError: $('#run-error'), result: $('#result'), summary: $('#summary'), kpis: $('#kpis'),
    btnDl: $('#btn-dl'), dlName: $('#dl-name'), previewBody: $('#preview tbody'), previewNote: $('#preview-note'),
  };

  /* ------------------------------------------------------------------ */
  /* Étape 1 : référentiels                                              */
  /* ------------------------------------------------------------------ */

  async function addRefs(files) {
    for (const file of files) {
      const error = checkExtension(file.name);
      const entry = { name: file.name, status: error ? 'error' : 'loading', error, override: '' };
      const pos = state.refs.findIndex((r) => r.name === file.name);
      if (pos >= 0) state.refs[pos] = entry; else state.refs.push(entry);
      renderRefs();
      if (error) continue;
      try {
        Object.assign(entry, await T.loadReference(file), { status: 'ok' });
      } catch (e) {
        entry.status = 'error';
        entry.error = e.message || String(e);
      }
      renderRefs();
    }
    refsChanged();
  }

  function refsChanged() {
    const ok = state.refs.filter((r) => r.status === 'ok');
    const built = T.buildIndex(ok);
    state.index = built.index;
    state.conflicts = built.conflicts;
    clearResult();
    renderRefs();
    updateRunState();
  }

  function renderRefs() {
    el.refList.replaceChildren(...state.refs.map((ref, i) => {
      let meta;
      let dotClass = 'none';
      let select = null;
      if (ref.status === 'loading') {
        meta = 'Lecture en cours…';
      } else if (ref.status === 'error') {
        meta = ref.error;
      } else {
        const auto = T.autoLabel(ref);
        const cat = ref.override || auto;
        dotClass = catClass(cat);
        const parts = [nf.format(ref.count) + ' EAN', 'détecté : ' + (auto === 'plusieurs' ? 'plusieurs catégories' : auto)];
        if (ref.duplicates) parts.push(nf.format(ref.duplicates) + ' doublon' + (ref.duplicates > 1 ? 's' : '') + ' ignoré' + (ref.duplicates > 1 ? 's' : ''));
        meta = parts.join(' · ');
        select = h('select', {
          'aria-label': 'Catégorie pour ' + ref.name,
          onchange: (e) => { ref.override = e.target.value; refsChanged(); },
        },
        h('option', { value: '' }, auto === 'plusieurs' ? 'Auto : selon segmentation_0' : 'Auto : ' + auto),
        T.CATEGORIES.map((c) => h('option', { value: c, selected: ref.override === c }, 'Forcer : ' + c)));
      }
      return h('li', { class: 'ref ' + ref.status },
        h('span', { class: 'dot ' + dotClass }),
        h('div', { class: 'ref-body' },
          h('div', { class: 'ref-name', title: ref.name }, ref.name),
          h('div', { class: 'ref-meta' }, meta)),
        select,
        h('button', {
          type: 'button', class: 'icon-btn', title: 'Retirer ce référentiel', 'aria-label': 'Retirer ' + ref.name,
          onclick: () => { state.refs.splice(i, 1); refsChanged(); },
        }, '✕'));
    }));

    const ok = state.refs.filter((r) => r.status === 'ok');
    const notes = [];
    if (state.conflicts.length) {
      const ex = state.conflicts[0];
      notes.push(nf.format(state.conflicts.length) + ' EAN figurent dans deux référentiels avec des catégories différentes. ' +
        'Le référentiel placé le plus haut dans la liste l\'emporte. Exemple : ' + ex.ean + ' gardé en ' + ex.kept.cat +
        ' au lieu de ' + ex.dropped.cat + '.');
    }
    el.refNotice.textContent = notes.join(' ');
    show(el.refNotice, ok.length > 0 && notes.length > 0);
  }

  /* ------------------------------------------------------------------ */
  /* Étape 2 : fichier brut                                              */
  /* ------------------------------------------------------------------ */

  function showRawError(msg) {
    el.rawError.textContent = msg;
    show(el.rawError, !!msg);
  }

  async function loadRaw(files) {
    const file = files[0];
    showRawError('');
    clearResult();
    const error = checkExtension(file.name);
    if (error) { showRawError(error); return; }

    if (state.raw) for (const s of state.raw.book.sheets) s._xml = null;
    state.raw = null;
    renderRaw();
    updateRunState();
    el.dropRaw.classList.add('loading');
    el.dropRaw.querySelector('.drop-text').textContent = 'Lecture de « ' + file.name + ' »…';
    try {
      const book = await X.openWorkbook(await file.arrayBuffer());
      const first = book.sheets.findIndex((s) => s.state === 'visible');
      state.raw = { file, book };
      await selectSheet(first >= 0 ? first : 0);
    } catch (e) {
      state.raw = null;
      showRawError('Impossible de lire « ' + file.name + ' » : ' + (e.message || e));
    } finally {
      el.dropRaw.classList.remove('loading');
    }
    renderRaw();
    updateRunState();
  }

  async function selectSheet(i) {
    const raw = state.raw;
    raw.book.sheets.forEach((s, k) => { if (k !== i) s._xml = null; });
    raw.sheetIdx = i;
    raw.sheet = raw.book.sheets[i];
    raw.rows = await X.readRows(raw.book, raw.sheet, { maxRows: PREVIEW_ROWS });
    raw.lastRow = await X.lastRowNumber(raw.book, raw.sheet);
    const found = T.findHeader(raw.rows);
    raw.headerFound = !!found;
    applyHeaderRow(found ? found.headerRow : (raw.rows[0] ? raw.rows[0].r : 1));
  }

  function applyHeaderRow(r) {
    const raw = state.raw;
    raw.headerRow = r;
    const row = raw.rows.find((x) => x.r === r);
    raw.headers = row ? row.cells.slice() : [];
    const found = T.findHeader([{ r: 1, cells: raw.headers }]);
    raw.eanCol = found ? found.eanCol : 0;
    raw.labelCol = T.findLabelColumn(raw.headers, outName());
  }

  function outName() { return el.fOut.value.trim() || 'produit'; }

  function columnCount() {
    const raw = state.raw;
    let n = raw.headers.length - 1;
    for (const row of raw.rows) if (row.r >= raw.headerRow) n = Math.max(n, row.cells.length - 1);
    return Math.max(n, 0);
  }

  function columnOptions(selected, withNone) {
    const raw = state.raw;
    const opts = [];
    if (withNone) opts.push(h('option', { value: '0', selected: !selected }, '(aucune)'));
    else if (!selected) opts.push(h('option', { value: '0', selected: true, disabled: true }, 'Choisir la colonne…'));
    for (let c = 1; c <= columnCount(); c++) {
      const name = formatCell(raw.headers[c]);
      opts.push(h('option', { value: String(c), selected: c === selected }, X.numToCol(c) + ' · ' + (name || '(sans titre)')));
    }
    return opts;
  }

  function renderRaw() {
    const raw = state.raw;
    el.dropRaw.classList.toggle('compact', !!raw);
    el.dropRaw.querySelector('.drop-text').replaceChildren(
      h('strong', null, raw ? 'Changer de fichier brut' : 'Glisser le fichier brut ici'),
      raw ? ' : glisser un autre fichier ou cliquer' : ' ou cliquer pour le choisir');
    show(el.rawPanel, !!raw);
    if (!raw) return;

    el.rawName.textContent = raw.file.name;
    const meta = [formatSize(raw.file.size), 'feuille « ' + raw.sheet.name + ' »'];
    if (raw.lastRow) meta.unshift(nf.format(raw.lastRow) + ' lignes');
    meta.push(raw.headerFound ? 'en-tête trouvé ligne ' + raw.headerRow : 'colonne EAN non trouvée automatiquement');
    el.rawMeta.textContent = meta.join(' · ');

    show(el.fSheetWrap, raw.book.sheets.length > 1);
    el.fSheet.replaceChildren(...raw.book.sheets.map((s, i) =>
      h('option', { value: String(i), selected: i === raw.sheetIdx }, s.name + (s.state !== 'visible' ? ' (masquée)' : ''))));
    el.fHeader.value = String(raw.headerRow);
    el.fEan.replaceChildren(...columnOptions(raw.eanCol, false));
    el.fLabel.replaceChildren(...columnOptions(raw.labelCol, true));
    renderHeadPreview();
  }

  function renderHeadPreview() {
    const raw = state.raw;
    const ncol = columnCount();
    const existing = T.findOutputColumn(raw.headers, outName());
    const outCol = existing || ncol + 1;
    const cols = [];
    for (let c = 1; c <= Math.max(ncol, outCol); c++) cols.push(c);
    const cls = (c) => (c === raw.eanCol ? 'is-ean' : c === outCol ? 'is-out' : null);

    const head = h('tr', null, h('th', { class: 'num' }, ''), cols.map((c) =>
      h('th', { class: cls(c) }, X.numToCol(c))));
    const rows = raw.rows.filter((r) => r.r >= raw.headerRow).slice(0, 5);
    const body = rows.map((row) => {
      const isHeader = row.r === raw.headerRow;
      return h('tr', { class: isHeader ? 'is-header' : null },
        h('td', { class: 'num' }, String(row.r)),
        cols.map((c) => {
          let v = formatCell(row.cells[c]);
          if (c === outCol && !existing) v = isHeader ? outName() : '…';
          else if (c === outCol && !isHeader) v = '…';
          return h('td', { class: cls(c), title: v }, v);
        }));
    });
    if (!rows.length) {
      body.push(h('tr', null, h('td', { colspan: String(cols.length + 1) }, 'Aucune donnée à partir de cette ligne.')));
    }
    el.headPreview.replaceChildren(h('thead', null, head), h('tbody', null, body));
  }

  /* ------------------------------------------------------------------ */
  /* Étape 3 : triage                                                    */
  /* ------------------------------------------------------------------ */

  function updateRunState() {
    const okRefs = state.refs.filter((r) => r.status === 'ok');
    const loading = state.refs.some((r) => r.status === 'loading');
    let hint = '';
    let ready = false;
    if (state.busy) hint = '';
    else if (loading) hint = 'Lecture des référentiels en cours…';
    else if (!okRefs.length) hint = 'Ajoutez au moins un référentiel.';
    else if (!state.raw) hint = 'Chargez un fichier brut.';
    else if (!state.raw.eanCol) hint = 'Choisissez la colonne EAN du fichier brut.';
    else {
      ready = true;
      hint = nf.format(state.index.size) + ' EAN de référence prêts.';
    }
    el.btnRun.disabled = !ready || state.busy;
    el.runHint.textContent = hint;
  }

  function setProgress(fraction, label) {
    show(el.progress, true);
    el.progressBar.style.width = Math.round(Math.min(1, Math.max(0, fraction)) * 100) + '%';
    el.progressLabel.textContent = label;
  }

  function clearResult() {
    state.result = null;
    show(el.result, false);
    show(el.runError, false);
  }

  async function run() {
    if (state.busy || !state.raw) return;
    const raw = state.raw;
    state.busy = true;
    clearResult();
    updateRunState();
    setProgress(0, 'Classement des lignes…');
    const name = outName();
    const cfg = {
      headerRow: raw.headerRow,
      eanCol: raw.eanCol,
      labelCol: raw.labelCol,
      outName: name,
      notFoundText: el.fNf.value.trim(),
      targetCol: T.findOutputColumn(raw.headers, name),
    };
    const t0 = performance.now();
    try {
      const res = await T.runTriage(raw.book, raw.sheet, cfg, state.index, (f, phase) => {
        if (phase === 'zip') setProgress(0.95, 'Création du fichier Excel…');
        else setProgress(f * 0.9, 'Classement des lignes… ' + Math.round(f * 100) + ' %');
      });
      const ext = /\.xlsm$/i.test(raw.file.name) ? '.xlsm' : '.xlsx';
      res.fileName = raw.file.name.replace(/\.(xlsx|xlsm)$/i, '') + ' - trié' + ext;
      res.seconds = (performance.now() - t0) / 1000;
      res.outName = name;
      res.reused = !!cfg.targetCol;
      state.result = res;
      state.filter = 'all';
      setProgress(1, 'Terminé.');
      show(el.progress, false);
      renderResult();
    } catch (e) {
      show(el.progress, false);
      el.runError.textContent = 'Erreur pendant le triage : ' + (e.message || e);
      show(el.runError, true);
    } finally {
      state.busy = false;
      updateRunState();
    }
  }

  function renderResult() {
    const res = state.result;
    const st = res.stats;
    const classified = [...st.byCat.values()].reduce((a, b) => a + b, 0);
    const col = X.numToCol(res.targetCol);

    el.summary.replaceChildren(
      h('strong', null, nf.format(classified) + ' lignes classées'),
      ' sur ' + nf.format(st.rows) + ', en ' + secs.format(res.seconds) + ' s. ',
      res.reused
        ? 'La colonne « ' + res.outName + ' » existante, en ' + col + ', a été mise à jour.'
        : 'La colonne « ' + res.outName + ' » a été ajoutée en ' + col + '.');

    const cats = [...T.CATEGORIES];
    for (const c of st.byCat.keys()) if (!cats.includes(c)) cats.push(c);
    const tiles = [{ key: 'all', label: 'Toutes les lignes', n: st.rows, dot: null }];
    for (const c of cats) tiles.push({ key: c, label: c, n: st.byCat.get(c) || 0, dot: catClass(c) });
    tiles.push({ key: T.NOT_FOUND, label: 'Non trouvé', n: st.notFound, dot: 'nf' });
    if (st.noEan) tiles.push({ key: T.NO_EAN, label: 'Sans EAN', n: st.noEan, dot: 'ne' });

    el.kpis.replaceChildren(...tiles.map((t) => h('button', {
      type: 'button',
      class: 'kpi' + (t.n === 0 ? ' zero' : ''),
      'aria-pressed': String(state.filter === t.key),
      onclick: () => { state.filter = t.key; renderResult(); },
    },
    h('span', { class: 'k-label' }, t.dot ? h('span', { class: 'dot ' + t.dot }) : null, t.label),
    h('span', { class: 'k-value' }, nf.format(t.n)),
    h('span', { class: 'k-pct' }, t.key === 'all' ? 'lignes de données' : st.rows ? pct.format(t.n / st.rows) : '—'))));

    el.dlName.textContent = res.fileName + ' · ' + formatSize(res.blob.size);
    renderPreview();
    show(el.result, true);
  }

  function renderPreview() {
    const res = state.result;
    let rows;
    if (state.filter === 'all') {
      rows = [];
      for (const list of res.samples.values()) rows.push(...list);
      rows.sort((a, b) => a.r - b.r);
      rows = rows.slice(0, 300);
    } else {
      rows = res.samples.get(state.filter) || [];
    }
    el.previewBody.replaceChildren(...rows.map((it) => h('tr', null,
      h('td', { class: 'num' }, String(it.r)),
      h('td', { class: 'mono' }, it.ean || '—'),
      h('td', { class: 'label' }, it.label),
      h('td', null, h('span', { class: 'chip' }, h('span', { class: 'dot ' + catClass(it.cat) }), it.cat)))));
    if (!rows.length) {
      el.previewBody.append(h('tr', null, h('td', { colspan: '4', class: 'muted' }, 'Aucune ligne dans cette catégorie.')));
    }
    el.previewNote.textContent = 'Aperçu limité à 300 lignes par catégorie. Le fichier téléchargé contient toutes les lignes.';
  }

  function download() {
    const res = state.result;
    if (!res) return;
    const url = URL.createObjectURL(res.blob);
    const a = h('a', { href: url, download: res.fileName });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  /* ------------------------------------------------------------------ */
  /* Branchements                                                        */
  /* ------------------------------------------------------------------ */

  function init() {
    let supported = true;
    try {
      new DecompressionStream('deflate-raw');
      new CompressionStream('deflate-raw');
    } catch (e) {
      supported = false;
    }
    show(el.support, !supported);

    // Un fichier lâché à côté d'une zone ne doit pas être ouvert par le navigateur.
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => e.preventDefault());

    setupDrop(el.dropRef, el.inRef, addRefs);
    setupDrop(el.dropRaw, el.inRaw, loadRaw);

    el.fSheet.addEventListener('change', async () => {
      clearResult();
      await selectSheet(+el.fSheet.value);
      renderRaw();
      updateRunState();
    });
    el.fHeader.addEventListener('change', () => {
      const r = Math.min(PREVIEW_ROWS, Math.max(1, parseInt(el.fHeader.value, 10) || 1));
      clearResult();
      applyHeaderRow(r);
      renderRaw();
      updateRunState();
    });
    el.fEan.addEventListener('change', () => {
      state.raw.eanCol = +el.fEan.value;
      clearResult();
      renderHeadPreview();
      updateRunState();
    });
    el.fLabel.addEventListener('change', () => {
      state.raw.labelCol = +el.fLabel.value;
      clearResult();
    });
    el.fOut.addEventListener('input', () => {
      clearResult();
      if (state.raw) renderHeadPreview();
    });
    el.fNf.addEventListener('input', clearResult);
    el.btnRun.addEventListener('click', run);
    el.btnDl.addEventListener('click', download);
    document.querySelector('#btn-to-modele').addEventListener('click', () => {
      const res = state.result;
      if (!res || !window.ModeleApp) return;
      window.ModeleApp.useSourceFile(new File([res.blob], res.fileName));
      location.hash = '#modele';
    });

    renderRefs();
    renderRaw();
    updateRunState();
  }

  init();
})();
