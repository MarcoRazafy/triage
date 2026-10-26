/*
 * ui.js
 * Petits outils d'interface partagés par les écrans « Trier » et « Remplir le modèle ».
 */
(function () {
  'use strict';

  const T = window.Triage;
  const nf = new Intl.NumberFormat('fr-FR');
  const pct = new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 1 });
  const dec = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 3 });
  const secs = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 });

  const CAT_CLASS = { 'Fromage LS': 'fromage', 'CORPS GRAS': 'corps', 'LAIT': 'lait', 'CREME': 'creme' };
  CAT_CLASS[T.NOT_FOUND] = 'nf';
  CAT_CLASS[T.NO_EAN] = 'ne';
  const catClass = (c) => CAT_CLASS[c] || 'other';

  /** Crée un élément DOM : h('div', {class: 'x'}, 'texte', autreElement). */
  function h(tag, attrs, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') node.className = v;
      else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
      else node.setAttribute(k, v === true ? '' : v);
    }
    for (const c of children.flat()) {
      if (c == null || c === false) continue;
      node.append(c instanceof Node ? c : document.createTextNode(String(c)));
    }
    return node;
  }

  function show(node, visible) { node.hidden = !visible; }

  function formatCell(v) {
    if (v == null) return '';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v) : dec.format(v);
    if (typeof v === 'boolean') return v ? 'VRAI' : 'FAUX';
    return String(v).trim();
  }

  function formatSize(bytes) {
    if (bytes < 1024 * 1024) return secs.format(bytes / 1024) + ' Ko';
    return secs.format(bytes / 1024 / 1024) + ' Mo';
  }

  function checkExtension(name) {
    if (/\.xls$/i.test(name)) return 'Le format .xls n\'est pas pris en charge. Ouvrez « ' + name + ' » dans Excel ou LibreOffice et enregistrez-le en .xlsx.';
    if (!/\.(xlsx|xlsm)$/i.test(name)) return '« ' + name + ' » n\'est pas un fichier Excel .xlsx.';
    return null;
  }

  function setupDrop(zone, input, onFiles) {
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('over'));
    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.classList.remove('over');
      if (e.dataTransfer.files.length) onFiles([...e.dataTransfer.files]);
    });
    zone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    input.addEventListener('change', () => {
      const files = [...input.files];
      input.value = '';
      if (files.length) onFiles(files);
    });
  }

  /** Barre de progression : box contient .bar span et .label. */
  function setProgress(box, fraction, label) {
    show(box, true);
    box.querySelector('.bar span').style.width = Math.round(Math.min(1, Math.max(0, fraction)) * 100) + '%';
    box.querySelector('.label').textContent = label;
  }

  function download(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = h('a', { href: url, download: fileName });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  window.UI = { nf, pct, dec, secs, catClass, h, show, formatCell, formatSize, checkExtension, setupDrop, setProgress, download };
})();
