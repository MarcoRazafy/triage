/*
 * nav.js
 * Navigation entre l'accueil et les deux écrans, par l'adresse (#trier, #modele).
 * Le bouton Retour du navigateur fonctionne donc normalement.
 */
(function () {
  'use strict';
  const VIEWS = {
    accueil: { id: 'view-home', title: '' },
    trier: { id: 'view-trier', title: 'Trier' },
    modele: { id: 'view-modele', title: 'Remplir le modèle' },
  };
  const crumbs = document.getElementById('crumbs');
  const crumbTitle = document.getElementById('crumb-title');

  function route() {
    const key = (location.hash || '#accueil').slice(1);
    const current = VIEWS[key] ? key : 'accueil';
    for (const [name, v] of Object.entries(VIEWS)) document.getElementById(v.id).hidden = name !== current;
    crumbs.hidden = current === 'accueil';
    crumbTitle.textContent = VIEWS[current].title;
    document.title = VIEWS[current].title ? VIEWS[current].title + ' · Triage des produits' : 'Triage des produits';
    window.scrollTo(0, 0);
  }

  window.addEventListener('hashchange', route);
  route();
})();
