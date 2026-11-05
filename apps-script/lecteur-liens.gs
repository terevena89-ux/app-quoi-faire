/**
 * Quoi faire ? — « lecteur de liens ».
 * Petit Apps Script SÉPARÉ de l'ancienne appli (qui reste intacte en archive).
 * Il ne fait que lire des pages web à la demande de l'appli, car un navigateur
 * n'a pas le droit d'aller lire librement d'autres sites :
 *   - event   : page d'un événement -> titre, dates, lieu, prix, image
 *   - texte   : texte collé d'une page -> dates, lieu, prix
 *   - place   : lien Google Maps -> nom, adresse, position
 *   - geocode : adresse -> position sur la carte
 * Il ne voit jamais tes données (elles sont dans Firebase).
 *
 * Installation : voir les étapes dans le chat (Déployer > Application Web,
 * « Exécuter en tant que : moi », « Accès : Tout le monde »).
 */

/* ---------- Porte d'entrée : l'appli envoie { action, ... } ---------- */

function doPost(e) {
  let r;
  try {
    const q = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let data;
    if (q.action === 'event') data = extractEvent(q.url);
    else if (q.action === 'texte') data = parseEventText(q.texte, q.type, q.titre);
    else if (q.action === 'place') data = extractPlace(q.url);
    else if (q.action === 'geocode') data = geocode_(q.adresse);
    else throw new Error('Action inconnue');
    r = { ok: true, data: data };
  } catch (err) {
    r = { ok: false, erreur: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

// Ouvrir l'adresse /exec dans un navigateur affiche ce message : pratique pour vérifier que ça marche.
function doGet() {
  return ContentService.createTextOutput('Lecteur de liens de « Quoi faire ? » : en marche.');
}

// À lancer une fois depuis l'éditeur : Google demande l'autorisation de lire des pages web
// et d'utiliser le géocodage (adresse -> position sur la carte).
function autoriser() {
  Maps.newGeocoder().setRegion('fr').geocode('Paris');
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  return 'OK';
}

/* ---------- Extraction d'infos depuis un lien ---------- */

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36';

function extractEvent(url) {
  url = String(url || '').trim();
  if (!/^https?:\/\/\S+$/i.test(url)) throw new Error('Ce n’est pas un lien valide.');
  let html;
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': UA, 'Accept-Language': 'fr-FR,fr;q=0.9,en;q=0.5' }
    });
    if (res.getResponseCode() >= 400) {
      return fromUrlOnly_(url, 'Le site bloque la lecture automatique (code ' + res.getResponseCode() + ') : colle le texte de la page ci-dessous pour compléter.');
    }
    html = res.getContentText();
  } catch (err) {
    return fromUrlOnly_(url, 'Page injoignable : colle le texte de la page ci-dessous pour compléter.');
  }
  return parseEventHtml_(html, url, todayIsoFromDate_(new Date()));
}

function emptyEvent_(url, info) {
  return { lien: url, titre: '', type: '', lieu: '', zone: '', debut: '', fin: '', prix: '', image: '', desc: '', infos: [info] };
}

// Site qui bloque la lecture (ex. Shotgun, protégé contre les robots) : on devine ce qu'on peut depuis
// l'adresse de la page (« …/events/tonic-walter-central-chapelle » -> « Tonic Walter Central Chapelle »).
function fromUrlOnly_(url, info) {
  const out = emptyEvent_(url, info);
  out.bloque = true;
  const morceaux = url.replace(/[?#].*$/, '').replace(/\/+$/, '').split('/');
  let slug = morceaux.length > 3 ? morceaux[morceaux.length - 1] : '';
  try { slug = decodeURIComponent(slug); } catch (e) { /* on garde tel quel */ }
  if (slug && !/^\d+$/.test(slug)) {
    out.titre = slug.replace(/\.[a-z]+$/i, '').replace(/[-_]+/g, ' ').replace(/\b\d{5,}\b/g, '').replace(/\s+/g, ' ').trim()
      .replace(/(^|\s)\S/g, function (c) { return c.toUpperCase(); });
  }
  out.type = /shotgun\.live|dice\.fm|residentadvisor|ra\.co\//i.test(url) ? 'Concert / Festival' : (guessType_(url) || 'Autre');
  return out;
}

/* Texte copié depuis la page (quand le site bloque la lecture) : date, lieu et prix. */
// titre : celui de la fiche (souvent deviné depuis l'adresse) ; une ligne du texte qui en reprend tous
// les mots le remplace, car elle est plus complète.
function parseEventText(texte, type, titre) {
  const brut = String(texte || '').slice(0, 100000);
  const today = todayIsoFromDate_(new Date());
  const out = { titre: titreDuTexte_(brut, titre), debut: '', fin: '', lieu: '', zone: '', prix: '' };
  const d = datesDuTexte_(brut, today) || datesFromText_(brut, today, type);
  if (d) { out.debut = d.debut || ''; out.fin = d.fin || ''; }
  if (type === 'Ciné' && out.debut && !out.fin) out.fin = plusUnMois_(out.debut);
  out.lieu = lieuDuTexte_(brut);
  out.zone = guessZone_(out.lieu) || guessZone_(brut, true);
  out.prix = prixDuTexte_(brut);
  return out;
}

function titreDuTexte_(brut, titre) {
  const norme = function (s) { return sansAccents_(String(s).toLowerCase()); };
  const mots = norme(titre || '').split(/[^a-z0-9]+/).filter(function (w) { return w.length > 1 && !/^\d+$/.test(w); });
  if (mots.length < 2) return '';
  const ligne = brut.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(function (s) { return s.length > 3 && s.length < 150; })
    .filter(function (l) { const n = norme(l); return mots.every(function (w) { return new RegExp('\\b' + w + '\\b').test(n); }); })[0];
  if (!ligne || ligne === String(titre).trim()) return '';
  if (ligne !== ligne.toUpperCase()) return ligne;
  // ligne en capitales : on reprend la casse du titre actuel, chiffres romains en capitales
  const casse = {};
  String(titre).split(/\s+/).forEach(function (w) { casse[norme(w)] = w; });
  return ligne.split(' ').map(function (w, i) {
    if (casse[norme(w)]) return casse[norme(w)];
    if (/^[IVXL]+$/.test(w)) return w;
    const bas = w.toLowerCase();
    return i === 0 ? bas.charAt(0).toUpperCase() + bas.slice(1) : bas;
  }).join(' ');
}

// La première date qui apparaît dans le texte (le bas de page parle souvent d'autres événements).
// Une période « 12 oct. - 3 janv. » l'emporte sur sa propre date de début.
function datesDuTexte_(brut, today) {
  const t = sansAccents_(brut.toLowerCase()).replace(/\s+/g, ' ');
  const anRef = +today.slice(0, 4);
  const J = '(\\d{1,2})(?:er)?';
  const M = '(janvier|janv|fevrier|fevr|fev|mars|avril|avr|mai|juin|juillet|juil|aout|septembre|sept|octobre|oct|novembre|nov|decembre|dec)\\b\\.?';
  const Y = '(?:\\s+(\\d{4}))?';
  const cands = [];
  let m;
  const reP = new RegExp('\\b' + J + '(?:\\s+' + M + ')?' + Y + '\\s*(?:au|-|–|—)\\s*' + J + '\\s+' + M + Y, 'g');
  while ((m = reP.exec(t))) {
    const fin = fin_(+m[4], MOIS[m[5]], m[6], today, anRef);
    const debut = debut_(+m[1], MOIS[m[2] || m[5]], m[3], fin);
    if (fin && fin >= today && debut && debut <= fin) cands.push({ i: m.index, j: m.index + m[0].length, debut: debut, fin: fin });
  }
  const reS = new RegExp('\\b' + J + '\\s+' + M + Y, 'g');
  while ((m = reS.exec(t))) {
    const i = m.index;
    if (cands.some(function (c) { return c.i <= i && i < c.j; })) continue;
    const d = fin_(+m[1], MOIS[m[2]], m[3], today, anRef);
    if (d && d >= today) cands.push({ i: i, debut: d, fin: d });
  }
  cands.sort(function (a, b) { return a.i - b.i; });
  return cands[0] || null;
}

// La ligne avec un code postal d'Île-de-France, précédée du nom du lieu s'il est sur la ligne d'avant.
function lieuDuTexte_(brut) {
  const cp = /\b(75|77|78|91|92|93|94|95)\d{3}\b/;
  const lignes = brut.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
  for (let i = 0; i < lignes.length; i++) {
    if (!cp.test(lignes[i]) || lignes[i].length > 140) continue;
    const avant = lignes[i - 1] || '';
    return (avant && avant.length < 50 && !/\d|€/.test(avant) ? avant + ' – ' : '') + lignes[i];
  }
  const m = brut.replace(/\s+/g, ' ').match(/[^.;|]{0,60}\b(75|77|78|91|92|93|94|95)\d{3}\s+[A-Za-zÀ-ÿ'’ -]{2,30}/);
  return m ? m[0].trim() : '';
}

// Le plus petit prix en euros du texte (tarif le moins cher), sinon « Gratuit » si la page le dit.
function prixDuTexte_(brut) {
  const re = /(\d{1,3}(?:[.,]\d{1,2})?)\s?(?:€|eur\b|euros?\b)|€\s?(\d{1,3}(?:[.,]\d{1,2})?)/gi;
  const nums = [];
  let m;
  while ((m = re.exec(brut))) {
    const n = parseFloat((m[1] || m[2]).replace(',', '.'));
    if (n > 0 && n < 1000) nums.push(n);
  }
  if (nums.length) {
    const n = Math.min.apply(null, nums);
    return (n % 1 ? n.toFixed(2).replace('.', ',') : String(n)) + ' €';
  }
  return /gratuit|entree libre/i.test(sansAccents_(brut)) ? 'Gratuit' : '';
}

function parseEventHtml_(html, url, today) {
  const out = emptyEvent_(url, '');
  out.infos = [];

  const ld = findLdEvent_(html);
  if (ld) {
    fromLd_(ld, out);
    out.infos.push('Infos lues dans les données structurées de la page.');
  }

  const meta = metas_(html);
  if (!out.titre) out.titre = meta['og:title'] || meta['twitter:title'] || titleTag_(html);
  if (!out.desc) out.desc = meta['og:description'] || meta['description'] || '';
  if (!out.image) out.image = meta['og:image'] || meta['twitter:image'] || '';
  if (out.image) out.image = absUrl_(out.image, url);

  const texte = pageText_(html);
  if (!out.type) out.type = guessType_(url + ' ' + out.titre + ' ' + (meta['og:type'] || '')) || 'Autre';
  if (!out.debut && !out.fin) {
    const d = datesFromText_(texte, today, out.type);
    if (d) {
      out.debut = d.debut || '';
      out.fin = d.fin || '';
      out.infos.push('Dates devinées dans le texte de la page : vérifie-les.');
    } else {
      out.infos.push('Pas de date trouvée : à compléter.');
    }
  }
  if (out.type === 'Ciné' && out.debut && !out.fin) {
    out.fin = plusUnMois_(out.debut);
    out.infos.push('Film : fin mise par défaut à 1 mois après la sortie (modifiable).');
  }
  if (!out.zone) out.zone = guessZone_(out.lieu) || guessZone_(texte.slice(0, 30000), true);

  out.titre = clean_(out.titre).slice(0, 200);
  out.lieu = clean_(out.lieu).slice(0, 200);
  out.desc = clean_(out.desc).slice(0, 400);
  return out;
}

function findLdEvent_(html) {
  const re = /<script[^>]*application\/ld\+json[^>]*>([\s\S]*?)<\/script>/gi;
  const nodes = [];
  let m;
  while ((m = re.exec(html))) {
    try { collect_(JSON.parse(m[1].trim()), nodes); } catch (e) { /* JSON invalide : on ignore */ }
  }
  const evs = nodes.filter(function (n) {
    return /Event|Festival|Exhibition|Movie|Screening/i.test([].concat(n['@type'] || []).join(' '));
  });
  return evs.filter(function (n) { return n.startDate; })[0] || evs[0] || null;
}

function collect_(x, acc) {
  if (Array.isArray(x)) x.forEach(function (y) { collect_(y, acc); });
  else if (x && typeof x === 'object') {
    acc.push(x);
    if (x['@graph']) collect_(x['@graph'], acc);
  }
}

function fromLd_(n, out) {
  const t = [].concat(n['@type'] || []).join(' ');
  out.titre = n.name || n.headline || '';
  out.debut = isoDay_(n.startDate || (/Movie/i.test(t) ? (n.datePublished || n.dateCreated) : ''));
  out.fin = isoDay_(n.endDate);

  const loc = [].concat(n.location || [])[0];
  if (loc) {
    if (typeof loc === 'string') out.lieu = loc;
    else {
      const a = loc.address || {};
      const adr = typeof a === 'string' ? a : [a.streetAddress, a.postalCode, a.addressLocality].filter(Boolean).join(' ');
      out.lieu = [loc.name, adr].filter(Boolean).join(' – ');
      const cp = String((typeof a === 'object' && a.postalCode) || '');
      if (/^\d{5}$/.test(cp)) out.zone = zoneFromCp_(cp);
      else if (typeof a === 'object' && a.addressCountry && !/^(FR|France)$/i.test(String(a.addressCountry.name || a.addressCountry))) out.zone = 'Hors IDF';
    }
  }
  if (!out.zone) out.zone = guessZone_(out.lieu);

  const off = [].concat(n.offers || [])[0];
  if (n.isAccessibleForFree === true || n.isAccessibleForFree === 'true') out.prix = 'Gratuit';
  else if (off) {
    const p = off.price !== undefined ? off.price : off.lowPrice;
    if (p !== undefined && p !== null && p !== '') {
      out.prix = Number(p) === 0 ? 'Gratuit' : p + ' ' + (!off.priceCurrency || off.priceCurrency === 'EUR' ? '€' : off.priceCurrency);
    }
  }
  const img = [].concat(n.image || [])[0];
  out.image = typeof img === 'string' ? img : (img && img.url) || '';
  out.desc = n.description || '';
  out.type = typeFromLd_(t);
}

function typeFromLd_(t) {
  if (/Exhibition/i.test(t)) return 'Expo';
  if (/Music|Festival/i.test(t)) return 'Concert / Festival';
  if (/Movie|Screening/i.test(t)) return 'Ciné';
  if (/Theater|Comedy|Dance/i.test(t)) return 'Spectacle';
  if (/Education|Course/i.test(t)) return 'Atelier';
  return '';
}

function guessType_(s) {
  s = sansAccents_(String(s).toLowerCase());
  const regles = [
    [/expo/, 'Expo'],
    [/concert|\blive\b|tournee|festival/, 'Concert / Festival'],
    [/cinema|\bfilm|allocine|seance/, 'Ciné'],
    [/theatre|spectacle|stand.?up|humour|danse|opera|ballet|cirque|comedie/, 'Spectacle'],
    [/visite|balade/, 'Visite'],
    [/atelier|\bcours\b|masterclass/, 'Atelier']
  ];
  for (let i = 0; i < regles.length; i++) if (regles[i][0].test(s)) return regles[i][1];
  return '';
}

function zoneFromCp_(cp) {
  const d = cp.slice(0, 2);
  if (d === '75') return 'Paris';
  return ['77', '78', '91', '92', '93', '94', '95'].indexOf(d) >= 0 ? 'IDF' : 'Hors IDF';
}

function guessZone_(s, cpSeulement) {
  s = String(s || '');
  const m = s.match(/\b(75|77|78|91|92|93|94|95)\d{3}\b/);
  if (m) return m[1] === '75' ? 'Paris' : 'IDF';
  if (!cpSeulement && /\bparis\b/i.test(s)) return 'Paris';
  return '';
}

function metas_(html) {
  const map = {};
  const re = /<meta\s[^>]*>/gi;
  let m;
  while ((m = re.exec(html))) {
    const tag = m[0];
    const k = (attr_(tag, 'property') || attr_(tag, 'name')).toLowerCase();
    const v = attr_(tag, 'content');
    if (k && v && !map[k]) map[k] = decode_(v);
  }
  return map;
}

function attr_(tag, a) {
  const m = tag.match(new RegExp('\\s' + a + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\')', 'i'));
  return m ? (m[1] !== undefined ? m[1] : m[2]) : '';
}

function titleTag_(html) {
  const m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  return m ? m[1] : '';
}

function pageText_(html) {
  return decode_(html
    .replace(/<(script|style|noscript|svg)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' '))
    .replace(/\s+/g, ' ');
}

function absUrl_(u, base) {
  if (/^https?:\/\//i.test(u)) return u;
  const m = base.match(/^(https?:\/\/[^\/]+)/i);
  if (/^\/\//.test(u)) return 'https:' + u;
  if (m && u.charAt(0) === '/') return m[1] + u;
  return u;
}

function decode_(s) {
  return String(s || '')
    .replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCodePoint(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCodePoint(+d); })
    .replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"').replace(/&(apos|rsquo|lsquo);/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&eacute;/g, 'é').replace(/&egrave;/g, 'è').replace(/&ecirc;/g, 'ê').replace(/&agrave;/g, 'à')
    .replace(/&ccedil;/g, 'ç').replace(/&ocirc;/g, 'ô').replace(/&ucirc;/g, 'û').replace(/&icirc;/g, 'î')
    .replace(/&amp;/g, '&');
}

function clean_(s) {
  return decode_(String(s || '').replace(/<[^>]+>/g, ' ')).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function sansAccents_(s) { return s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }

function isoDay_(s) {
  if (!s) return '';
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  const d = new Date(s);
  return isNaN(d) ? '' : todayIsoFromDate_(d);
}

function todayIsoFromDate_(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

/* ---------- Adresses : lien Google Maps -> nom, adresse, position ---------- */

// Les liens courts (maps.app.goo.gl) redirigent vers l'adresse longue, qui contient le nom et les coordonnées.
function extractPlace(url) {
  url = String(url || '').trim();
  if (!/^https?:\/\/\S+$/i.test(url)) throw new Error('Ce n’est pas un lien valide.');
  const out = { lien: url, nom: '', cat: '', style: '', adresse: '', ville: '', prix: '', statut: 'À tester', lat: '', lng: '', infos: [] };
  let u = url, html = '';
  for (let k = 0; k < 6; k++) {
    let res;
    try {
      res = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: false,
        headers: { 'User-Agent': UA, 'Accept-Language': 'fr-FR,fr;q=0.9' } });
    } catch (e) { break; }
    const code = res.getResponseCode();
    const h = res.getAllHeaders();
    const loc = h.Location || h.location;
    if (code >= 300 && code < 400 && loc) { u = absUrl_(String(loc), u); continue; }
    if (code < 400) html = res.getContentText();
    break;
  }
  placeFromUrl_(u, out);
  if (html) placeFromHtml_(html, out);
  if (out.lat === '' && out.adresse) {
    const g = geocode_(out.adresse);
    if (g) { out.lat = g.lat; out.lng = g.lng; }
  }
  if (!out.adresse && out.lat !== '') out.adresse = reverse_(out.lat, out.lng);
  out.ville = villeDe_(out.adresse);
  out.cat = catAdresse_(out.nom + ' ' + out.style);
  out.nom = clean_(out.nom).slice(0, 150);
  if (!out.nom) out.infos.push('Nom introuvable : complète à la main.');
  if (out.lat === '') out.infos.push('Position introuvable : ajoute l’adresse, elle sera placée sur la carte à l’enregistrement.');
  else out.infos.push('Infos lues depuis le lien : vérifie-les.');
  return out;
}

function placeFromUrl_(u, out) {
  let s = u;
  try { s = decodeURIComponent(u.replace(/\+/g, ' ')); } catch (e) { /* lien mal encodé : on lit tel quel */ }
  let m = s.match(/\/maps\/place\/([^\/@?]+)/);
  if (m) out.nom = m[1].trim();
  m = s.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) ||
    s.match(/[?&](?:q|query|ll|center|destination)=(-?\d+\.\d+),\s*(-?\d+\.\d+)/);
  if (m) { out.lat = +m[1]; out.lng = +m[2]; }
  else if (!out.nom && (m = s.match(/[?&](?:q|query)=([^&]+)/))) out.nom = m[1].trim();
  // « Nom, 12 rue …, 75011 Paris » : le nom d'abord, l'adresse ensuite
  m = out.nom.match(/^([^,]+),\s*(.*\d.*)$/);
  if (m) { out.nom = m[1].trim(); out.adresse = m[2].trim(); }
}

// La page Google Maps donne souvent « Nom · Adresse » en titre, et la catégorie et le prix en description.
function placeFromHtml_(html, out) {
  const meta = metas_(html);
  const titre = meta['og:title'] || '';
  const parts = titre.split(' · ');
  if (parts.length > 1) {
    if (!out.nom) out.nom = parts[0];
    if (!out.adresse) out.adresse = parts.slice(1).join(', ');
  } else if (!out.nom && titre && !/google maps/i.test(titre)) out.nom = titre;
  const desc = meta['og:description'] || '';
  desc.split(' · ').forEach(function (p) {
    p = p.trim();
    if (/^€{1,4}$/.test(p)) out.prix = ['€', '€€', '€€€', '€€€'][p.length - 1];
    else if (!out.style && /restaurant|bar|café|cafe|musée|galerie|boutique|librairie|parc|bistrot|brasserie|cave|salon/i.test(p) && p.length < 50) out.style = p;
  });
  if (out.lat === '') {
    const m = (meta['og:image'] || '').match(/center=(-?\d+\.\d+)%2C(-?\d+\.\d+)/);
    if (m) { out.lat = +m[1]; out.lng = +m[2]; }
  }
}

function catAdresse_(s) {
  s = sansAccents_(String(s).toLowerCase());
  const regles = [
    [/restau|bistro|brasserie|pizz|creperie|ramen|sushi|burger|cantine|traiteur/, 'Restau'],
    [/\bbar\b|pub|cocktail|cave a vin|biere|brewery/, 'Bar'],
    [/cafe|coffee|salon de the|patisserie|boulangerie/, 'Café'],
    [/musee|museum|galerie|fondation|exposition/, 'Musée'],
    [/librairie|boutique|magasin|shop|store|friperie|epicerie|marche/, 'Boutique'],
    [/escape|escalade|bowling|karaoke|spa|hammam|cinema|theatre|piscine|salle/, 'Activité'],
    [/parc|jardin|monument|eglise|basilique|chateau|cimetiere|place|pont/, 'Lieu à visiter']
  ];
  for (let i = 0; i < regles.length; i++) if (regles[i][0].test(s)) return regles[i][1];
  return '';
}

function geocode_(adr) {
  if (!adr) return null;
  try {
    const r = Maps.newGeocoder().setRegion('fr').setLanguage('fr').geocode(adr);
    const x = r && r.results && r.results[0];
    return x ? { lat: x.geometry.location.lat, lng: x.geometry.location.lng, adresse: x.formatted_address } : null;
  } catch (e) { return null; }
}

function reverse_(lat, lng) {
  try {
    const r = Maps.newGeocoder().setLanguage('fr').reverseGeocode(lat, lng);
    const x = r && r.results && r.results[0];
    return x ? x.formatted_address : '';
  } catch (e) { return ''; }
}

// « 12 rue Oberkampf, 75011 Paris » -> « Paris 11e » ; hors Paris, le nom de la ville.
function villeDe_(adr) {
  const s = String(adr || '');
  let m = s.match(/\b75(0\d\d|116)\b/);
  if (m) {
    const n = m[1] === '116' ? 16 : +m[1];
    return n >= 1 && n <= 20 ? 'Paris ' + (n === 1 ? '1er' : n + 'e') : 'Paris';
  }
  m = s.match(/\b\d{5}\s+([^,]+)/);
  if (m) return m[1].trim();
  return /\bparis\b/i.test(s) ? 'Paris' : '';
}

/* Dates en français dans le texte : « du 12 octobre au 3 janvier 2027 », « jusqu'au 15 février », « 12/11/2026 »… */

const MOIS = { janvier: 1, janv: 1, fevrier: 2, fevr: 2, fev: 2, mars: 3, avril: 4, avr: 4, mai: 5, juin: 6, juillet: 7, juil: 7,
  aout: 8, septembre: 9, sept: 9, octobre: 10, oct: 10, novembre: 11, nov: 11, decembre: 12, dec: 12 };

function datesFromText_(texte, today, type) {
  const t = sansAccents_(String(texte).toLowerCase()).replace(/[’`]/g, "'").replace(/\s+/g, ' ').slice(0, 300000);
  const JS = '(?:(?:lun|mar|mer|jeu|ven|sam|dim)[a-z]*\\.?\\s+)?';
  const J = '(\\d{1,2})(?:er)?';
  const M = '(janvier|janv|fevrier|fevr|fev|mars|avril|avr|mai|juin|juillet|juil|aout|septembre|sept|octobre|oct|novembre|nov|decembre|dec)\\.?';
  const Y = '(?:\\s+(\\d{4}))?';
  const ref = today || todayIsoFromDate_(new Date());
  const anRef = +ref.slice(0, 4);
  let m;

  // Film : date de sortie (« 23 septembre 2026 en salle », « date de sortie 8 octobre 2026 »)
  if (type === 'Ciné') {
    m = t.match(new RegExp('\\b' + J + '\\s+' + M + '\\s+(\\d{4})\\s+en salles?')) ||
      t.match(new RegExp('(?:date de sortie|sortie le|en salles? (?:depuis )?le|au cinema le)\\s+' + JS + J + '\\s+' + M + '\\s+(\\d{4})'));
    if (m) return { debut: iso_(+m[3], MOIS[m[2]], +m[1]), fin: '' };
  }

  // (du) 12 (octobre) (2026) au 3 janvier (2027) — on garde la première période pas encore terminée
  const reP = new RegExp('(?:\\b(?:du|depuis le)\\s+)?' + JS + '\\b' + J + '(?:\\s+' + M + ')?' + Y + '\\s*(?:au|-|–|—)\\s*' + JS + J + '\\s+' + M + Y, 'g');
  while ((m = reP.exec(t))) {
    const fin = fin_(+m[4], MOIS[m[5]], m[6], ref, anRef);
    const debut = debut_(+m[1], MOIS[m[2] || m[5]], m[3], fin);
    if (fin && fin >= ref && debut && debut <= fin) return { debut: debut, fin: fin };
  }
  // du 12/10/2026 au 03/01/2027
  m = t.match(/\bdu\s+(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})\s+au\s+(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})/);
  if (m) {
    const debut = iso_(an_(m[3]), +m[2], +m[1]);
    const fin = iso_(an_(m[6]), +m[5], +m[4]);
    if (debut && fin) return { debut: debut, fin: fin };
  }
  // jusqu'au 15 février (2027)
  m = t.match(new RegExp("jusqu'(?:au|a)\\s+(?:le\\s+)?" + JS + J + '\\s+' + M + Y));
  if (m) {
    const fin = fin_(+m[1], MOIS[m[2]], m[3], ref, anRef);
    if (fin) return { debut: '', fin: fin };
  }
  // à partir du 12 novembre (2026)
  m = t.match(new RegExp("a partir du\\s+" + JS + J + '\\s+' + M + Y));
  if (m) {
    const debut = fin_(+m[1], MOIS[m[2]], m[3], ref, anRef);
    if (debut) return { debut: debut, fin: '' };
  }
  if (type === 'Ciné') return null;
  // date unique avec année (concert, spectacle) : première date à venir
  const re = new RegExp(JS + '\\b' + J + '\\s+' + M + '\\s+(\\d{4})', 'g');
  while ((m = re.exec(t))) {
    const d = iso_(+m[3], MOIS[m[2]], +m[1]);
    if (d && d >= ref) return { debut: d, fin: d };
  }
  return null;
}

// Durée d'affiche par défaut d'un film : on ne connaît jamais la dernière séance.
function plusUnMois_(iso) {
  let y = +iso.slice(0, 4), mo = +iso.slice(5, 7) + 1;
  if (mo > 12) { mo = 1; y++; }
  const j = Math.min(+iso.slice(8, 10), new Date(y, mo, 0).getDate());
  return iso_(y, mo, j);
}

function an_(y) { y = +y; return y < 100 ? 2000 + y : y; }

function iso_(y, mo, j) {
  if (!y || !mo || !j || mo > 12 || j > 31) return '';
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + j).slice(-2);
}

// Sans année : année en cours, ou suivante si la date est passée depuis plus de 2 mois.
function fin_(j, mo, y, ref, anRef) {
  if (y) return iso_(+y, mo, j);
  let d = iso_(anRef, mo, j);
  const limite = iso_(anRef, +ref.slice(5, 7), +ref.slice(8, 10));
  if (d && d < limite && (+ref.slice(5, 7) - mo + 12) % 12 > 2) d = iso_(anRef + 1, mo, j);
  return d;
}

function debut_(j, mo, y, fin) {
  if (y) return iso_(+y, mo, j);
  if (!fin) return '';
  let d = iso_(+fin.slice(0, 4), mo, j);
  if (d > fin) d = iso_(+fin.slice(0, 4) - 1, mo, j);
  return d;
}
