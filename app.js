// =========================================================
// Quoi faire ? — le « cerveau » de l'appli (version Firebase).
// L'interface est reprise de l'ancienne version Apps Script ;
// seuls l'enregistrement des données et l'Excel ont changé.
// Plan du fichier :
//   1. Réglages de base et listes de valeurs
//   2. Stockage (Firebase) et lecteur de liens (mini Apps Script)
//   3. Utilitaires
//   4. Idées, événements, calendrier, adresses, Radar, notes (affichage)
//   5. Formulaires
//   6. Apparence, Excel (export / import)
//   7. Clics et démarrage
// =========================================================

import { initializeApp } from "https://www.gstatic.com/firebasejs/12.4.0/firebase-app.js";
import {
  getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect,
  onAuthStateChanged, signOut
} from "https://www.gstatic.com/firebasejs/12.4.0/firebase-auth.js";
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  collection, doc, onSnapshot, setDoc, deleteDoc, writeBatch
} from "https://www.gstatic.com/firebasejs/12.4.0/firebase-firestore.js";
import { firebaseConfig } from "./firebase-config.js";
import { chargerExcelJS, ajouterFeuille, telechargerClasseur, lireClasseur } from "./excel-style.js";

/* ---------- 1. Réglages de base ---------- */

const VERSION = "v1.3.2";
const NOM_APPLI = "app-quoi-faire"; // dossier dans Firestore et nom des fichiers Excel

// Adresse du mini Apps Script « lecteur de liens » (celle qui finit par /exec).
// Vide = pas encore branché : on peut quand même tout remplir à la main.
const LECTEUR_URL = "https://script.google.com/macros/s/AKfycbw7TcvRAan2cbuHaLLqGOlq__GZe64IXZ-XdpRL964U3CkReaUkn-D4fdnrMiwIVoDh/exec";

// Les colonnes de chaque tableau : clé dans l'appli <-> en-tête dans Excel.
// Ce sont les mêmes en-têtes que l'ancien Google Sheet, pour pouvoir le réimporter.
const IDEES = {
  nom: "Idées", arr: "idees",
  // groupe : sous-catégorie (ex. « Balade dans Paris ») ; au tirage, un groupe compte pour une seule idée
  keys: ["id", "nom", "cat", "avec", "duree", "lieu", "ie", "budget", "energie", "derniere", "notes", "epingle", "groupe"],
  headers: ["ID", "Activité", "Catégorie", "Avec qui", "Durée", "Lieu", "Intérieur/Extérieur", "Budget", "Énergie", "Dernière fois", "Notes", "Épinglée", "Groupe"]
};
const EVTS = {
  nom: "Événements", arr: "evts",
  // niveau : prix de 0 à 3 (Gratuit, €, €€, €€€), déduit du prix en euros s'il est vide
  // heure : pour un événement réservé (« 20:00 »), facultative
  keys: ["id", "titre", "type", "lieu", "zone", "debut", "fin", "prix", "lien", "image", "desc", "avec", "statut", "ajoute", "niveau", "heure"],
  headers: ["ID", "Titre", "Type", "Lieu", "Zone", "Début", "Fin", "Prix", "Lien", "Image", "Description", "Avec qui", "Statut", "Ajouté le", "Niveau de prix", "Heure"]
};
// Radar : les sites des lieux à surveiller
const SOURCES = {
  nom: "Sources", arr: "sources",
  keys: ["id", "nom", "lien", "cat", "notes"],
  headers: ["ID", "Nom", "Lien", "Catégorie", "Notes"]
};
// Adresses : restos, bars, musées… à tester ou adorés, avec leur position pour la carte
const ADRESSES = {
  nom: "Adresses", arr: "adresses",
  keys: ["id", "nom", "cat", "style", "adresse", "ville", "prix", "statut", "lien", "lat", "lng", "notes", "epingle", "ajoute"],
  headers: ["ID", "Nom", "Catégorie", "Style", "Adresse", "Ville", "Prix", "Statut", "Lien", "Latitude", "Longitude", "Notes", "Épinglée", "Ajouté le"]
};
const TABLES = [IDEES, EVTS, ADRESSES, SOURCES];
const DATE_KEYS = ["derniere", "debut", "fin", "ajoute"];

const LISTES = {
  avec: ["Solo", "Couple", "Amis"],
  cat: ["Culture", "Sortie", "Sport", "Cuisine", "Créatif", "Apprendre", "Détente", "Maison", "Nature", "Jeux", "Projets"],
  duree: ["Soirée", "Demi-journée", "Journée"],
  lieu: ["Maison", "Paris", "IDF", "Hors IDF"],
  ie: ["Intérieur", "Extérieur", "Les deux"],
  budget: ["Gratuit", "€", "€€", "€€€"],
  energie: ["Chill", "Moyen", "Actif"],
  type: ["Expo", "Ciné", "Concert / Festival", "Spectacle", "Visite", "Atelier", "Autre"],
  zone: ["Paris", "IDF", "Hors IDF"],
  statut: ["À voir", "Réservé", "Fait", "Pas intéressé"],
  srcCat: ["Salles de concert", "Sites d’info", "Salles de spectacle", "Ciné", "Musées"],
  adrCat: ["Restau", "Bar", "Café", "Musée", "Lieu à visiter", "Boutique", "Activité"],
  adrStatut: ["À tester", "Adoré"]
};

// Mode démo (adresse se terminant par ?demo) : données gardées dans ce navigateur
// seulement, sans connexion. Sert à tester l'appli sans compte.
const MODE_DEMO = new URLSearchParams(location.search).has("demo");

// Lien reçu via le menu « Partager » d'Android : il arrive dans url, text ou title selon l'appli d'origine.
let SHARED = MODE_DEMO ? "" : (() => {
  const p = new URLSearchParams(location.search);
  const s = [p.get("url"), p.get("text"), p.get("title")].filter(Boolean).join(" ");
  // Recharger l'appli ne doit pas rajouter le même lien une deuxième fois.
  if (s) history.replaceState(null, "", location.pathname);
  return s;
})();

const S = {
  idees: [], evts: [], sources: [], adresses: [], notes: "", L: LISTES, tab: "idees", picks: [],
  pret: false, compte: "",
  // filtres communs au tirage et à la liste d'idées
  f: { avec: "Couple", duree: "Soirée", ou: "Peu importe", meteo: "Peu importe", budget: "Tout" },
  mode: "tirage", q: "", tri: "cat", filt: false, lf: { cat: "Toutes", pin: "Toutes" }, open: {},
  ef: { type: "Tout", zone: "Tout", prix: "Tout" }, archive: false, later: false,
  af: { statut: "Tous", cat: "Toutes", prix: "Tout" }, aview: "Liste",
  // view : Cartes / Liste / Calendrier ; calMode : Semaine / Mois
  fopen: {}, // filtres dépliés, par onglet
  addOpen: {}, // cadre « Ajouter » déplié, par onglet
  view: "Cartes", calMode: "Semaine", calWeek: "", calMonth: "", calDay: "",
  // vues ouvertes au démarrage, choisies dans Paramètres (synchronisées entre appareils)
  reglages: { vueIdees: "tirage", vueEvts: "Cartes" }
};
const REG_DEFAUT = { vueIdees: "tirage", vueEvts: "Cartes" };
const MO = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."];

/* ---------- 2. Stockage : Firebase ---------- */
// Rangement dans Firestore :
//   users/{mon uid}/app-quoi-faire/donnees            -> { notes }
//   users/{mon uid}/app-quoi-faire/donnees/idees/{id} -> une idée par document (pareil pour evts, sources, adresses)
// Un document par élément : deux appareils qui modifient deux idées différentes ne se gênent pas.
// Si les deux modifient la même : la dernière modification gagne.

let db = null, auth = null, uid = null, arrets = [];
const coll = arr => collection(db, "users", uid, NOM_APPLI, "donnees", arr);
const refDonnees = () => doc(db, "users", uid, NOM_APPLI, "donnees");

function erreurEcriture(err) {
  console.error(err);
  toast("Échec de l’enregistrement");
}
// On n'attend pas la réponse du serveur : hors ligne, Firestore garde l'écriture
// en attente et l'envoie tout seul au retour du réseau.
function ecrire(arr, o) {
  if (MODE_DEMO) return sauverDemo();
  const propre = {};
  Object.keys(o).forEach(k => { if (o[k] !== undefined && k !== "st") propre[k] = o[k]; });
  setDoc(doc(coll(arr), String(o.id)), propre).catch(erreurEcriture);
}
function effacer(arr, id) {
  if (MODE_DEMO) return sauverDemo();
  deleteDoc(doc(coll(arr), String(id))).catch(erreurEcriture);
}

// Mode démo : tout l'état est recopié dans le navigateur, juste après chaque modification
let demoT = null;
function sauverDemo() {
  clearTimeout(demoT);
  demoT = setTimeout(() => {
    const d = { idees: S.idees, evts: S.evts, sources: S.sources, adresses: S.adresses, notes: S.notes, reglages: S.reglages };
    try { localStorage.setItem("qf-demo", JSON.stringify(d)); } catch (e) { /* stockage plein ou bloqué */ }
  }, 0);
}
function demarrerDemo() {
  let d = {};
  try { d = JSON.parse(localStorage.getItem("qf-demo")) || {}; } catch (e) { /* rien de gardé */ }
  ["idees", "evts", "sources", "adresses"].forEach(arr => { S[arr] = Array.isArray(d[arr]) ? d[arr] : []; });
  S.notes = String(d.notes || "");
  S.reglages = Object.assign({}, REG_DEFAUT, d.reglages);
  uid = "demo";
  S.compte = "Mode démo (données sur cet appareil seulement)";
  majSynchro();
  demarrerAffichage();
}

const today_ = () => today();
function newId() {
  try { return crypto.randomUUID().slice(0, 8); } catch (e) { return Math.random().toString(16).slice(2, 10); }
}
// Un identifiant de document ne peut pas contenir « / »
const idPropre = id => String(id == null ? "" : id).trim().replace(/\//g, "-").replace(/^\.+$/, "");

// Enregistre un élément : on garde toutes les colonnes du tableau (vides si besoin)
// et la date de création « cree », qui sert à garder l'ordre d'ajout.
function upsert(T, o) {
  if (!o.id) o.id = newId();
  const out = {};
  T.keys.forEach(k => { out[k] = o[k] === undefined || o[k] === null ? "" : o[k]; });
  out.cree = o.cree || Date.now();
  ecrire(T.arr, out);
  return out;
}

// Prix en euros -> niveau 0 à 3 : gratuit, jusqu'à 15 €, jusqu'à 40 €, au-delà.
function niveauPrix(prix) {
  const s = String(prix == null ? "" : prix).toLowerCase();
  if (/gratuit|libre|free/.test(s)) return "Gratuit";
  const m = s.replace(",", ".").match(/[0-9]+([.][0-9]+)?/);
  if (!m) return "";
  const n = +m[0];
  return n === 0 ? "Gratuit" : n <= 15 ? "€" : n <= 40 ? "€€" : "€€€";
}

// « 12 rue Oberkampf, 75011 Paris » -> « Paris 11e » ; hors Paris, le nom de la ville.
function villeDe(adr) {
  const s = String(adr || "");
  let m = s.match(/\b75(0\d\d|116)\b/);
  if (m) {
    const n = m[1] === "116" ? 16 : +m[1];
    return n >= 1 && n <= 20 ? "Paris " + (n === 1 ? "1er" : n + "e") : "Paris";
  }
  m = s.match(/\b\d{5}\s+([^,]+)/);
  if (m) return m[1].trim();
  return /\bparis\b/i.test(s) ? "Paris" : "";
}

// Le mini Apps Script qui lit les pages web (le navigateur n'a pas le droit de le faire lui-même).
// Il ne reçoit que le lien ou le texte à lire, jamais tes données.
async function lecteur(action, params) {
  if (!LECTEUR_URL) throw new Error("lecteur de liens pas encore branché");
  if (!navigator.onLine) throw new Error("pas de réseau");
  // Envoi en « texte simple » : c'est ce qu'Apps Script accepte depuis une autre adresse.
  const rep = await fetch(LECTEUR_URL, { method: "POST", body: JSON.stringify(Object.assign({ action }, params)) });
  const r = await rep.json();
  if (!r.ok) throw new Error(r.erreur || "lecture impossible");
  return r.data;
}

// Les mêmes noms de fonctions que l'ancien serveur Apps Script :
// l'interface les appelle avec call('saveIdee', …) comme avant.
const API = {
  saveIdee: o => upsert(IDEES, o),
  deleteIdee: id => effacer("idees", id),
  saveEvt: o => {
    if (!o.statut) o.statut = "À voir";
    sansFinInventee(o);
    if (!o.ajoute) o.ajoute = today_();
    if (!o.niveau) o.niveau = niveauPrix(o.prix);
    return upsert(EVTS, o);
  },
  deleteEvt: id => effacer("evts", id),
  saveSource: o => upsert(SOURCES, o),
  deleteSource: id => effacer("sources", id),
  // Sans position connue, on demande au lecteur de liens de placer l'adresse sur la carte.
  saveAdresse: async o => {
    if (!o.statut) o.statut = "À tester";
    if (!o.ajoute) o.ajoute = today_();
    if ((o.lat === "" || o.lat === undefined || o.lat === null) && o.adresse) {
      try {
        const g = await lecteur("geocode", { adresse: o.adresse });
        if (g) { o.lat = g.lat; o.lng = g.lng; }
      } catch (e) { /* hors ligne ou pas branché : enregistrée sans position */ }
    }
    if (!o.ville) o.ville = villeDe(o.adresse);
    return upsert(ADRESSES, o);
  },
  deleteAdresse: id => effacer("adresses", id),
  saveNotes: txt => {
    if (MODE_DEMO) { sauverDemo(); return true; }
    setDoc(refDonnees(), { notes: String(txt || "").slice(0, 500000) }, { merge: true }).catch(erreurEcriture);
    return true;
  },
  extractEvent: url => lecteur("event", { url }),
  parseEventText: (texte, type, titre) => lecteur("texte", { texte, type, titre }),
  extractPlace: url => lecteur("place", { url })
};
function sauverReglages() {
  if (MODE_DEMO) return sauverDemo();
  setDoc(refDonnees(), { reglages: S.reglages }, { merge: true }).catch(erreurEcriture);
}

// Film : le lecteur de liens met une fin « au hasard » à +1 mois. On ne la garde pas :
// un film n'a pas de date de fin connue (le calendrier l'affiche quand même sur 1 mois).
function sansFinInventee(o) {
  if (o.type === "Ciné" && o.debut && o.fin === plusUnMois(o.debut)) o.fin = "";
  return o;
}

function call(fn, ...args) {
  return Promise.resolve().then(() => API[fn](...args))
    .catch(e => { toast("Oups : " + (e && e.message || e)); throw e; });
}

// Transforme en texte toujours dans le même ordre (Firestore peut renvoyer
// les champs dans un autre ordre) : sert à voir si quelque chose a vraiment changé.
const enTexte = v => JSON.stringify(v, (k, x) =>
  x && typeof x === "object" && !Array.isArray(x) ? Object.fromEntries(Object.entries(x).sort()) : x);
const parCree = (a, b) => (a.cree || 0) - (b.cree || 0);

// Indicateur de synchro : synchronisé / en attente d'envoi / hors ligne
const enAttente = {};
function majSynchro(cle, meta) {
  if (cle) enAttente[cle] = !!(meta && meta.hasPendingWrites);
  const el = $("#sync");
  if (!el || !uid) { if (el) el.hidden = true; return; }
  const etat = !navigator.onLine ? "horsligne" : Object.values(enAttente).some(Boolean) ? "attente" : "ok";
  const txt = { ok: "synchro", attente: "en attente", horsligne: "hors ligne" }[etat];
  el.className = "sync " + etat;
  el.title = { ok: "Tout est enregistré en ligne", attente: "Modifications pas encore envoyées", horsligne: "Hors ligne : tout est gardé sur l’appareil" }[etat];
  el.innerHTML = ICON["nuage_" + etat] + txt;
  el.hidden = false;
}
window.addEventListener("online", () => majSynchro());
window.addEventListener("offline", () => majSynchro());

// Un changement venu d'un autre appareil pendant que je tape : on attend que j'aie fini.
let renduEnAttente = false;
function rerendre() {
  if (!S.pret) return;
  const a = document.activeElement;
  if (a && a.closest && a.closest("#main") && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName)) { renduEnAttente = true; return; }
  render();
}
document.addEventListener("focusout", () => { if (renduEnAttente) setTimeout(() => { renduEnAttente = false; rerendre(); }, 50); });

function ecouter() {
  const attendus = new Set(["idees", "evts", "sources", "adresses", "donnees"]);
  const premier = cle => { if (attendus.delete(cle) && !attendus.size) demarrerAffichage(); };
  ["idees", "evts", "sources", "adresses"].forEach(arr => {
    arrets.push(onSnapshot(coll(arr), { includeMetadataChanges: true }, snap => {
      majSynchro(arr, snap.metadata);
      const recu = snap.docs.map(d => d.data()).sort(parCree);
      if (enTexte(recu) !== enTexte(S[arr])) {
        S[arr] = recu;
        if (arr === "idees") S.picks = S.picks.filter(id => recu.some(i => i.id === id));
        rerendre();
      }
      premier(arr);
    }, err => { console.error(err); toast("Accès refusé aux données"); }));
  });
  arrets.push(onSnapshot(refDonnees(), { includeMetadataChanges: true }, snap => {
    majSynchro("donnees", snap.metadata);
    const n = snap.exists() ? String(snap.data().notes || "") : "";
    // réglages : ils ne changent pas la vue déjà ouverte, seulement celle du prochain démarrage
    const rg = snap.exists() && snap.data().reglages;
    if (rg) S.reglages = Object.assign({}, REG_DEFAUT, rg);
    // pas d'écrasement pendant que j'écris mes notes
    if (n !== S.notes && !notesSale && document.activeElement?.id !== "notes") {
      S.notes = n;
      if (S.tab === "notes") rerendre();
    }
    premier("donnees");
  }, err => console.error(err)));
}

function demarrerAffichage() {
  S.pret = true;
  // vues de départ choisies dans Paramètres
  S.mode = S.reglages.vueIdees === "liste" ? "liste" : "tirage";
  S.view = ["Cartes", "Liste", "Calendrier"].includes(S.reglages.vueEvts) ? S.reglages.vueEvts : "Cartes";
  suggest();
  // lien partagé : Google Maps -> Adresses, le reste -> À venir
  const lien = SHARED; SHARED = "";
  if (lien && estLienMaps(lien)) { S.tab = "adresses"; render(); extractAdr(lien); }
  else if (lien) { S.tab = "evts"; render(); extract(lien); }
  else render(true);
}

function demarrerFirebase() {
  const app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  // Cache local permanent : l'appli fonctionne hors ligne avec la dernière copie connue
  db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });
  onAuthStateChanged(auth, user => {
    arrets.forEach(f => f()); arrets = [];
    if (!user) { uid = null; S.pret = false; majSynchro(); renderConnexion(); return; }
    uid = user.uid;
    S.compte = user.displayName || user.email || "";
    ecouter();
  });
}

async function seConnecter(btn) {
  const fournisseur = new GoogleAuthProvider();
  const err = $("#lerr");
  if (err) err.textContent = "";
  btn.disabled = true;
  try {
    await signInWithPopup(auth, fournisseur);
  } catch (e) {
    // Si la fenêtre de connexion est bloquée, on passe par une redirection
    if (e.code === "auth/popup-blocked" || e.code === "auth/operation-not-supported-in-this-environment") signInWithRedirect(auth, fournisseur);
    else if (e.code !== "auth/popup-closed-by-user" && e.code !== "auth/cancelled-popup-request" && err) err.textContent = "Connexion impossible (" + e.code + ").";
  } finally { btn.disabled = false; }
}

function renderConnexion() {
  $("#sub").textContent = new Date().toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" });
  $("#main").innerHTML = `<div class="empty"><span class="big">👋</span>Connecte-toi avec ton compte Google<br>pour retrouver tes idées sur tous tes appareils.
    <div class="actions" style="justify-content:center;margin-top:16px"><button class="btn primary" data-act="login">Se connecter avec Google</button></div>
    <div class="meta" id="lerr" style="margin-top:10px"></div></div>`;
}

/* ---------- 3. Utilitaires ---------- */
const $ = s => document.querySelector(s);
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const list = v => String(v || "").split(",").map(s => s.trim()).filter(Boolean);
const rank = (arr, v) => Math.max(0, arr.indexOf(v));
function today() { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
function days(a, b) { return Math.round((Date.parse(b + "T12:00:00") - Date.parse(a + "T12:00:00")) / 864e5); }
function fmt(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-").map(Number);
  return d + (d === 1 ? "er" : "") + " " + MO[m - 1] + (y !== new Date().getFullYear() ? " " + y : "");
}
// annuler : fonction appelée si on touche « Annuler » (le message reste alors un peu plus longtemps)
let toastT;
function toast(msg, annuler) {
  let t = $(".toast");
  if (!t) { t = document.createElement("div"); t.className = "toast"; document.body.appendChild(t); }
  t.innerHTML = "";
  const span = document.createElement("span");
  span.textContent = msg;
  t.appendChild(span);
  if (annuler) {
    const b = document.createElement("button");
    b.textContent = "Annuler";
    b.onclick = () => { t.hidden = true; annuler(); };
    t.appendChild(b);
  }
  t.hidden = false;
  clearTimeout(toastT); toastT = setTimeout(() => t.hidden = true, annuler ? 6000 : 3200);
}
const plur = (n, mot) => n + " " + mot + (n > 1 ? "s" : "");

/* ---------- 4. Idées : filtres et tirage ---------- */
// Prix de 0 à 3 partout : Gratuit, €, €€, €€€. Le filtre garde ce qui ne dépasse pas le niveau choisi.
const PRIX_MAX = { Gratuit: 0, "€ max": 1, "€€ max": 2 };
const PRIX_CHOIX = ["Tout", "Gratuit", "€ max", "€€ max"];
function prixOk(niveau, choix) {
  if (!(choix in PRIX_MAX)) return true;
  const r = S.L.budget.indexOf(niveau);
  return r < 0 ? choix !== "Gratuit" : r <= PRIX_MAX[choix];
}

function matches(i) {
  const f = S.f;
  if (f.avec !== "Tous" && !list(i.avec).includes(f.avec)) return false;
  if (rank(S.L.duree, i.duree) > rank(S.L.duree, f.duree)) return false;
  if (f.ou === "Maison" && i.lieu !== "Maison") return false;
  if (f.ou === "Dehors" && !["Paris", "IDF"].includes(i.lieu)) return false;
  if (f.ou === "Hors IDF" && i.lieu !== "Hors IDF") return false;
  if (f.meteo === "Il pleut" && i.ie === "Extérieur") return false;
  return prixOk(i.budget, f.budget);
}
const epingle = i => !!i.epingle && i.epingle !== "non";
const grp = i => String(i.groupe || "").trim();
const auHasard = arr => arr[Math.floor(Math.random() * arr.length)];

// Une entrée par idée seule, et une seule par groupe quelle que soit sa taille.
function entrees(pool) {
  const out = [], parG = {};
  pool.forEach(i => {
    const g = grp(i);
    if (!g) out.push({ items: [i] });
    else if (parG[g]) parG[g].items.push(i);
    else out.push(parG[g] = { g, items: [i] });
  });
  return out;
}
// Les idées épinglées (qui collent aux filtres) restent dans le tirage ; le hasard complète.
// Un groupe tiré propose une de ses idées au hasard (sous-tirage).
function suggest() {
  const ents = entrees(S.idees.filter(matches));
  const pris = [];
  ents.forEach(e => { const p = e.items.find(epingle); if (p && pris.length < 3) pris.push({ e, i: p }); });
  const reste = ents.filter(e => !pris.some(p => p.e === e));
  while (pris.length < 3 && reste.length) {
    const e = reste.splice(Math.floor(Math.random() * reste.length), 1)[0];
    pris.push({ e, i: auHasard(e.items) });
  }
  S.picks = pris.map(p => p.i.id);
}
async function togglePin(id) {
  const i = S.idees.find(x => x.id === id);
  const avant = i.epingle;
  i.epingle = epingle(i) ? "" : "oui";
  render();
  toast(i.epingle ? "📌 Épinglée : elle reste dans tes tirages" : "Désépinglée");
  try { await call("saveIdee", i); } catch (e) { i.epingle = avant; render(); }
}

/* ---------- événements ---------- */
// « À voir » et « Réservé » : encore à venir. « Fait » et « Pas intéressé » : rangés dans les passés.
const actif = e => !e.statut || e.statut === "À voir" || e.statut === "Réservé";
const reserve = e => e.statut === "Réservé";

// Trois genres de dates :
//   jour    : un seul jour précis (concert, atelier…) — début = fin, ou pas de fin pour ces types-là
//   ouvert  : pas de fin connue (films à l'affiche, spectacles en tournée…)
//   periode : une date de fin (expos, festivals…)
const JOUR_UNIQUE = ["Concert / Festival", "Atelier", "Visite"];
function evGenre(e) {
  if (!e.debut) return e.fin ? "periode" : "sansdate";
  if (e.type === "Ciné") return "ouvert";
  if (e.fin === e.debut || (!e.fin && JOUR_UNIQUE.includes(e.type))) return "jour";
  return e.fin ? "periode" : "ouvert";
}
// Fin « réelle » pour ranger dans les passés : un film disparaît 1 mois après sa sortie.
function finEffective(e) {
  const g = evGenre(e);
  if (g === "jour") return e.debut;
  if (e.type === "Ciné" && e.debut) return e.fin || plusUnMois(e.debut);
  return e.fin;
}
const dansN = n => n <= 0 ? "Aujourd’hui" : n === 1 ? "Demain" : n < 15 ? "Dans " + n + " j" : "Dans " + Math.round(n / 7) + " sem.";

function evState(e) {
  const t = today(), g = evGenre(e), fin = finEffective(e);
  if (!actif(e)) return { arch: true, label: e.statut, key: e.fin || e.debut || "" };
  if (fin && fin < t) return { arch: true, label: "Terminé", key: fin };
  if (e.debut && e.debut > t) { const n = days(t, e.debut); return { label: dansN(n), cls: "soon", key: e.debut }; }
  if (g === "jour") return { label: "Aujourd’hui", cls: "urgent", key: e.debut };
  if (g === "ouvert") return { label: "À l’affiche", key: fin || "9998", ouvert: true };
  if (e.fin) {
    const n = days(t, e.fin);
    const label = n === 0 ? "Dernier jour !" : n < 15 ? "Finit dans " + n + " j" : "Finit dans " + Math.round(n / 7) + " sem.";
    return { label, cls: n <= 14 ? "urgent" : "", key: e.fin };
  }
  return { label: "Sans date", key: "9999" };
}

// « 20:00 » -> « 20 h », « 20:30 » -> « 20 h 30 »
const fmtHeure = h => { const m = String(h || "").match(/^(\d{1,2}):(\d{2})/); return m ? +m[1] + " h" + (m[2] !== "00" ? " " + m[2] : "") : ""; };
const jourSemaine = iso => new Date(iso + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "short" });

// La ligne de date, avec l'icône de son genre : [icône, texte]
function dateLigne(e) {
  const g = evGenre(e), t = today(), h = fmtHeure(e.heure);
  if (g === "sansdate") return [ICON.g_sans, "sans date"];
  if (g === "jour") return [ICON.g_jour, jourSemaine(e.debut) + " " + fmt(e.debut) + (h ? " · " + h : "")];
  if (g === "ouvert") return [ICON.g_ouvert, (e.debut <= t ? "depuis le " : "sortie le ") + fmt(e.debut) + (h ? " · " + h : "")];
  if (!e.debut || e.debut <= t) return [ICON.g_periode, "jusqu’au " + fmt(e.fin)];
  return [ICON.g_periode, fmt(e.debut) + " → " + fmt(e.fin)];
}
// Période déjà commencée : part du temps écoulé, en % (barre de temps restant)
function progression(e) {
  if (evGenre(e) !== "periode" || !e.debut || !e.fin) return null;
  const t = today();
  if (e.debut > t || e.fin < t) return null;
  const total = days(e.debut, e.fin);
  return total > 0 ? Math.round(100 * days(e.debut, t) / total) : null;
}
// Étiquettes de la carte : compte à rebours, et « Réservé » si c'est le cas
function badges(e) {
  const h = fmtHeure(e.heure);
  return `<span class="badge ${e.st.cls || ""}">${esc(e.st.label)}</span>` +
    (reserve(e) ? `<span class="badge resa">${ICON.check}Réservé${h && evGenre(e) !== "jour" ? " · " + esc(h) : ""}</span>` : "");
}
const lieuPrix = e => [e.lieu || e.zone, e.prix !== "" && e.prix != null ? (typeof e.prix === "number" ? e.prix + " €" : e.prix) : niveauEvt(e)].filter(Boolean).join(" · ");
function evSorted() {
  return S.evts.map(e => Object.assign({ st: evState(e) }, e)).sort((a, b) => a.st.key < b.st.key ? -1 : a.st.key > b.st.key ? 1 : 0);
}

/* ---------- émojis ---------- */
const EMO_CAT = { Culture: "🏛️", Sortie: "🍹", Sport: "🏃", Cuisine: "🍳", Créatif: "🎨", Apprendre: "📚", Détente: "🛋️",
  Maison: "🏠", Nature: "🌿", Jeux: "🎲", Projets: "🗺️" };
const EMO_TYPE = { Expo: "🖼️", Ciné: "🎬", "Concert / Festival": "🎶", Concert: "🎶", Festival: "🎶", Spectacle: "🎭",
  Visite: "🧭", Atelier: "✂️", Autre: "✨" };
const catEmo = c => EMO_CAT[c] || "💡";

// Petites illustrations plates multicolores des catégories d'idées (thèmes Suisse).
// Une par catégorie de base ; une ampoule pour les catégories créées en plus.
const ill = d => `<svg viewBox="0 0 24 24" aria-hidden="true">${d}</svg>`;
const ILLU = {
  Culture: ill('<path d="M12 3l9 5H3z" fill="#ff6f61"/><rect x="3" y="8" width="18" height="2" fill="#2f3e8f"/><rect x="5" y="10" width="2.4" height="7" fill="#fff"/><rect x="10.8" y="10" width="2.4" height="7" fill="#fff"/><rect x="16.6" y="10" width="2.4" height="7" fill="#fff"/><rect x="3" y="17" width="18" height="3" rx="1" fill="#4e69c6"/>'),
  Sortie: ill('<path d="M5 4h14l-7 8z" fill="#ff4d8d"/><path d="M7.6 6.6h8.8L12 11.4z" fill="#ffb3cf"/><rect x="11.2" y="11.6" width="1.6" height="6.6" fill="#5a6b8c"/><rect x="8" y="18.2" width="8" height="2" rx="1" fill="#5a6b8c"/><circle cx="17" cy="5" r="2.6" fill="#ffd23f"/><path d="M8 2l3 6" stroke="#2ec4b6" stroke-width="1.6" stroke-linecap="round"/>'),
  Sport: ill('<circle cx="12" cy="12" r="9" fill="#ff8a3d"/><path d="M3 12h18M12 3v18M5.6 5.6c3 3 3 9.8 0 12.8M18.4 5.6c-3 3-3 9.8 0 12.8" stroke="#9c3d0c" stroke-width="1.4" fill="none"/>'),
  Cuisine: ill('<path d="M3.5 12.5h17a8.5 6.5 0 0 1-17 0z" fill="#e2574c"/><rect x="3" y="11.5" width="18" height="1.6" rx=".8" fill="#9e2f27"/><circle cx="8.5" cy="10.5" r="2" fill="#ffd23f"/><circle cx="12.5" cy="9.8" r="1.8" fill="#4caf50"/><circle cx="15.8" cy="10.8" r="1.5" fill="#ff9f1c"/><path d="M9.5 6c0-1.6 1.6-1.6 1.6-3.2M13.5 6c0-1.6 1.6-1.6 1.6-3.2" stroke="#8a94a8" stroke-width="1.4" fill="none" stroke-linecap="round"/>'),
  "Créatif": ill('<path d="M12 3c5 0 9 3.6 9 8 0 2.4-1.8 3.4-3.6 3.4h-2c-1.3 0-2 1-1.4 2.2.7 1.4-.2 3.4-2 3.4C7 20 3 16 3 11.5S7 3 12 3z" fill="#f6d6a8"/><circle cx="7.5" cy="10.5" r="1.7" fill="#e2574c"/><circle cx="10.5" cy="7" r="1.7" fill="#ffd23f"/><circle cx="15" cy="7.2" r="1.7" fill="#4caf50"/><circle cx="17.6" cy="10.8" r="1.7" fill="#4e69c6"/>'),
  Apprendre: ill('<path d="M2 6c3-1.5 7-1.5 10 1v13c-3-2.5-7-2.5-10-1z" fill="#7c6cf0"/><path d="M22 6c-3-1.5-7-1.5-10 1v13c3-2.5 7-2.5 10-1z" fill="#a99df7"/><path d="M4 8c2.4-.8 4.6-.6 6.5.8V18c-1.9-1.3-4.1-1.5-6.5-.8z" fill="#fff"/><path d="M20 8c-2.4-.8-4.6-.6-6.5.8V18c1.9-1.3 4.1-1.5 6.5-.8z" fill="#fff"/><path d="M16 7v6l1.3-1 1.3 1V6.6z" fill="#ff4d6d"/>'),
  "Détente": ill('<path d="M4 9h13v5a5 5 0 0 1-5 5H9a5 5 0 0 1-5-5z" fill="#2ec4b6"/><path d="M17 10.5h1.5a2.5 2.5 0 0 1 0 5H17" stroke="#1b8a80" stroke-width="1.8" fill="none"/><rect x="3" y="19" width="15" height="2" rx="1" fill="#f6d6a8"/><rect x="5.5" y="9" width="10" height="2" fill="#a7f0e6"/><path d="M8.5 6c0-1.4 1.3-1.4 1.3-2.8M12 6c0-1.4 1.3-1.4 1.3-2.8" stroke="#8a94a8" stroke-width="1.4" fill="none" stroke-linecap="round"/>'),
  Maison: ill('<rect x="5" y="10" width="14" height="10" fill="#ffd23f"/><path d="M2.5 11L12 3l9.5 8z" fill="#e2574c"/><rect x="10" y="14" width="4" height="6" fill="#8d5a3b"/><rect x="6.5" y="12.5" width="2.6" height="2.6" fill="#74c0fc"/><rect x="14.9" y="12.5" width="2.6" height="2.6" fill="#74c0fc"/>'),
  Nature: ill('<circle cx="18" cy="6" r="3" fill="#ffd23f"/><rect x="10.9" y="14" width="2.2" height="7" fill="#8d5a3b"/><circle cx="12" cy="10" r="6" fill="#4caf50"/><circle cx="9.5" cy="8.5" r="2.4" fill="#7ed67f"/><path d="M3 21h18" stroke="#2e7d32" stroke-width="2" stroke-linecap="round"/>'),
  Jeux: ill('<rect x="10" y="3" width="11" height="11" rx="2.5" fill="#4caf50" transform="rotate(12 15.5 8.5)"/><rect x="3" y="9" width="12" height="12" rx="2.6" fill="#e2574c"/><circle cx="6.5" cy="12.5" r="1.2" fill="#fff"/><circle cx="9" cy="15" r="1.2" fill="#fff"/><circle cx="11.5" cy="17.5" r="1.2" fill="#fff"/><circle cx="15.5" cy="6.5" r="1.1" fill="#fff"/><circle cx="17.5" cy="10" r="1.1" fill="#fff"/>'),
  Projets: ill('<path d="M3 6.5l6-3 6 3 6-3v14l-6 3-6-3-6 3z" fill="#bfe8d4"/><path d="M9 3.5l6 3v14l-6-3z" fill="#7fd1ae"/><path d="M5 14c2-1.2 3 1 5 0s3-3.2 5-2.2 3 2.2 4 1" stroke="#ff7a1a" stroke-width="1.8" fill="none" stroke-dasharray="2 1.6" stroke-linecap="round"/><circle cx="18.5" cy="12.2" r="2" fill="#e2574c"/>')
};
const ILLU_DEFAUT = ill('<path d="M12 3a6 6 0 0 0-3.6 10.8c.6.5 1 1.2 1 2V17h5.2v-1.2c0-.8.4-1.5 1-2A6 6 0 0 0 12 3z" fill="#ffd23f"/><rect x="9.4" y="17.6" width="5.2" height="2" rx=".8" fill="#8a94a8"/><rect x="10.2" y="20" width="3.6" height="1.6" rx=".8" fill="#5a6b8c"/><path d="M10.5 9.5l1.5 2 1.5-2" stroke="#e08a00" stroke-width="1.3" fill="none"/>');
const catIllu = c => `<span class="ill" data-c="${esc(c || "")}">${ILLU[c] || ILLU_DEFAUT}</span>`;
const typeEmo = t => EMO_TYPE[t] || "✨";

const svg = d => `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
const NUAGE = '<path d="M7 18a4.5 4.5 0 0 1-.6-9 6 6 0 0 1 11.4 1.6A3.7 3.7 0 0 1 17.5 18z"/>';
const ICON = {
  check: svg('<path d="M5 12.5l4.5 4.5L19 7.5"/>'),
  // genres de dates : un seul jour / sans fin connue / période — et le billet « réservé »
  g_jour: svg('<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4M15 3v4"/><circle cx="12" cy="15" r="1.4" fill="currentColor"/>'),
  g_ouvert: svg('<circle cx="5" cy="12" r="2" fill="currentColor"/><path d="M7 12h3M13 12h2M18 12h.01M21 12h.01"/>'),
  g_periode: svg('<circle cx="5" cy="12" r="2" fill="currentColor"/><circle cx="19" cy="12" r="2" fill="currentColor"/><path d="M7 12h10"/>'),
  g_sans: svg('<circle cx="12" cy="12" r="8"/><path d="M12 8v4l2.5 2"/>'),
  billet: svg('<path d="M4 8h16v3a2 2 0 0 0 0 4v3H4v-3a2 2 0 0 0 0-4z"/><path d="M9 12.5l2 2 4-4"/>'),
  // mini pop-up
  dots: svg('<circle cx="5" cy="12" r="1.3" fill="currentColor"/><circle cx="12" cy="12" r="1.3" fill="currentColor"/><circle cx="19" cy="12" r="1.3" fill="currentColor"/>'),
  x: svg('<path d="M6 6l12 12M18 6L6 18"/>'),
  trash: svg('<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>'),
  ouvrir: svg('<path d="M14 4h6v6"/><path d="M20 4l-9 9"/><path d="M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>'),
  iti: svg('<path d="M3 11l18-8-8 18-2-8z"/>'),
  shuffle: svg('<path d="M16 3h5v5"/><path d="M4 20L21 3"/><path d="M21 16v5h-5"/><path d="M15 15l6 6"/><path d="M4 4l5 5"/>'),
  edit: svg('<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>'),
  next: svg('<path d="M5 12h14"/><path d="M13 6l6 6-6 6"/>'),
  filtre: svg('<path d="M4 5h16l-6 7.5V18l-4 2v-7.5z"/>'),
  pin: svg('<path d="M9 3h6l-1 6 3.5 3.5V15h-11v-2.5L10 9z"/><path d="M12 15v6"/>'),
  spark: svg('<path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/><path d="M19 15l.7 2 2 .7-2 .7-.7 2-.7-2-2-.7 2-.7z"/>'),
  prev: svg('<path d="M19 12H5"/><path d="M11 6l-6 6 6 6"/>'),
  map: svg('<path d="M12 21s-6-5.6-6-11a6 6 0 0 1 12 0c0 5.4-6 11-6 11z"/><circle cx="12" cy="10" r="2.2"/>'),
  target: svg('<circle cx="12" cy="12" r="7"/><circle cx="12" cy="12" r="2.5"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3"/>'),
  list: svg('<path d="M9 6h11M9 12h11M9 18h11"/><path d="M4.5 6h.01M4.5 12h.01M4.5 18h.01"/>'),
  dice: svg('<rect x="4" y="4" width="16" height="16" rx="3"/><circle cx="9" cy="9" r=".6" fill="currentColor"/><circle cx="15" cy="15" r=".6" fill="currentColor"/><circle cx="15" cy="9" r=".6" fill="currentColor"/><circle cx="9" cy="15" r=".6" fill="currentColor"/>'),
  // indicateur de synchro
  nuage_ok: svg(NUAGE + '<path d="M9.5 13.5l2 2 3.5-3.5"/>'),
  nuage_attente: svg(NUAGE + '<path d="M12 15.5v-4M10.3 13l1.7-1.7 1.7 1.7"/>'),
  nuage_horsligne: svg(NUAGE + '<path d="M3 3l18 18"/>')
};

/* ---------- rendu ---------- */
// acc : couleur d'accent de la ligne (1 à 5), pour distinguer les filtres d'un coup d'œil
function chipRow(label, key, opts, store, icon, acc) {
  return `<div class="row" data-acc="${acc || 1}"><span class="lbl">${label}</span><div class="chips">${opts.map(o =>
    `<button class="chip ${store[key] === o ? "on" : ""}" data-act="${store === S.f ? "f" : store === S.af ? "af" : "ef"}" data-k="${key}" data-v="${esc(o)}">${icon && icon(o) ? `<span class="emoji">${icon(o)} </span>` : ""}${esc(o)}</button>`).join("")}</div></div>`;
}

// Filtres repliés derrière une petite pastille qui résume les choix (en gras : un vrai choix,
// point orange : au moins un filtre actif). On la touche pour déplier les lignes de pastilles.
// lignes : [libellé, clé, choix possibles, où ranger le choix, icône, couleur d'accent]
const NEUTRES = ["Tout", "Tous", "Toutes", "Peu importe"];
const RESUME_NEUTRE = { ou: "partout", budget: "tout prix", prix: "tout prix", type: "tous types", zone: "toutes zones", statut: "tous statuts" };
function blocFiltres(zone, lignes) {
  const morceaux = [];
  let actifs = 0;
  lignes.forEach(([, key, , store]) => {
    const v = store[key];
    if (NEUTRES.includes(v)) { if (RESUME_NEUTRE[key]) morceaux.push(RESUME_NEUTRE[key]); }
    else { actifs++; morceaux.push(`<b>${esc(v)}</b>`); }
  });
  const ouvert = !!S.fopen[zone];
  return `<button class="fresume ${ouvert ? "open" : ""}" data-act="fopen" data-v="${zone}" aria-expanded="${ouvert}" title="Filtres">${ICON.filtre}<span>${morceaux.join(" · ")}</span>${actifs ? '<i class="dot"></i>' : ""}</button>` +
    (ouvert ? `<div class="fpli">${lignes.map(([l, k, o, st, ic, acc]) => chipRow(l, k, o, st, ic, acc)).join("")}</div>` : "");
}

function renderIdees() {
  const t = today();
  // films à l'affiche exclus : ils n'ont pas de vraie date limite
  const urgents = evSorted().filter(e => !e.st.arch && !e.st.ouvert && e.st.key !== "9999" && days(t, e.st.key) <= 14).slice(0, 3);
  const quand = { Soirée: "ce soir", "Demi-journée": "cet aprèm", Journée: "aujourd’hui" }[S.f.duree] || "aujourd’hui";
  const seg = [["tirage", ICON.dice, "Tirage"], ["liste", ICON.list, "Liste"]].map(([m, ic, l]) =>
    `<button class="${S.mode === m ? "on" : ""}" data-act="mode" data-v="${m}" aria-pressed="${S.mode === m}">${ic} ${l}</button>`).join("");
  return `<section class="hero"><h2>On fait quoi ${quand}&nbsp;?</h2></section>` +
    (urgents.length ? `<div class="alert"><div class="eyebrow">À ne pas rater</div>${urgents.map(e =>
      `<div class="ln"><span>${esc(e.titre)}</span><strong>${esc(e.st.label)}</strong></div>`).join("")}</div>` : "") +
    blocFiltres("idees", [
      ["Avec qui", "avec", ["Tous"].concat(S.L.avec), S.f, null, 1],
      ["J’ai", "duree", S.L.duree.slice(), S.f, null, 2],
      ["Où", "ou", ["Peu importe", "Maison", "Dehors", "Hors IDF"], S.f, null, 3],
      ["Météo", "meteo", ["Peu importe", "Il pleut"], S.f, null, 4],
      ["Prix", "budget", PRIX_CHOIX, S.f, null, 5]]) +
    (S.mode === "liste" ? renderListe(seg) : renderTirage(seg));
}

function renderTirage(seg) {
  const picks = S.picks.map(id => S.idees.find(i => i.id === id)).filter(Boolean);
  const n = entrees(S.idees.filter(matches)).length;
  const vide = S.idees.length
    ? "Aucune idée pour ces filtres.<br>Élargis un peu, ou ajoute des idées depuis la liste."
    : "Pas encore d’idées.<br>Ajoute-en depuis la liste, ou importe ton ancien fichier dans Paramètres.";
  return `<div class="modebar"><div class="seg">${seg}</div><span class="grow"></span>
      ${n > picks.length ? `<button class="btn primary sm" data-act="suggest">${ICON.dice} Nouveau tirage</button>` : ""}</div>` +
    (picks.length ? `<div class="pgrid">${picks.map(ideaCard).join("")}</div>` : `<div class="empty"><span class="big">🫥</span>${vide}</div>`);
}

// Petite carte : les 3 idées tiennent côte à côte sur un téléphone.
// Pas de boutons : la toucher ouvre le mini pop-up (épingler, fait aujourd'hui, modifier…).
function ideaCard(i) {
  const pin = epingle(i);
  return `<div class="card pc ${pin ? "pinned" : ""}" data-act="idee-pop" data-id="${esc(i.id)}">
    ${pin ? `<span class="pinmark" aria-label="Épinglée">${ICON.pin}</span>` : ""}
    <div class="emo" data-c="${esc(i.cat)}">${catEmo(i.cat)}</div>${catIllu(i.cat)}
    <div class="kcat">${esc(i.cat || "")}</div>
    ${grp(i) ? `<div class="gk">${esc(grp(i))}</div>` : ""}
    <h3>${esc(i.nom)}</h3>
    <div class="ptags">${[i.lieu, i.budget].filter(Boolean).map(esc).join(" · ")}</div>
  </div>`;
}

// Titre de page (le grand titre et le compteur ne s'affichent que dans les thèmes Suisse)
// avec, à droite, le bouton « + Ajouter » qui déplie le cadre d'ajout.
function titrePage(zone, titre, compte) {
  const ouvert = !!S.addOpen[zone];
  return `<div class="ptitle"><h2>${titre}</h2><span class="pcount">${esc(compte)}</span>
    <button class="addbtn ${ouvert ? "on" : ""}" data-act="add-open" data-v="${zone}" aria-expanded="${ouvert}">${ouvert ? "× Fermer" : "<b>+</b> Ajouter"}</button></div>`;
}
// Cadre d'ajout : coller un lien (lu automatiquement) ou remplir à la main
function blocAjout(zone, titre, idChamp, aide, actLire, idBtn, actMain) {
  if (!S.addOpen[zone]) return "";
  return `<div class="card addc"><div class="lbl">${titre}</div>
    <div class="add"><input type="url" id="${idChamp}" placeholder="${aide}" inputmode="url">
      <button class="btn primary" data-act="${actLire}" id="${idBtn}">Lire</button>
      <button class="btn" data-act="${actMain}">${ICON.edit} À la main</button></div></div>`;
}

function renderEvts() {
  const all = evSorted();
  const vis = all.filter(e => (S.ef.type === "Tout" || e.type === S.ef.type) && (S.ef.zone === "Tout" || e.zone === S.ef.zone) && prixOk(niveauEvt(e), S.ef.prix));
  const act = vis.filter(e => !e.st.arch), arch = vis.filter(e => e.st.arch);
  // par défaut : ce qui est déjà ouvert ou commence dans les 7 jours ; le reste derrière un lien
  const lim = addDays(today(), 7);
  // les réservés remontent en tête (le tri par date est gardé à l'intérieur de chaque groupe)
  const now = act.filter(e => !e.debut || e.debut <= lim).sort((a, b) => reserve(b) - reserve(a));
  const later = act.filter(e => e.debut && e.debut > lim).sort((a, b) => reserve(b) - reserve(a));
  // titre de page (affiché seulement par certains thèmes, ex. Suisse)
  return titrePage("evts", "À venir", act.length + " en cours") +
    blocAjout("evts", "Ajouter un événement", "url", "Coller un lien…", "extract", "go", "new-evt") +
    `<div class="row"><div class="chips">${["Cartes", "Liste", "Calendrier"].map(v =>
      `<button class="chip ${S.view === v ? "on" : ""}" data-act="view" data-v="${v}">${v}</button>`).join("")}</div></div>` +
    blocFiltres("evts", [
      ["Type", "type", ["Tout"].concat(S.L.type), S.ef, o => EMO_TYPE[o], 2],
      ["Zone", "zone", ["Tout"].concat(S.L.zone), S.ef, null, 3],
      ["Prix", "prix", PRIX_CHOIX, S.ef, null, 5]]) +
    (S.view === "Calendrier" ? renderCal(vis) :
    (now.length ? grilleEvts(now) : `<div class="empty"><span class="big">🌱</span>Rien en cours cette semaine.<br>${later.length ? "Regarde plus loin juste en dessous, ou colle" : "Colle"} un lien juste au-dessus !</div>`) +
    (later.length ? `<button class="link" data-act="later">${S.later ? "Masquer" : "Voir"} ce qui commence plus tard (${later.length})</button>` +
      (S.later ? grilleEvts(later) : "") + "<br>" : "") +
    (arch.length ? `<button class="link" data-act="archive">${S.archive ? "Masquer" : "Voir"} les passés et faits (${arch.length})</button>` +
      (S.archive ? grilleEvts(arch) : "") : ""));
}
// Cartes (2 par ligne) ou liste (une ligne par événement), selon la vue choisie
const grilleEvts = arr => S.view === "Liste"
  ? `<div class="evlist">${arr.map(evLigne).join("")}</div>`
  : `<div class="evgrid">${arr.map(evCard).join("")}</div>`;

/* ---------- calendrier ---------- */
function addDays(iso, n) {
  const d = new Date(iso + "T12:00:00"); d.setDate(d.getDate() + n);
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}
function plusUnMois(iso) {
  let [y, m, d] = iso.split("-").map(Number);
  m++; if (m > 12) { m = 1; y++; }
  d = Math.min(d, new Date(y, m, 0).getDate());
  return y + "-" + String(m).padStart(2, "0") + "-" + String(d).padStart(2, "0");
}
// Période pendant laquelle l'événement est « ouvert » dans le calendrier.
// Sans fin connue (films…) : affiché sur 1 mois après le début, avec une barre qui s'efface.
function evRange(e) {
  const g = evGenre(e);
  if (g === "jour") return [e.debut, e.debut];
  if (g === "ouvert") return [e.debut, e.fin || plusUnMois(e.debut)];
  if (e.debut) return [e.debut, e.fin || e.debut];
  if (e.fin) { const d = e.ajoute && e.ajoute < e.fin ? e.ajoute : today(); return [d < e.fin ? d : e.fin, e.fin]; }
  return null;
}
const lundi = iso => addDays(iso, -((new Date(iso + "T12:00:00").getDay() + 6) % 7));
function shiftMonth(k) {
  let [y, m] = S.calMonth.split("-").map(Number);
  m += k; if (m < 1) { m = 12; y--; } if (m > 12) { m = 1; y++; }
  S.calMonth = y + "-" + String(m).padStart(2, "0");
}
// Flèches du calendrier : une semaine ou un mois de plus / de moins
function shiftCal(k) {
  if (S.calMode === "Mois") return shiftMonth(k);
  S.calWeek = addDays(S.calWeek, 7 * k);
  S.calDay = addDays(S.calDay, 7 * k);
}
// Bascule semaine / mois : on garde le jour sélectionné
function setCalMode(m) {
  S.calMode = m;
  S.calMonth = S.calDay.slice(0, 7);
  S.calWeek = lundi(S.calDay);
}

// Une rangée de 7 jours avec les barres des événements (sert pour la semaine et pour le mois).
// large : vue semaine (nom des jours, heure dans les blocs d'un seul jour)
function rangeeSemaine(ws, periodes, t, mois, large) {
  const we = addDays(ws, 6);
  const jours = [...Array(7)].map((_, i) => addDays(ws, i));
  const lignes = []; // dernière colonne occupée par ligne de frise
  const barres = periodes.filter(x => x.r[0] <= we && x.r[1] >= ws).map(x => {
    const s = x.r[0] < ws ? 0 : jours.indexOf(x.r[0]);
    const f = x.r[1] > we ? 6 : jours.indexOf(x.r[1]);
    let k = lignes.findIndex(fin => fin < s);
    if (k < 0) { k = lignes.length; lignes.push(f); } else lignes[k] = f;
    const unJour = x.r[0] === x.r[1];
    const txt = unJour && large && fmtHeure(x.e.heure) ? fmtHeure(x.e.heure) : x.e.titre;
    const cls = [unJour && "jour", evGenre(x.e) === "ouvert" && x.r[1] <= we && "ouvert",
      x.r[0] < ws && "cont-l", x.r[1] > we && "cont-r", reserve(x.e) && "resa"].filter(Boolean).join(" ");
    return `<div class="bar ${cls}" data-t="${esc(x.e.type)}" style="grid-column:${s + 1}/${f + 2};grid-row:${k + 2}">${esc(txt)}</div>`;
  }).join("");
  const cases = jours.map((iso, i) => {
    const cls = [mois && iso.slice(0, 7) !== mois && "out", iso < t && "past", iso === t && "today", iso === S.calDay && "sel"].filter(Boolean).join(" ");
    return `<button class="day ${cls}" style="grid-column:${i + 1}" data-act="day" data-v="${iso}">${large ? `<small>${jourSemaine(iso)}</small>` : ""}<span class="num">${+iso.slice(8)}</span></button>`;
  }).join("");
  return `<div class="week ${large ? "large" : ""}" style="grid-template-rows:${large ? 40 : 24}px ${lignes.length ? `repeat(${lignes.length}, 18px)` : ""} 6px">${cases}${barres}</div>`;
}

function renderCal(vis) {
  const t = today();
  if (!S.calDay) S.calDay = t;
  if (!S.calMonth) S.calMonth = S.calDay.slice(0, 7);
  if (!S.calWeek) S.calWeek = lundi(S.calDay);
  const evs = vis.filter(actif).map(e => ({ e, r: evRange(e) })).filter(x => x.r);
  const sansDate = vis.filter(e => actif(e) && !evRange(e)).length;
  const on = iso => evs.filter(x => x.r[0] <= iso && iso <= x.r[1]);
  // Tout est dans la frise ; les événements d'un seul jour sont des blocs encadrés.
  const periodes = evs.slice().sort((a, b) => a.r[0] < b.r[0] ? -1 : a.r[0] > b.r[0] ? 1 : (a.r[1] > b.r[1] ? -1 : 1));

  let titre, corps;
  if (S.calMode === "Mois") {
    const [y, m] = S.calMonth.split("-").map(Number);
    const debutGrille = lundi(S.calMonth + "-01");
    const finMois = S.calMonth + "-" + String(new Date(y, m, 0).getDate()).padStart(2, "0");
    corps = `<div class="dows">${["L", "M", "M", "J", "V", "S", "D"].map(d => `<div class="dow">${d}</div>`).join("")}</div>`;
    for (let ws = debutGrille; ws <= finMois; ws = addDays(ws, 7)) corps += rangeeSemaine(ws, periodes, t, S.calMonth, false);
    titre = new Date(y, m - 1, 1).toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  } else {
    const a = S.calWeek, b = addDays(a, 6);
    const moisDe = iso => new Date(iso + "T12:00:00").toLocaleDateString("fr-FR", { month: "long" });
    titre = a.slice(5, 7) === b.slice(5, 7) ? `${+a.slice(8)} – ${+b.slice(8)} ${moisDe(b)}` : `${+a.slice(8)} ${moisDe(a)} – ${+b.slice(8)} ${moisDe(b)}`;
    corps = rangeeSemaine(a, periodes, t, "", true);
  }

  const types = [...new Set(evs.map(x => x.e.type).filter(Boolean))];
  const jour = on(S.calDay).map(x => Object.assign({ st: evState(x.e) }, x.e))
    .sort((a, b) => reserve(b) - reserve(a) || ((a.fin || "9999") < (b.fin || "9999") ? -1 : 1));
  const titreJour = new Date(S.calDay + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" });
  const unite = S.calMode === "Mois" ? "Mois" : "Semaine";

  return `<div class="row"><div class="chips">${["Semaine", "Mois"].map(v =>
      `<button class="chip ${S.calMode === v ? "on" : ""}" data-act="calmode" data-v="${v}">${v}</button>`).join("")}</div></div>
    <div class="calhead"><button class="ibtn" data-act="cal-prev" aria-label="${unite} précédent${unite === "Semaine" ? "e" : ""}">${ICON.prev}</button><b>${titre}</b><button class="ibtn go" data-act="cal-next" aria-label="${unite} suivant${unite === "Semaine" ? "e" : ""}">${ICON.next}</button></div>
    <div class="cal">${corps}</div>
    <div class="legend">${types.map(ty => `<span><i class="sw" data-t="${esc(ty)}"></i>${esc(ty)}</span>`).join("")}
      <span><i class="sw jour"></i>un seul jour</span><span><i class="sw ouvert"></i>sans fin connue</span><span><i class="sw resa"></i>réservé</span>
      ${sansDate ? `<span>${sansDate} sans date, visible${sansDate > 1 ? "s" : ""} dans la liste</span>` : ""}</div>
    <div class="dayhead">${titreJour}</div>
    ${jour.length ? `<div class="evlist">${jour.map(evLigne).join("")}</div>` : '<div class="empty">Rien ce jour-là.</div>'}`;
}

const niveauEvt = e => e.niveau || niveauPrix(e.prix);
const imageEvt = e => e.image ? `<img src="${esc(e.image)}" referrerpolicy="no-referrer" alt="" loading="lazy" onerror="this.remove()">` : "";

// Carte compacte : image basse avec seulement les étiquettes, puis titre sur 2 lignes,
// ligne de date avec l'icône de son genre, barre de temps restant, lieu · prix.
// Pas de boutons : la toucher ouvre le mini pop-up.
function evCard(e) {
  const [ic, quand] = dateLigne(e);
  const p = progression(e);
  return `<article class="card evc ${reserve(e) ? "reserve" : ""}" data-act="ev-pop" data-id="${esc(e.id)}">
    <div class="cover" data-t="${esc(e.type)}"><span class="cemo">${typeEmo(e.type)}</span>${imageEvt(e)}
      <div class="bds">${badges(e)}</div>
    </div>
    <div class="evbody">
      <div class="kline"><span>${esc(e.type || "")}</span>${reserve(e) ? '<span class="ok">✓ Réservé</span>' : `<span class="${e.st.cls || ""}">${esc(e.st.label)}</span>`}</div>
      <h3>${esc(e.titre || "(sans titre)")}</h3>
      <div class="quand">${ic}<span>${esc(quand)}</span></div>
      ${p !== null ? `<div class="prog"><i style="width:${p}%"></i></div>` : ""}
      <div class="evinfo"><div>${esc(lieuPrix(e))}</div></div>
    </div></article>`;
}

// Mode liste : une ligne par événement, petite image à gauche (toucher la ligne = mini pop-up)
function evLigne(e) {
  const [ic, quand] = dateLigne(e);
  return `<div class="item evl ${reserve(e) ? "reserve" : ""}" data-act="ev-pop" data-id="${esc(e.id)}">
    <div class="th" data-t="${esc(e.type)}"><span>${typeEmo(e.type)}</span>${imageEvt(e)}</div>
    <div class="tx"><b>${esc(e.titre || "(sans titre)")}</b><div class="quand">${ic}<span>${esc(quand)}</span></div><div class="meta">${esc(lieuPrix(e))}</div></div>
    <div class="dr">${badges(e)}</div></div>`;
}

const TRIS = [["cat", "Par catégorie"], ["nom", "A → Z"]];
const LF_DEFAUT = { cat: "Toutes", pin: "Toutes" };

// En plus des filtres du haut : catégorie et épinglées (repliés derrière l'entonnoir)
function lfMatch(i) {
  const f = S.lf;
  return (f.cat === "Toutes" || i.cat === f.cat) && (f.pin === "Toutes" || epingle(i));
}

function renderListe(seg) {
  const cats = ["Toutes"].concat([...new Set(S.idees.map(i => i.cat).filter(Boolean))].sort((a, b) => a.localeCompare(b, "fr")));
  const actif = S.tri !== "cat" || Object.keys(LF_DEFAUT).some(k => S.lf[k] !== LF_DEFAUT[k]);
  const ligne = (label, k, opts, icone) => `<div class="row"><span class="lbl">${label}</span><div class="chips">${opts.map(o =>
    `<button class="chip ${S.lf[k] === o ? "on" : ""}" data-act="lf" data-k="${k}" data-v="${esc(o)}">${icone && icone(o) ? `<span class="emoji">${icone(o)} </span>` : ""}${esc(o)}</button>`).join("")}</div></div>`;
  return `<div class="modebar"><div class="seg">${seg}</div>
      <input type="search" id="q" placeholder="Rechercher…" value="${esc(S.q)}">
      <button class="ibtn fbtn ${S.filt ? "open" : ""} ${actif ? "actif" : ""}" data-act="filt" title="Trier et filtrer" aria-label="Trier et filtrer">${ICON.filtre}</button></div>
    ${S.filt ? `<div class="fpanel">
      <div class="row"><span class="lbl">Trier</span><div class="chips">${TRIS.map(([k, l]) =>
        `<button class="chip ${S.tri === k ? "on" : ""}" data-act="tri" data-v="${k}">${l}</button>`).join("")}</div></div>
      ${ligne("Catégorie", "cat", cats, c => c === "Toutes" ? "" : catEmo(c))}
      ${ligne("Afficher", "pin", ["Toutes", "Épinglées"], o => o === "Épinglées" ? "📌" : "")}
      ${actif ? `<button class="link" data-act="lf-reset">Tout réinitialiser</button>` : ""}
    </div>` : ""}
    <div id="items">${renderItems()}</div>
    <button class="btn primary fab" data-act="new-idee">+ Nouvelle idée</button>`;
}

// Catégorie d'un groupe : la plus fréquente parmi toutes ses idées.
function catGroupe(g) {
  const n = {};
  S.idees.forEach(i => { if (grp(i) === g && i.cat) n[i.cat] = (n[i.cat] || 0) + 1; });
  return Object.keys(n).sort((a, b) => n[b] - n[a])[0] || "";
}

function renderItems() {
  const q = S.q.trim().toLowerCase();
  const items = S.idees.filter(i => matches(i) && lfMatch(i) && (!q || (i.nom + " " + grp(i)).toLowerCase().includes(q)));
  if (!items.length) return '<div class="empty">Aucune idée pour ces filtres.</div>';
  const nom = (a, b) => String(a.nom).localeCompare(String(b.nom), "fr");
  const ligne = (i, cls, avecGroupe) => { const on = epingle(i); return `<div class="item li ${cls || ""}" data-act="edit-idee" data-id="${esc(i.id)}"><div class="emo" data-c="${esc(i.cat)}">${catEmo(i.cat)}</div>${catIllu(i.cat)}
      <div class="txt">${esc(i.nom)}${avecGroupe && grp(i) ? `<span class="gm">${esc(grp(i))}</span>` : ""}</div>
      <button class="pact ${on ? "on" : ""}" data-act="pin" data-id="${esc(i.id)}" title="${on ? "Désépingler" : "Épingler"}" aria-label="${on ? "Désépingler" : "Épingler"} ${esc(i.nom)}" aria-pressed="${on}">${ICON.pin}</button></div>`; };

  // les épinglées d'abord, dans un bloc à part ; le reste en idées seules et en groupes repliables
  const pins = items.filter(epingle).sort(nom);
  const unites = entrees(items.filter(i => !epingle(i))).map(e => e.g
    ? { nom: e.g, g: e.g, cat: catGroupe(e.g), items: e.items.sort(nom) }
    : { nom: e.items[0].nom, cat: e.items[0].cat, i: e.items[0] });
  const parNom = (a, b) => String(a.nom).localeCompare(String(b.nom), "fr");
  unites.sort(S.tri === "cat" ? (a, b) => String(a.cat).localeCompare(String(b.cat), "fr") || parNom(a, b) : parNom);

  let cat = null;
  const reste = unites.map(u => {
    // en tri par catégorie, un petit intertitre à chaque nouvelle catégorie
    const titre = S.tri === "cat" && u.cat !== cat ? `<div class="ghead"><span class="emoji">${catEmo(u.cat)} </span>${esc(u.cat || "Sans catégorie")}</div>` : "";
    cat = u.cat;
    if (!u.g) return titre + ligne(u.i);
    const ouvert = !!S.open[u.g] || !!q;
    return titre + `<div class="item li grp ${ouvert ? "open" : ""}" data-act="grp" data-v="${esc(u.g)}" aria-expanded="${ouvert}">
        <div class="emo" data-c="${esc(u.cat)}">${catEmo(u.cat)}</div>${catIllu(u.cat)}<div class="txt">${esc(u.g)}</div>
        <span class="gn">${u.items.length}</span><span class="chev">›</span></div>` +
      (ouvert ? u.items.map(i => ligne(i, "sub")).join("") : "");
  }).join("");
  return `<div class="count">${plur(items.length, "idée")}</div>` +
    (pins.length ? `<div class="pinbloc"><div class="ghead"><span class="emoji">📌 </span>Épinglées</div>${pins.map(i => ligne(i, "", true)).join("")}</div>` : "") +
    (pins.length && unites.length && S.tri !== "cat" ? '<div class="ghead">Toutes les autres</div>' : "") + reste;
}

/* ---------- Adresses : à tester ou adorées, en liste ou sur une carte ---------- */
const EMO_ADR = { Restau: "🍽️", Bar: "🍸", Café: "☕", Musée: "🏛️", "Lieu à visiter": "📍", Boutique: "🛍️", Activité: "🎯" };
const adrEmo = c => EMO_ADR[c] || "📌";
const aPos = a => a.lat !== "" && a.lat != null && !isNaN(+a.lat) && +a.lat !== 0;
// catégories par défaut d'abord, puis celles créées au fil de l'eau
const adrCats = () => [...new Set((S.L.adrCat || []).concat(S.adresses.map(a => a.cat).filter(Boolean)))];
const adrVisibles = () => S.adresses.filter(a => (S.af.statut === "Tous" || a.statut === S.af.statut) &&
  (S.af.cat === "Toutes" || a.cat === S.af.cat) && prixOk(a.prix, S.af.prix));
const quoiChercher = a => encodeURIComponent([a.nom, a.adresse].filter(Boolean).join(" "));
const mapsUrl = a => a.lien || "https://www.google.com/maps/search/?api=1&query=" + quoiChercher(a);
const itineraire = a => "https://www.google.com/maps/dir/?api=1&destination=" + (aPos(a) ? a.lat + "," + a.lng : quoiChercher(a));
// lien Google Maps (partagé depuis l'appli Maps ou copié)
const estLienMaps = s => /maps[.]app[.]goo[.]gl|goo[.]gl[/]maps|google[.][a-z.]+[/]maps|maps[.]google[.]/i.test(s);

function renderAdresses() {
  const vis = adrVisibles();
  const cats = ["Toutes"].concat(adrCats().filter(c => S.adresses.some(a => a.cat === c)));
  return titrePage("adresses", "Adresses", String(S.adresses.length)) +
    blocAjout("adresses", "Ajouter une adresse", "aurl", "Coller un lien Maps…", "adr-extract", "ago", "new-adr") +
    `<div class="row"><div class="chips">${["Liste", "Carte"].map(v =>
      `<button class="chip ${S.aview === v ? "on" : ""}" data-act="aview" data-v="${v}">${v}</button>`).join("")}</div></div>` +
    blocFiltres("adresses", [
      ["Statut", "statut", ["Tous"].concat(S.L.adrStatut), S.af, o => ({ Adoré: "♥", "À tester": "🔖" })[o], 1]]
      .concat(cats.length > 2 ? [["Type", "cat", cats, S.af, c => EMO_ADR[c], 4]] : [])
      .concat([["Prix", "prix", PRIX_CHOIX, S.af, null, 5]])) +
    (S.aview === "Carte" ? renderCarte(vis) : renderAdrListe(vis));
}

// dist : distance affichée à la place du prix (liste « les plus proches »)
function adrLigne(a, dist) {
  const on = epingle(a);
  const meta = [a.style, a.ville].filter(Boolean).join(" · ");
  // toucher la ligne : mini pop-up (Maps, itinéraire, épingler, modifier…)
  return `<div class="item li adr" data-act="adr-pop" data-id="${esc(a.id)}"><div class="emo" data-c="${esc(a.cat)}">${adrEmo(a.cat)}</div>
    <div class="txt">${esc(a.nom)}${meta ? `<span class="gm">${esc(meta)}</span>` : ""}</div>
    ${a.statut === "Adoré" ? '<span class="heart" title="Adoré">♥</span>' : ""}
    ${on ? `<span class="pinmark" aria-label="Épinglée">${ICON.pin}</span>` : ""}
    ${dist || a.prix ? `<span class="when">${esc(dist || a.prix)}</span>` : ""}</div>`;
}

function renderAdrListe(vis) {
  if (!S.adresses.length) return `<div class="empty"><span class="big">📍</span>Aucune adresse pour l’instant.<br>Dans Google Maps : Partager → Copier le lien, puis colle-le ici.</div>`;
  if (!vis.length) return '<div class="empty">Aucune adresse pour ces filtres.</div>';
  const nom = (a, b) => String(a.nom).localeCompare(String(b.nom), "fr");
  const pins = vis.filter(epingle).sort(nom);
  const ordre = adrCats();
  const autres = vis.filter(a => !epingle(a)).sort((a, b) => (ordre.indexOf(a.cat) - ordre.indexOf(b.cat)) || nom(a, b));
  let cat = null;
  return `<div class="count">${plur(vis.length, "adresse")}</div>` +
    (pins.length ? `<div class="pinbloc"><div class="ghead"><span class="emoji">📌 </span>Épinglées</div>${pins.map(a => adrLigne(a)).join("")}</div>` : "") +
    autres.map(a => {
      const titre = a.cat !== cat ? `<div class="ghead"><span class="emoji">${adrEmo(a.cat)} </span>${esc(a.cat || "Autres")}</div>` : "";
      cat = a.cat;
      return titre + adrLigne(a);
    }).join("");
}

function renderCarte(vis) {
  const sans = vis.filter(a => !aPos(a)).length;
  return `<div class="mapbar"><button class="btn sm" data-act="locate">${ICON.target} Autour de moi</button>
      <span class="meta">${sans ? plur(sans, "adresse") + " sans position" : ""}</span></div>
    <div id="map"></div><div id="near"></div>`;
}

/* carte : Leaflet + OpenStreetMap, chargés seulement quand on ouvre la carte (version fixe) */
let leafletLib = null, LMAP = null, moiMk = null, MOI = null;
function loadLeaflet() {
  return leafletLib || (leafletLib = new Promise((ok, ko) => {
    const base = "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/";
    const css = document.createElement("link");
    css.rel = "stylesheet"; css.href = base + "leaflet.min.css";
    document.head.appendChild(css);
    const s = document.createElement("script");
    s.src = base + "leaflet.min.js";
    s.onload = () => ok(window.L);
    s.onerror = () => { leafletLib = null; ko(new Error("carte indisponible")); };
    document.head.appendChild(s);
  }));
}
function fermerCarte() { if (LMAP) { LMAP.remove(); LMAP = null; moiMk = null; } }

async function drawMap() {
  const el = $("#map");
  if (!el) return;
  let L;
  try { L = await loadLeaflet(); } catch (e) { el.innerHTML = '<div class="empty">Carte indisponible (pas de connexion ?)</div>'; return; }
  if (!document.body.contains(el)) return;
  fermerCarte();
  const pts = adrVisibles().filter(aPos);
  LMAP = L.map(el, { zoomControl: false });
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 19, attribution: "© OpenStreetMap" }).addTo(LMAP);
  pts.forEach(a => L.circleMarker([+a.lat, +a.lng], { radius: 8, weight: 2, color: "#fff", fillOpacity: 1,
    fillColor: a.statut === "Adoré" ? "#ff641f" : "#4e69c6" }).bindPopup(popupAdr(a)).addTo(LMAP));
  if (MOI) { addMoi(); LMAP.setView(MOI, 15); }
  else if (pts.length) LMAP.fitBounds(L.latLngBounds(pts.map(a => [+a.lat, +a.lng])), { padding: [30, 30], maxZoom: 15 });
  else LMAP.setView([48.8566, 2.3522], 12);
  renderNear();
}

function popupAdr(a) {
  return `<b>${esc(a.nom)}</b>${a.statut === "Adoré" ? ' <span style="color:#ff641f">♥</span>' : ""}<br>
    <span>${esc([a.style, a.ville, a.prix].filter(Boolean).join(" · "))}</span><br>
    <a href="${esc(itineraire(a))}" target="_blank" rel="noopener">Itinéraire</a> · <button class="link" data-act="edit-adr" data-id="${esc(a.id)}">Modifier</button>`;
}

function addMoi() {
  if (!LMAP || !MOI) return;
  if (moiMk) moiMk.remove();
  moiMk = window.L.circleMarker(MOI, { radius: 7, weight: 3, color: "#fff", fillColor: "#1f8a5b", fillOpacity: 1 }).addTo(LMAP).bindTooltip("Toi");
}

function locate(btn) {
  if (!navigator.geolocation) return toast("Position indisponible sur cet appareil");
  btn.disabled = true;
  navigator.geolocation.getCurrentPosition(p => {
    btn.disabled = false;
    MOI = [p.coords.latitude, p.coords.longitude];
    addMoi(); if (LMAP) LMAP.setView(MOI, 15);
    renderNear();
  }, () => { btn.disabled = false; toast("Position refusée ou indisponible ici"); },
  { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 });
}

// distance à vol d'oiseau, en km
function dist(p, q) {
  const r = x => x * Math.PI / 180;
  const a = Math.sin(r(q[0] - p[0]) / 2) ** 2 + Math.cos(r(p[0])) * Math.cos(r(q[0])) * Math.sin(r(q[1] - p[1]) / 2) ** 2;
  return 12742 * Math.asin(Math.sqrt(a));
}
const fmtDist = d => d < 1 ? Math.max(50, Math.round(d * 20) * 50) + " m" : d.toFixed(1).replace(".", ",") + " km";

function renderNear() {
  const box = $("#near");
  if (!box) return;
  if (!MOI) { box.innerHTML = ""; return; }
  const proches = adrVisibles().filter(aPos).map(a => ({ a, d: dist(MOI, [+a.lat, +a.lng]) })).sort((x, y) => x.d - y.d).slice(0, 6);
  box.innerHTML = proches.length ? '<div class="ghead">Les plus proches</div>' + proches.map(x => adrLigne(x.a, fmtDist(x.d))).join("") : "";
}

async function togglePinAdr(id) {
  const a = S.adresses.find(x => x.id === id);
  const avant = a.epingle;
  a.epingle = epingle(a) ? "" : "oui";
  render();
  toast(a.epingle ? "📌 Épinglée" : "Désépinglée");
  try { await call("saveAdresse", a); } catch (e) { a.epingle = avant; render(); }
}

async function extractAdr(text) {
  const mots = String(text || "").split("\n").join(" ").split(" ");
  const url = mots.find(w => /^https?:\/\/./i.test(w));
  if (!url) return toast("Colle un lien qui commence par http");
  const dup = S.adresses.find(a => a.lien === url);
  if (dup) { toast("Déjà dans tes adresses"); return openForm("adr", dup); }
  const btn = $("#ago");
  if (btn) { btn.disabled = true; btn.textContent = "Lecture…"; }
  try {
    const out = await call("extractPlace", url);
    const infos = out.infos; delete out.infos;
    openForm("adr", out, infos);
  } catch (e) {
    openForm("adr", { lien: url, statut: "À tester" }, ["Lecture impossible : complète à la main."]);
  } finally { if (btn) { btn.disabled = false; btn.textContent = "Ajouter"; } }
}

/* ---------- Radar : les sites à surveiller ---------- */
const EMO_SRC = { "Salles de concert": "🎸", "Sites d’info": "📰", "Salles de spectacle": "🎭", Ciné: "🎬", Musées: "🏛️" };
// catégories par défaut d'abord, puis celles créées au fil de l'eau
function srcCats() {
  return [...new Set((S.L.srcCat || []).concat(S.sources.map(s => s.cat).filter(Boolean).sort((a, b) => a.localeCompare(b, "fr"))))];
}
const domaine = u => String(u).replace(/^https?:\/\/(www[.])?/i, "").replace(/\/$/, "");

function renderRadar() {
  const groupes = srcCats().map(c => ({ cat: c, items: S.sources.filter(s => s.cat === c) })).filter(g => g.items.length);
  return `<section class="hero mini"><div class="eyebrow">Radar · ${plur(S.sources.length, "site")} suivi${S.sources.length > 1 ? "s" : ""}</div>
      <h2>Quoi de neuf chez tes lieux&nbsp;?</h2></section>
    <div class="radar-bar">
      <button class="btn primary" data-act="prompt">${ICON.spark} Prompt IA</button>
      <button class="btn" data-act="new-src">+ Ajouter un site</button>
    </div>` +
    (groupes.length ? groupes.map(g => `<div class="card srcg">
      <div class="srch"><div class="emo" data-c="${esc(g.cat)}">${EMO_SRC[g.cat] || "📌"}</div>
        <h3>${esc(g.cat)}<small>${g.items.length}</small></h3>
        <button class="btn" data-act="open-cat" data-v="${esc(g.cat)}">Tout ouvrir</button></div>
      ${g.items.map(s => `<div class="src">
        <a class="srcl" href="${esc(s.lien)}" target="_blank" rel="noopener"><b>${esc(s.nom)}</b><span>${esc(s.notes || domaine(s.lien))}</span></a>
        <button class="ibtn sm" data-act="edit-src" data-id="${esc(s.id)}" title="Modifier" aria-label="Modifier">${ICON.edit}</button>
        <a class="ibtn sm go" href="${esc(s.lien)}" target="_blank" rel="noopener" title="Ouvrir" aria-label="Ouvrir ${esc(s.nom)}">${ICON.next}</a>
      </div>`).join("")}</div>`).join("")
    : `<div class="empty"><span class="big">📡</span>Aucun site pour l’instant.<br>Ajoute les salles, musées et sites que tu consultes souvent.</div>`);
}

// Ouvre tous les sites d'une catégorie. Si le navigateur bloque l'ouverture groupée,
// on propose de les ouvrir un par un (chaque appui compte comme un clic autorisé).
function openCat(cat) {
  const liste = S.sources.filter(s => s.cat === cat && s.lien);
  const bloques = [];
  liste.forEach(s => {
    const w = window.open(s.lien, "_blank");
    if (w) { try { w.opener = null; } catch (e) { /* autre origine */ } } else bloques.push(s);
  });
  if (!bloques.length) return toast(plur(liste.length, "onglet") + " ouvert" + (liste.length > 1 ? "s" : ""));
  form = null;
  $("#modal").innerHTML = `<div class="overlay" data-act="close-bg"><div class="sheet">
    <h2>${esc(cat)}</h2>
    <div class="infos">Ton navigateur a bloqué l’ouverture groupée. Touche chaque site ci-dessous, ou autorise les pop-ups pour cette appli.</div>
    ${bloques.map(s => `<a class="item" href="${esc(s.lien)}" target="_blank" rel="noopener" data-act="opened">
      <div class="txt">${esc(s.nom)}<div class="meta">${esc(domaine(s.lien))}</div></div>${ICON.next}</a>`).join("")}
    <div class="foot"><span></span><button class="btn" data-act="close">Fermer</button></div>
  </div></div>`;
}

/* ---------- prompt pour une IA ---------- */
const QUAND = [["soir", "Ce soir"], ["demain", "Demain"], ["we", "Ce week-end"], ["semaine", "Cette semaine"], ["mois", "Ce mois-ci"], ["3mois", "Les 3 prochains mois"]];
const QUOI = [["voir", "Ce que je peux voir ou faire"], ["inscr", "Les inscriptions et billetteries ouvertes"], ["deux", "Les deux"]];

function openPrompt() {
  form = null;
  const sel = (id, opts, v) => `<select id="${id}">${opts.map(([k, l]) => `<option value="${k}" ${k === v ? "selected" : ""}>${esc(l)}</option>`).join("")}</select>`;
  $("#modal").innerHTML = `<div class="overlay" data-act="close-bg"><div class="sheet" id="pform">
    <h2>Prompt IA</h2>
    <div class="field"><label>Sites à scanner</label><div class="checks">${srcCats().filter(c => S.sources.some(s => s.cat === c)).map(c =>
      `<label><input type="checkbox" name="pc" value="${esc(c)}" checked> ${esc(c)}</label>`).join("")}</div></div>
    <div class="grid2">
      <div class="field"><label for="p_quand">Quand</label>${sel("p_quand", QUAND, "soir")}</div>
      <div class="field"><label for="p_avec">Avec qui</label>${sel("p_avec", [["", "Peu importe"]].concat(S.L.avec.map(a => [a, a])), S.f.avec)}</div>
    </div>
    <div class="grid2">
      <div class="field"><label for="p_quoi">Je cherche</label>${sel("p_quoi", QUOI, "voir")}</div>
      <div class="field"><label for="p_budget">Budget</label>${sel("p_budget", [["", "Peu importe"], ["gratuit", "Gratuit seulement"], ["20", "Moins de 20 €"], ["50", "Moins de 50 €"]], "")}</div>
    </div>
    <div class="field"><label for="p_plus">Précisions</label><input type="text" id="p_plus" placeholder="ex. plutôt jazz, pas trop tard, près de Bastille…"></div>
    <div class="field"><label for="p_out">Ton prompt</label><textarea id="p_out" class="pout" readonly></textarea></div>
    <div class="foot"><button class="btn" data-act="close">Fermer</button><button class="btn primary" data-act="copy-prompt">Copier le prompt</button></div>
  </div></div>`;
  majPrompt();
}

function fmtLong(iso) { return new Date(iso + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }); }

function periode(k) {
  const t = today(), dow = new Date(t + "T12:00:00").getDay();
  const du = (a, b) => `du ${fmtLong(a)} au ${fmtLong(b)}`;
  if (k === "soir") return `ce soir, ${fmtLong(t)}`;
  if (k === "demain") return `demain, ${fmtLong(addDays(t, 1))}`;
  if (k === "we") {
    if (dow === 0) return `aujourd’hui, ${fmtLong(t)} (dimanche)`;
    const sam = dow === 6 ? t : addDays(t, 6 - dow);
    return `ce week-end, ${du(sam, addDays(sam, 1))}`;
  }
  if (k === "semaine") return `cette semaine, ${du(t, addDays(t, dow === 0 ? 0 : 7 - dow))}`;
  if (k === "mois") { const [y, m] = t.split("-").map(Number); return `ce mois-ci, ${du(t, `${y}-${String(m).padStart(2, "0")}-${String(new Date(y, m, 0).getDate()).padStart(2, "0")}`)}`; }
  return `dans les 3 prochains mois, ${du(t, addDays(t, 90))}`;
}

function buildPrompt() {
  const cats = [...document.querySelectorAll('input[name="pc"]:checked')].map(c => c.value);
  const v = id => $("#" + id).value;
  const sites = cats.map(c => `## ${c}\n` + S.sources.filter(s => s.cat === c).map(s => `- ${s.nom} : ${s.lien}${s.notes ? " (" + s.notes + ")" : ""}`).join("\n")).join("\n\n");
  const quoi = v("p_quoi");
  const avec = { Solo: "seul", Couple: "en couple", Amis: "entre amis" }[v("p_avec")];
  const budget = { gratuit: "uniquement des choses gratuites", 20: "moins de 20 € par personne", 50: "moins de 50 € par personne" }[v("p_budget")];
  const lignes = [
    "Tu m’aides à choisir mes sorties à Paris. Ouvre chacune des pages ci-dessous et lis leur programmation à jour (concerts, spectacles, films, expositions, événements…).",
    "", sites, "",
    "Ce que je veux savoir, pour " + periode(v("p_quand")) + (avec ? ", " + avec : "") + " :"
  ];
  if (quoi !== "inscr") lignes.push("- ce que je peux aller voir ou faire sur cette période ;");
  if (quoi !== "voir") lignes.push("- les événements dont les inscriptions ou la billetterie sont ouvertes en ce moment, ouvrent bientôt ou ferment bientôt (avec les dates).");
  if (budget) lignes.push("Budget : " + budget + ".");
  if (v("p_plus").trim()) lignes.push("Précisions : " + v("p_plus").trim());
  lignes.push("", "Pour chaque proposition, donne : le nom, le lieu, la date et l’heure, le prix, s’il reste des places, et le lien direct. Classe par date, et mets en premier ce qui est le plus immanquable.",
    "Si une page est inaccessible ou ne donne pas l’info, dis-le clairement au lieu de deviner.");
  return lignes.join("\n");
}
function majPrompt() { const o = $("#p_out"); if (o) o.value = buildPrompt(); }

async function copyPrompt() {
  const ta = $("#p_out");
  try { await navigator.clipboard.writeText(ta.value); }
  catch (e) { ta.removeAttribute("readonly"); ta.select(); document.execCommand("copy"); ta.setAttribute("readonly", ""); }
  toast("Prompt copié : colle-le dans ton IA");
}

/* ---------- bloc-notes ---------- */
let notesT = null, notesSale = false;
function renderNotes() {
  return `<div class="notesbar"><b>Notes</b><span class="meta" id="nstate">${notesSale ? "Modifié…" : "Enregistré automatiquement"}</span></div>
    <textarea class="notes" id="notes" placeholder="Une idée en vrac, un resto conseillé, un film à voir…">${esc(S.notes)}</textarea>`;
}
function noteChanged(v) {
  S.notes = v; notesSale = true;
  const st = $("#nstate"); if (st) st.textContent = "Modifié…";
  clearTimeout(notesT); notesT = setTimeout(saveNotesNow, 1200);
}
async function saveNotesNow() {
  clearTimeout(notesT);
  if (!notesSale) return;
  notesSale = false;
  try {
    await call("saveNotes", S.notes);
    const st = $("#nstate"); if (st && !notesSale) st.textContent = "Enregistré ✓";
  } catch (e) { notesSale = true; }
}

// anim = true : les cartes apparaissent en cascade (changement d'onglet, nouveau tirage…).
function render(anim) {
  if (!S.pret) { if (!uid && auth) renderConnexion(); return; }
  $("#sub").textContent = new Date().toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" });
  document.querySelectorAll("[data-tab]").forEach(b => b.classList.toggle("on", b.dataset.tab === S.tab));
  const main = $("#main");
  main.classList.remove("anim");
  fermerCarte();
  main.innerHTML = S.tab === "idees" ? renderIdees() : S.tab === "evts" ? renderEvts() : S.tab === "adresses" ? renderAdresses() :
    S.tab === "notes" ? renderNotes() : renderRadar();
  if (S.tab === "adresses" && S.aview === "Carte") drawMap();
  if (anim) {
    main.querySelectorAll(".card,.item").forEach((el, i) => { el.style.animationDelay = Math.min(i, 12) * 55 + "ms"; });
    void main.offsetWidth; // relance l'animation
    main.classList.add("anim");
  }
}

/* ---------- 5. Formulaires ---------- */
const F_IDEE = [
  ["nom", "Activité", "text"], ["cat", "Catégorie", "cat"], ["groupe", "Groupe (facultatif)", "groupe"], ["avec", "Avec qui", "checks", "avec"],
  ["duree", "Durée", "select", "duree", "half"], ["lieu", "Lieu", "select", "lieu", "half"],
  ["ie", "Intérieur / extérieur", "select", "ie", "half"], ["budget", "Prix (0 à 3)", "select", "budget", "half"],
  ["energie", "Énergie", "select", "energie", "half"],
  ["notes", "Notes", "textarea"]
];
const F_EVT = [
  ["titre", "Titre", "text"], ["lien", "Lien", "url"],
  ["type", "Type", "select", "type", "half"], ["zone", "Zone", "select", "zone", "half"],
  ["lieu", "Lieu", "text"],
  ["debut", "Début", "date", "", "half"], ["fin", "Fin", "date", "", "half"],
  ["prix", "Prix en €", "text", "", "half"], ["niveau", "Prix (0 à 3)", "select", "budget", "half"],
  ["statut", "Statut", "select", "statut", "half"], ["heure", "Heure (si réservé)", "time", "", "half"],
  ["avec", "Avec qui", "checks", "avec"], ["desc", "Description", "textarea"]
];
const F_SRC = [["nom", "Nom", "text"], ["lien", "Lien", "url"], ["cat", "Catégorie", "srccat"], ["notes", "Notes", "textarea"]];
const F_ADR = [
  ["nom", "Nom", "text"], ["cat", "Catégorie", "adrcat", "", "half"], ["style", "Style", "style", "", "half"],
  ["adresse", "Adresse", "text"], ["prix", "Prix (0 à 3)", "select", "budget", "half"], ["statut", "Statut", "select", "adrStatut", "half"],
  ["lien", "Lien (Google Maps ou site)", "url"], ["notes", "Notes", "textarea"]
];
// Ce qui change d'un type de fiche à l'autre (« supprime » : message après suppression)
const KINDS = {
  idee: { fields: F_IDEE, save: "saveIdee", del: "deleteIdee", arr: "idees", req: "nom", nouveau: "Nouvelle idée", vide: "Donne un nom à l’activité", supprime: "Idée supprimée" },
  evt: { fields: F_EVT, save: "saveEvt", del: "deleteEvt", arr: "evts", req: "titre", nouveau: "Nouvel événement", vide: "Donne un titre à l’événement", supprime: "Événement supprimé" },
  src: { fields: F_SRC, save: "saveSource", del: "deleteSource", arr: "sources", req: "nom", nouveau: "Nouveau site", vide: "Donne un nom au site", supprime: "Site supprimé" },
  adr: { fields: F_ADR, save: "saveAdresse", del: "deleteAdresse", arr: "adresses", req: "nom", nouveau: "Nouvelle adresse", vide: "Donne un nom à l’adresse", supprime: "Adresse supprimée" }
};
let form = null;

function fieldHtml([k, label, type, src], v) {
  const id = "f_" + k;
  let input;
  if (type === "select") input = `<select id="${id}"><option value=""></option>${S.L[src].map(o => `<option ${o === v ? "selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
  else if (type === "checks") input = `<div class="checks">${S.L[src].map(o => `<label><input type="checkbox" name="${id}" value="${esc(o)}" ${list(v).includes(o) ? "checked" : ""}> ${esc(o)}</label>`).join("")}</div>`;
  else if (type === "textarea") input = `<textarea id="${id}">${esc(v)}</textarea>`;
  else if (type === "cat" || type === "srccat") {
    // liste proposée + saisie libre : une nouvelle catégorie se crée en la tapant
    const cats = type === "cat" ? [...new Set(S.L.cat.concat(S.idees.map(i => i.cat).filter(Boolean)))] : srcCats();
    input = `<input type="text" id="${id}" list="cats" value="${esc(v)}" placeholder="Choisis ou tape une nouvelle catégorie"><datalist id="cats">${cats.map(c => `<option value="${esc(c)}">`).join("")}</datalist>`;
  } else if (type === "adrcat" || type === "style") {
    // catégorie d'adresse ou style (indien, bistrot, art…) : propositions + saisie libre
    const opts = type === "adrcat" ? adrCats() : [...new Set(S.adresses.map(a => a.style).filter(Boolean))].sort((a, b) => a.localeCompare(b, "fr"));
    input = `<input type="text" id="${id}" list="l_${type}" value="${esc(v)}" placeholder="${type === "style" ? "ex. indien, bistrot, art" : "Choisis ou tape"}"><datalist id="l_${type}">${opts.map(c => `<option value="${esc(c)}">`).join("")}</datalist>`;
  } else if (type === "groupe") {
    // sous-catégorie : au tirage, toutes les idées d'un groupe comptent pour une seule
    const gs = [...new Set(S.idees.map(grp).filter(Boolean))].sort((a, b) => a.localeCompare(b, "fr"));
    input = `<input type="text" id="${id}" list="grps" value="${esc(v)}" placeholder="ex. Balade dans Paris"><datalist id="grps">${gs.map(g => `<option value="${esc(g)}">`).join("")}</datalist>`;
  } else input = `<input type="${type}" id="${id}" value="${esc(v)}">`;
  return `<div class="field"><label for="${id}">${label}</label>${input}</div>`;
}

// coller : la page n'a pas pu être lue (ex. Shotgun), le bloc « coller le texte » s'ouvre directement.
function openForm(kind, obj, infos, coller) {
  const fields = KINDS[kind].fields;
  form = { kind, fields, base: Object.assign({}, obj) };
  let html = "", half = [];
  const flush = () => { if (half.length) { html += `<div class="grid2">${half.join("")}</div>`; half = []; } };
  fields.forEach(f => { const h = fieldHtml(f, obj[f[0]]); if (f[4] === "half") half.push(h); else { flush(); html += h; } });
  flush();
  const titre = obj.id ? "Modifier" : KINDS[kind].nouveau;
  const notes = (infos || []).filter(Boolean);
  $("#modal").innerHTML = `<div class="overlay" data-act="close-bg"><div class="sheet">
    <h2>${titre}</h2>
    ${notes.length ? `<div class="infos">${notes.map(esc).join("<br>")}</div>` : ""}
    ${kind === "evt" ? `<details class="paste" ${coller ? "open" : ""}><summary>Compléter en collant le texte de la page</summary>
      <textarea id="f_texte" placeholder="Colle ici le texte de la page de l’événement"></textarea>
      <button class="btn sm" data-act="parse-texte">Compléter la fiche</button>
      <div class="meta">Sur le téléphone : appui long sur la page → Tout sélectionner → Copier. Seules les cases vides sont remplies.</div></details>` : ""}
    ${html}
    <div class="foot">
      ${obj.id ? `<button class="btn danger" data-act="delete">Supprimer</button>` : "<span></span>"}
      <span style="display:flex;gap:8px"><button class="btn" data-act="close">Annuler</button><button class="btn primary" data-act="save">Enregistrer</button></span>
    </div></div></div>`;
}

function readForm() {
  const o = Object.assign({}, form.base);
  form.fields.forEach(([k, , type]) => {
    if (type === "checks") o[k] = [...document.querySelectorAll(`input[name="f_${k}"]:checked`)].map(c => c.value).join(", ");
    else o[k] = $("#f_" + k).value.trim();
  });
  return o;
}

function closeForm() { $("#modal").innerHTML = ""; form = null; popup = null; }

/* ---------- mini pop-up de présentation (toucher une carte ou une ligne) ---------- */
// Les actions courantes en gros boutons ; Modifier et Supprimer rangés dans le menu « ⋯ ».
let popup = null; // { kind, id, menu }
function ouvrirPopup(kind, id) { popup = { kind, id, menu: false }; afficherPopup(); }
const btnPop = (act, ic, label, cls) => `<button class="pb ${cls || ""}" data-act="${act}"><i>${ic}</i>${label}</button>`;
const lienPop = (url, ic, label, cls) => `<a class="pb ${cls || ""}" href="${esc(url)}" target="_blank" rel="noopener"><i>${ic}</i>${label}</a>`;

function afficherPopup() {
  if (!popup) return;
  const K = KINDS[popup.kind], o = S[K.arr].find(x => x.id === popup.id);
  if (!o) return closeForm();
  form = null;
  let vignette, kicker = "", etiquettes = "", infos = [], notes = "", btns = [];
  if (popup.kind === "evt") {
    const e = Object.assign({ st: evState(o) }, o);
    vignette = `<div class="pimg" data-t="${esc(e.type)}"><span>${typeEmo(e.type)}</span>${imageEvt(e)}</div>`;
    kicker = e.type || "";
    etiquettes = badges(e);
    const [ic, quand] = dateLigne(e);
    infos.push(`<div class="quand">${ic}<span>${esc(quand)}</span></div>`);
    const p = progression(e);
    if (p !== null) infos.push(`<div class="prog"><i style="width:${p}%"></i></div>`);
    if (lieuPrix(e)) infos.push(`<div>${esc(lieuPrix(e))}</div>`);
    if (e.lien) btns.push(lienPop(e.lien, ICON.ouvrir, "Ouvrir", "p"));
    btns.push(btnPop("pop-resa", ICON.billet, reserve(e) ? "Réservé ✓" : "Réservé", reserve(e) ? "r on" : "r"));
    btns.push(actif(e) ? btnPop("pop-fait", ICON.check, "Fait") : btnPop("pop-avoir", ICON.prev, "Remettre"));
  } else if (popup.kind === "idee") {
    vignette = `<div class="pimg" data-c="${esc(o.cat)}"><span>${catEmo(o.cat)}</span></div>${catIllu(o.cat)}`;
    kicker = [o.cat, grp(o)].filter(Boolean).join(" · ");
    if (o.derniere) infos.push(`<div>Dernière fois : ${esc(fmt(o.derniere))}</div>`);
    notes = o.notes || "";
    btns.push(btnPop("pop-pin", ICON.pin, epingle(o) ? "Épinglée" : "Épingler", epingle(o) ? "on" : ""));
    btns.push(btnPop("pop-faitidee", ICON.check, o.derniere === today() ? "Fait aujourd’hui ✓" : "Fait aujourd’hui", o.derniere === today() ? "r on" : ""));
  } else {
    vignette = `<div class="pimg" data-c="${esc(o.cat)}"><span>${adrEmo(o.cat)}</span></div>`;
    kicker = [o.cat, o.style].filter(Boolean).join(" · ");
    if (o.statut === "Adoré") etiquettes = '<span class="badge adore">♥ Adoré</span>';
    const ligne = [o.adresse || o.ville, o.prix].filter(Boolean).join(" · ");
    if (ligne) infos.push(`<div>${esc(ligne)}</div>`);
    notes = o.notes || "";
    btns.push(lienPop(mapsUrl(o), ICON.map, "Maps", "p"));
    btns.push(lienPop(itineraire(o), ICON.iti, "Itinéraire"));
    btns.push(btnPop("pop-pin", ICON.pin, epingle(o) ? "Épinglée" : "Épingler", epingle(o) ? "on" : ""));
  }
  const titre = o.titre || o.nom || "(sans titre)";
  const menu = `<span class="pmenu"><button class="pmini ${popup.menu ? "on" : ""}" data-act="pop-menu" aria-label="Plus d’actions" aria-expanded="${popup.menu}">${ICON.dots}</button>
    ${popup.menu ? `<div class="pdd"><button data-act="pop-edit">${ICON.edit}Modifier</button><button class="sup" data-act="pop-del">${ICON.trash}Supprimer</button></div>` : ""}</span>`;
  $("#modal").innerHTML = `<div class="overlay centre" data-act="close-bg"><div class="pop" role="dialog" aria-label="${esc(titre)}">
    <div class="ptop">${vignette}
      <div class="ptx">${kicker ? `<div class="kicker">${esc(kicker)}</div>` : ""}<h3>${esc(titre)}</h3>${etiquettes ? `<div class="pbds">${etiquettes}</div>` : ""}</div>
      <div class="pcoin">${menu}<button class="pmini" data-act="close" aria-label="Fermer">${ICON.x}</button></div></div>
    ${infos.length ? `<div class="pinfo">${infos.join("")}</div>` : ""}
    ${notes ? `<div class="pnotes">${esc(notes)}</div>` : ""}
    <div class="pbtns">${btns.join("")}</div>
  </div></div>`;
}

// Change le statut d'un événement depuis le pop-up
async function statutPop(statut, message, fermer) {
  const ev = S.evts.find(x => x.id === popup.id);
  const saved = await call("saveEvt", Object.assign({}, ev, { statut }));
  Object.assign(ev, saved); render(); toast(message);
  if (fermer) closeForm(); else afficherPopup();
}

async function saveForm(btn) {
  const o = readForm();
  const K = KINDS[form.kind];
  if (!o[K.req]) return toast(K.vide);
  if (form.kind === "evt" && !o.niveau) o.niveau = niveauPrix(o.prix);
  // adresse modifiée : on recalcule la position et la ville
  if (form.kind === "adr" && o.adresse !== String(form.base.adresse || "")) { o.lat = ""; o.lng = ""; o.ville = ""; }
  if (form.kind === "src") {
    if (!o.lien) return toast("Ajoute le lien du site");
    if (!/^https?:\/\//i.test(o.lien)) o.lien = "https://" + o.lien;
    if (!o.cat) o.cat = "Autres";
  }
  btn.disabled = true; btn.textContent = "…";
  try {
    const saved = await call(K.save, o);
    const arr = S[K.arr];
    const i = arr.findIndex(x => x.id === saved.id);
    if (i >= 0) arr[i] = saved; else arr.push(saved);
    closeForm(); render(); toast("Enregistré");
  } catch (e) { btn.disabled = false; btn.textContent = "Enregistrer"; }
}

// Suppression immédiate, avec « Annuler » pendant quelques secondes.
async function deleteForm(btn) {
  btn.disabled = true;
  await supprimer(form.kind, form.base.id);
}
async function supprimer(kind, id) {
  const K = KINDS[kind];
  const avant = S[K.arr].find(x => x.id === id);
  await call(K.del, id);
  S[K.arr] = S[K.arr].filter(x => x.id !== id);
  if (kind === "idee") S.picks = S.picks.filter(x => x !== id);
  closeForm(); render();
  toast(K.supprime, avant && (() => {
    ecrire(K.arr, avant);
    S[K.arr] = S[K.arr].filter(x => x.id !== id).concat([avant]).sort(parCree);
    render(); toast("Restauré");
  }));
}

async function extract(text) {
  // premier mot qui commence par http (le texte partagé peut contenir un titre avant le lien)
  const mots = String(text || "").replace(/[\n\t]/g, " ").split(" ");
  const url = mots.find(w => /^https?:\/\/./i.test(w));
  if (!url) return toast("Colle un lien qui commence par http");
  const dup = S.evts.find(e => e.lien === url);
  if (dup) { toast("Déjà dans ta liste"); return openForm("evt", dup); }
  const btn = $("#go");
  if (btn) { btn.disabled = true; btn.textContent = "Lecture…"; }
  // texte partagé avec le lien (partage Android) : il peut contenir la date, le lieu ou le prix
  const reste = mots.filter(w => w !== url).join(" ").trim();
  try {
    const out = await call("extractEvent", url);
    const infos = out.infos, bloque = out.bloque; delete out.infos; delete out.bloque;
    if (reste.length > 15) {
      try { remplirVides(out, await call("parseEventText", reste, out.type, out.titre)); } catch (e) { /* le lien suffit */ }
    }
    sansFinInventee(out);
    openForm("evt", Object.assign({ statut: "À voir", avec: "" }, out), (infos || []).filter(i => !/^Film : fin/.test(i)), bloque);
  } catch (e) {
    openForm("evt", { lien: url, statut: "À voir" }, ["Lecture impossible : colle le texte de la page ci-dessous, ou complète à la main."], true);
  } finally { if (btn) { btn.disabled = false; btn.textContent = "Ajouter"; } }
}

// Ne remplit que les cases vides ; renvoie les noms de ce qui a été ajouté.
const NOMS_CHAMPS = { debut: "date", fin: "fin", lieu: "lieu", zone: "zone", prix: "prix" };
// Le titre, lui, est remplacé quand le texte en donne une version plus complète.
function remplirVides(o, r) {
  const ajout = Object.keys(NOMS_CHAMPS).filter(k => r[k] && !o[k]);
  ajout.forEach(k => { o[k] = r[k]; });
  if (ajout.includes("prix") && !o.niveau) o.niveau = niveauPrix(o.prix);
  const noms = ajout.map(k => NOMS_CHAMPS[k]);
  if (r.titre) { o.titre = r.titre; noms.unshift("titre"); }
  return noms;
}

async function completerTexte(btn) {
  const texte = $("#f_texte").value.trim();
  if (!texte) return toast("Colle d’abord le texte de la page");
  const o = readForm();
  btn.disabled = true; btn.textContent = "Lecture…";
  try {
    const ajout = remplirVides(o, sansFinInventee(Object.assign({ type: o.type }, await call("parseEventText", texte, o.type, o.titre))));
    openForm("evt", o, [ajout.length ? "Complété depuis le texte : " + ajout.join(", ") + ". Vérifie avant d’enregistrer." : "Rien de plus trouvé dans ce texte."]);
  } catch (e) { btn.disabled = false; btn.textContent = "Compléter la fiche"; }
}

/* ---------- 6. Apparence ---------- */
const THEMES = [
  { id: "auto", nom: "Automatique", desc: "Suisse le jour, Suisse nuit le soir (selon l’appareil)", sw: ["#f4f3f0", "#ec5a2e", "#141518", "#ff7a4d"] },
  { id: "suisse", nom: "Suisse", desc: "Papier et orange, épuré", sw: ["#f4f3f0", "#ffffff", "#161616", "#ec5a2e"] },
  { id: "suissenuit", nom: "Suisse nuit", desc: "Nuit et corail", sw: ["#141518", "#1e2024", "#ededea", "#ff7a4d"] },
  { id: "durevie", nom: "Dure Vie", desc: "Encre, pervenche et orange", sw: ["#f4f3f7", "#b3c6f7", "#171520", "#ff641f"] },
  { id: "pastel", nom: "Pastel", desc: "Doux et arrondi, comme avant", sw: ["#fbf6f1", "#fde4dc", "#fff1d6", "#ec8a72"] },
  { id: "sombre", nom: "Sombre", desc: "Dure Vie de nuit", sw: ["#121019", "#283463", "#f1eff6", "#ff641f"] }
];
const SOMBRE = window.matchMedia("(prefers-color-scheme: dark)");
function choixTheme() { try { return localStorage.getItem("qf-theme") || "auto"; } catch (e) { return "auto"; } }
function appliquerTheme() {
  const c = choixTheme();
  document.documentElement.dataset.theme = c === "auto" ? (SOMBRE.matches ? "suissenuit" : "suisse") : c;
}
// le téléphone passe en mode sombre (le soir, par exemple) : on suit si « Automatique »
SOMBRE.addEventListener("change", appliquerTheme);
function setTheme(id) {
  try { localStorage.setItem("qf-theme", id); } catch (e) { /* stockage indisponible : le thème vaut pour cette fois */ }
  appliquerTheme();
  document.querySelectorAll(".theme").forEach(b => b.classList.toggle("on", b.dataset.v === id));
}

/* ---------- Excel : export / import ---------- */
let importEnAttente = null;

// Une ligne de choix dans Paramètres (vue de départ d'un onglet)
const ligneReglage = (label, k, opts) => `<div class="row"><span class="lbl">${label}</span><div class="chips">${opts.map(([v, l]) =>
  `<button class="chip ${S.reglages[k] === v ? "on" : ""}" data-act="reg" data-k="${k}" data-v="${v}">${l}</button>`).join("")}</div></div>`;

// Choix du thème : un écran à part, ouvert depuis Paramètres
function openThemes() {
  form = null;
  const cur = choixTheme();
  $("#modal").innerHTML = `<div class="overlay" data-act="close-bg"><div class="sheet">
    <button class="link" data-act="settings">‹ Paramètres</button>
    <h2>Thème</h2>
    <div class="themes">${THEMES.map(t => `<button class="theme ${t.id === cur ? "on" : ""}" data-act="set-theme" data-v="${t.id}">
      <span class="swatch">${t.sw.map(c => `<i style="background:${c}"></i>`).join("")}</span>
      <span><b>${t.nom}</b><small>${t.desc}</small></span></button>`).join("")}</div>
    <div class="foot"><span></span><button class="btn" data-act="close">Fermer</button></div>
  </div></div>`;
}

// Paramètres : thème, vues de départ, Excel (export / import), compte et version
function openSettings() {
  form = null; importEnAttente = null;
  const t = THEMES.find(x => x.id === choixTheme()) || THEMES[0];
  $("#modal").innerHTML = `<div class="overlay" data-act="close-bg"><div class="sheet">
    <h2>Paramètres</h2>
    <div class="lbl">Apparence</div>
    <button class="setrow card" data-act="themes"><span>Thème</span><span class="meta">${esc(t.nom)} ›</span></button>
    ${uid ? `<div class="lbl" style="margin-top:22px">Affichage au démarrage</div>
    <div class="card">${ligneReglage("Idées", "vueIdees", [["tirage", "Tirage"], ["liste", "Liste"]])}
      ${ligneReglage("À venir", "vueEvts", [["Cartes", "Cartes"], ["Liste", "Liste"], ["Calendrier", "Calendrier"]])}
      <div class="meta">Les mêmes sur le téléphone et le PC.</div></div>
    <div class="lbl" style="margin-top:22px">Excel</div>
    <div class="card excel"><button class="btn" data-act="export">Exporter</button>
      <label class="btn">Importer<input type="file" id="file" accept=".xlsx" hidden></label>
      <div id="imp"></div></div>
    <div class="lbl" style="margin-top:22px">Compte</div>
    <div class="card"><div class="compte"><span class="meta">Connecté : ${esc(S.compte)}</span><button class="btn sm" data-act="logout">Se déconnecter</button></div>
      <div class="compte" style="margin-top:6px"><span class="meta">Synchro</span><span id="sync" class="sync" hidden></span></div>
      <div class="meta" style="margin-top:6px">Lecture des liens : ${LECTEUR_URL ? "branchée" : "pas encore branchée (remplissage à la main)"}</div></div>` : ""}
    <div class="version">Version ${VERSION.replace("v", "")}</div>
    <div class="foot"><span></span><button class="btn" data-act="close">Fermer</button></div>
  </div></div>`;
  majSynchro(); // l'indicateur de synchro vit maintenant ici
}

// Largeur des colonnes Excel (en caractères) ; 14 par défaut
const LARGEURS = { nom: 34, titre: 34, desc: 50, notes: 40, lien: 40, image: 30, adresse: 34, lieu: 26, avec: 18, cat: 16, style: 16, ville: 14, groupe: 20 };

// Date « 2026-10-20 » -> vraie date Excel (affichée 20/10/2026)
function versExcel(k, v) {
  if (DATE_KEYS.includes(k) && /^\d{4}-\d{2}-\d{2}$/.test(v || "")) {
    const [y, m, d] = v.split("-").map(Number);
    return new Date(Date.UTC(y, m - 1, d));
  }
  return v === undefined || v === null ? "" : v;
}

async function exporter(btn) {
  btn.disabled = true; btn.textContent = "Préparation…";
  try {
    const ExcelJS = await chargerExcelJS();
    const wb = new ExcelJS.Workbook();
    TABLES.forEach(T => {
      const cols = T.headers.map((h, i) => ({ header: h, key: T.keys[i], width: LARGEURS[T.keys[i]] || 14 }));
      const lignes = S[T.arr].map(o => { const l = {}; T.keys.forEach(k => { l[k] = versExcel(k, o[k]); }); return l; });
      const ws = ajouterFeuille(wb, T.nom, cols, lignes);
      T.keys.forEach((k, i) => { if (DATE_KEYS.includes(k)) ws.getColumn(i + 1).eachCell((c, n) => { if (n > 1) c.numFmt = "dd/mm/yyyy"; }); });
    });
    const wn = ajouterFeuille(wb, "Notes", [{ header: "Notes", key: "notes", width: 100 }], [{ notes: S.notes }]);
    wn.getCell("A2").alignment = { wrapText: true, vertical: "top" };
    await telechargerClasseur(wb, NOM_APPLI);
    toast("Fichier téléchargé");
  } catch (e) {
    toast("Export impossible : " + (e && e.message || e));
  } finally { btn.disabled = false; btn.textContent = "Exporter"; }
}

// Valeur brute d'une case Excel : texte, nombre, ou date « AAAA-MM-JJ »
function valeurCase(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return isNaN(v) ? "" : v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.hyperlink) return typeof v.text === "string" ? v.text : valeurCase(v.text) || v.hyperlink;
    if ("result" in v) return valeurCase(v.result);
    return "";
  }
  return v;
}
// Une feuille -> [{ en-tête: valeur }] (les lignes entièrement vides sont ignorées)
function lignesFeuille(ws) {
  const entetes = [];
  ws.getRow(1).eachCell((c, i) => { entetes[i] = String(valeurCase(c.value)).trim(); });
  const lignes = [];
  ws.eachRow((rang, n) => {
    if (n === 1) return;
    const o = {};
    let plein = false;
    entetes.forEach((h, i) => { if (!h) return; const v = valeurCase(rang.getCell(i).value); o[h] = v; if (v !== "") plein = true; });
    if (plein) lignes.push(o);
  });
  return lignes;
}
// Feuille Notes : tout le texte de la colonne A (sans l'en-tête « Notes » de l'export)
function texteNotes(ws) {
  const morceaux = [];
  ws.eachRow((rang, n) => {
    const v = String(valeurCase(rang.getCell(1).value)).trim();
    if (v && !(n === 1 && v.toLowerCase() === "notes")) morceaux.push(v);
  });
  return morceaux.join("\n");
}

const sansAccents = s => s.normalize("NFD").replace(/[̀-ͯ]/g, "");
const norm = s => sansAccents(String(s || "").toLowerCase()).replace(/[^a-z0-9]/g, "");
const nomCle = s => sansAccents(String(s || "").toLowerCase()).replace(/\s+/g, " ").trim();

function cibleImport(s) {
  const h = Object.keys(s.rows[0] || {}).map(norm);
  const n = norm(s.name);
  if (/idee/.test(n) || h.includes("activite")) return IDEES;
  if (/evenement/.test(n) || h.includes("titre")) return EVTS;
  if (/adresse/.test(n) || h.includes("style")) return ADRESSES;
  if (/source/.test(n) || (h.includes("nom") && h.includes("lien"))) return SOURCES;
  return null;
}

// Accepte 2026-10-20, 20/10/2026 ou un numéro de série Excel.
function dateIso(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  m = s.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{2,4})$/);
  if (!m) return "";
  const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
  return y + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[1]).padStart(2, "0");
}

function rowToObj(T, row) {
  const map = {};
  T.headers.forEach((h, i) => { map[norm(h)] = T.keys[i]; });
  const o = {};
  Object.keys(row).forEach(col => {
    const k = map[norm(col)];
    if (!k) return;
    let v = row[col];
    if (DATE_KEYS.includes(k)) v = dateIso(v);
    else if (typeof v === "boolean") v = v ? "oui" : "";
    else if (typeof v === "string") v = v.trim();
    o[k] = v;
  });
  // Ancien fichier Activités.xlsx : colonnes Solo / A plusieurs (0 ou 1)
  if (T === IDEES && !("avec" in o)) {
    let solo, plus;
    Object.keys(row).forEach(col => {
      if (norm(col) === "solo") solo = row[col];
      if (norm(col) === "aplusieurs") plus = row[col];
    });
    const avec = [];
    if (+solo === 1) avec.push("Solo");
    if (+plus === 1) avec.push("Couple", "Amis");
    if (avec.length) o.avec = avec.join(", ");
  }
  if (o.id !== undefined) o.id = idPropre(o.id);
  return o[T.keys[1]] ? o : null;
}

// Fusionne les lignes d'un onglet avec ce qui existe déjà.
// Même ID (ou même nom) : mise à jour, une case vide n'efface rien. Sinon : ajout.
// simuler = true : on compte sans rien changer.
function fusion(T, rows, simuler, aEcrire) {
  const data = S[T.arr].map(o => Object.assign({}, o));
  const parId = {}, parNom = {};
  data.forEach((r, i) => { parId[r.id] = i; parNom[nomCle(r[T.keys[1]])] = i; });
  const touches = new Set();
  let ajout = 0, maj = 0, cree = Date.now();
  rows.forEach(row => {
    const o = rowToObj(T, row);
    if (!o) return;
    const i = o.id && parId[o.id] !== undefined ? parId[o.id] : parNom[nomCle(o[T.keys[1]])];
    if (i !== undefined) {
      T.keys.forEach(k => { if (k !== "id" && k in o && o[k] !== "") data[i][k] = o[k]; });
      touches.add(i);
      maj++;
    } else {
      if (T === EVTS) {
        if (!o.statut) o.statut = "À voir";
        if (!o.ajoute) o.ajoute = today();
        if (!o.niveau) o.niveau = niveauPrix(o.prix);
      }
      if (T === ADRESSES) {
        if (!o.statut) o.statut = "À tester";
        if (!o.ajoute) o.ajoute = today();
        if (!o.ville) o.ville = villeDe(o.adresse);
      }
      o.id = o.id || newId();
      const full = {};
      T.keys.forEach(k => { full[k] = o[k] === undefined || o[k] === null ? "" : o[k]; });
      full.cree = cree++;
      data.push(full);
      parId[full.id] = parNom[nomCle(full[T.keys[1]])] = data.length - 1;
      touches.add(data.length - 1);
      ajout++;
    }
  });
  if (!simuler) {
    touches.forEach(i => aEcrire.push([T.arr, data[i]]));
    S[T.arr] = data;
  }
  return { ajout, maj };
}

function importer(sheets, simuler) {
  const res = { idees: { ajout: 0, maj: 0 }, evts: { ajout: 0, maj: 0 }, sources: { ajout: 0, maj: 0 }, adresses: { ajout: 0, maj: 0 }, ignores: [], notes: "" };
  const aEcrire = [];
  sheets.forEach(s => {
    if (norm(s.name) === "notes") { res.notes = s.notes; return; }
    const T = s.rows.length ? cibleImport(s) : null;
    if (!T) { if (s.rows.length) res.ignores.push(s.name); return; }
    const r = fusion(T, s.rows, simuler, aEcrire);
    res[T.arr].ajout += r.ajout;
    res[T.arr].maj += r.maj;
  });
  // Notes : on ajoute le texte du fichier s'il n'est pas déjà dans mes notes
  const nouvellesNotes = res.notes && !S.notes.includes(res.notes);
  if (!simuler) {
    if (nouvellesNotes) { S.notes = S.notes ? S.notes + "\n\n" + res.notes : res.notes; call("saveNotes", S.notes); }
    // envoi par paquets (Firestore accepte 500 écritures par paquet)
    if (MODE_DEMO) sauverDemo();
    else for (let i = 0; i < aEcrire.length; i += 400) {
      const lot = writeBatch(db);
      aEcrire.slice(i, i + 400).forEach(([arr, o]) => lot.set(doc(coll(arr), String(o.id)), o));
      lot.commit().catch(erreurEcriture);
    }
  }
  res.nouvellesNotes = !!nouvellesNotes;
  return res;
}

async function importFile(file) {
  const box = $("#imp");
  box.innerHTML = '<div class="infos">Lecture du fichier…</div>';
  try {
    const wb = await lireClasseur(file);
    const sheets = wb.worksheets.map(ws => norm(ws.name) === "notes"
      ? { name: ws.name, rows: [], notes: texteNotes(ws) }
      : { name: ws.name, rows: lignesFeuille(ws) });
    const r = importer(sheets, true);
    const total = ["idees", "evts", "sources", "adresses"].reduce((n, k) => n + r[k].ajout + r[k].maj, 0);
    if (!total && !r.nouvellesNotes) {
      box.innerHTML = '<div class="infos">Rien à importer : aucun onglet ne ressemble à Idées, Événements, Adresses ou Sources.</div>';
      return;
    }
    importEnAttente = sheets;
    box.innerHTML = `<div class="infos">
        <b>Idées</b> : ${plur(r.idees.ajout, "nouvelle")}, ${r.idees.maj} mise${r.idees.maj > 1 ? "s" : ""} à jour<br>
        <b>Événements</b> : ${r.evts.ajout} nouveau${r.evts.ajout > 1 ? "x" : ""}, ${r.evts.maj} mis à jour<br>
        <b>Radar</b> : ${r.sources.ajout} nouveau${r.sources.ajout > 1 ? "x" : ""} site${r.sources.ajout > 1 ? "s" : ""}, ${r.sources.maj} mis à jour<br>
        <b>Adresses</b> : ${plur(r.adresses.ajout, "nouvelle")}, ${r.adresses.maj} mise${r.adresses.maj > 1 ? "s" : ""} à jour
        ${r.nouvellesNotes ? "<br><b>Notes</b> : ajoutées à la suite des tiennes" : ""}
        ${r.ignores.length ? "<br>Onglets ignorés : " + r.ignores.map(esc).join(", ") : ""}</div>
      <div class="actions"><button class="btn primary" data-act="imp-ok">Importer</button></div>`;
  } catch (e) {
    box.innerHTML = '<div class="infos">Impossible de lire ce fichier (' + esc(e && e.message || e) + ").</div>";
  }
}

function confirmImport(btn) {
  if (!importEnAttente) return;
  btn.disabled = true; btn.textContent = "Import…";
  importer(importEnAttente, false);
  importEnAttente = null;
  suggest(); closeForm(); render(); toast("Import terminé");
}

/* ---------- 7. Clics et démarrage ---------- */
document.addEventListener("click", async e => {
  const tab = e.target.closest("[data-tab]");
  if (tab) { if (!S.pret) return; saveNotesNow(); S.tab = tab.dataset.tab; render(true); window.scrollTo(0, 0); return; }
  const el = e.target.closest("[data-act]");
  if (!el) return;
  const act = el.dataset.act, id = el.dataset.id;
  if (act === "close-bg" && e.target !== el) return;
  switch (act) {
    case "login": seConnecter(el); break;
    case "logout": if (MODE_DEMO) break; closeForm(); S.pret = false; signOut(auth); break;
    case "export": exporter(el); break;
    case "f": S.f[el.dataset.k] = el.dataset.v; suggest(); render(true); break;
    case "ef": S.ef[el.dataset.k] = el.dataset.v; render(true); break;
    case "suggest": suggest(); render(true); break;
    case "mode": S.mode = el.dataset.v; render(true); break;
    case "grp": S.open[el.dataset.v] = !S.open[el.dataset.v]; $("#items").innerHTML = renderItems(); break;
    case "later": S.later = !S.later; render(); break;
    case "af": S.af[el.dataset.k] = el.dataset.v; render(); break;
    case "aview": S.aview = el.dataset.v; render(true); break;
    case "adr-extract": extractAdr($("#aurl").value); break;
    case "new-adr": openForm("adr", { statut: "À tester" }); break;
    case "edit-adr": openForm("adr", S.adresses.find(x => x.id === id)); break;
    // mini pop-up : ouvrir, menu ⋯, et ses actions
    case "ev-pop": ouvrirPopup("evt", id); break;
    case "idee-pop": ouvrirPopup("idee", id); break;
    case "adr-pop": ouvrirPopup("adr", id); break;
    case "pop-menu": popup.menu = !popup.menu; afficherPopup(); break;
    case "pop-edit": { const { kind, id: pid } = popup; popup = null; openForm(kind, S[KINDS[kind].arr].find(x => x.id === pid)); break; }
    case "pop-del": { const { kind, id: pid } = popup; closeForm(); supprimer(kind, pid); break; }
    case "pop-resa": statutPop(reserve(S.evts.find(x => x.id === popup.id)) ? "À voir" : "Réservé", "Statut mis à jour"); break;
    case "pop-fait": statutPop("Fait", "Noté comme fait", true); break;
    case "pop-avoir": statutPop("À voir", "Remis dans les à venir"); break;
    case "pop-pin": {
      const p = popup;
      await (p.kind === "idee" ? togglePin(p.id) : togglePinAdr(p.id));
      popup = p; afficherPopup(); break;
    }
    case "pop-faitidee": {
      const i = S.idees.find(x => x.id === popup.id);
      i.derniere = i.derniere === today() ? "" : today();
      await call("saveIdee", i);
      render(); toast(i.derniere ? "Noté : fait aujourd’hui" : "Date retirée"); afficherPopup(); break;
    }
    case "apin": togglePinAdr(id); break;
    case "locate": locate(el); break;
    case "parse-texte": completerTexte(el); break;
    case "pin": togglePin(id); break;
    case "edit-idee": openForm("idee", S.idees.find(x => x.id === id)); break;
    case "new-idee": openForm("idee", { avec: S.f.avec }); break;
    case "edit-evt": openForm("evt", S.evts.find(x => x.id === id)); break;
    case "new-evt": openForm("evt", { statut: "À voir" }); break;
    case "ev-done": {
      const ev = S.evts.find(x => x.id === id);
      el.disabled = true;
      const saved = await call("saveEvt", Object.assign({}, ev, { statut: "Fait" }));
      Object.assign(ev, saved); render(); toast("Noté comme fait"); break;
    }
    case "ev-resa": {
      // bascule « À voir » <-> « Réservé » ; l'heure se met dans la fiche
      const ev = S.evts.find(x => x.id === id);
      const saved = await call("saveEvt", Object.assign({}, ev, { statut: reserve(ev) ? "À voir" : "Réservé" }));
      Object.assign(ev, saved); render();
      toast(reserve(ev) ? "Réservé : ajoute l’heure dans la fiche si tu veux" : "Plus marqué comme réservé"); break;
    }
    case "reg": {
      // vue de départ : enregistrée, et appliquée tout de suite
      S.reglages[el.dataset.k] = el.dataset.v;
      if (el.dataset.k === "vueIdees") S.mode = el.dataset.v; else S.view = el.dataset.v;
      el.parentElement.querySelectorAll(".chip").forEach(c => c.classList.toggle("on", c === el));
      sauverReglages(); render(); break;
    }
    case "calmode": setCalMode(el.dataset.v); render(); break;
    case "fopen": S.fopen[el.dataset.v] = !S.fopen[el.dataset.v]; render(); break;
    case "add-open": {
      S.addOpen[el.dataset.v] = !S.addOpen[el.dataset.v]; render();
      const champ = $(el.dataset.v === "evts" ? "#url" : "#aurl");
      if (champ) champ.focus();
      break;
    }
    case "extract": extract($("#url").value); break;
    case "archive": S.archive = !S.archive; render(); break;
    case "view": S.view = el.dataset.v; render(true); break;
    case "settings": openSettings(); break;
    case "themes": openThemes(); break;
    case "new-src": openForm("src", { cat: "" }); break;
    case "edit-src": openForm("src", S.sources.find(x => x.id === id)); break;
    case "open-cat": openCat(el.dataset.v); break;
    case "opened": el.style.opacity = ".45"; break;
    case "prompt": openPrompt(); break;
    case "copy-prompt": copyPrompt(); break;
    case "set-theme": setTheme(el.dataset.v); break;
    case "imp-ok": confirmImport(el); break;
    case "cal-prev": shiftCal(-1); render(); break;
    case "cal-next": shiftCal(1); render(); break;
    case "day": S.calDay = el.dataset.v; render(); break;
    case "lf": S.lf[el.dataset.k] = el.dataset.v; render(); break;
    case "lf-reset": S.lf = Object.assign({}, LF_DEFAUT); S.tri = "cat"; render(); break;
    case "tri": S.tri = el.dataset.v; render(); break;
    case "filt": S.filt = !S.filt; render(); break;
    case "save": saveForm(el); break;
    case "delete": deleteForm(el); break;
    case "close": case "close-bg": closeForm(); break;
  }
});
document.addEventListener("change", e => {
  if (e.target.id === "file" && e.target.files[0]) importFile(e.target.files[0]);
  if (e.target.closest("#pform")) majPrompt();
});
document.addEventListener("input", e => {
  if (e.target.closest("#pform") && e.target.id !== "p_out") majPrompt();
  if (e.target.id === "q") { S.q = e.target.value; $("#items").innerHTML = renderItems(); }
  if (e.target.id === "notes") noteChanged(e.target.value);
});
// Quitter l'appli (ou passer en arrière-plan sur le téléphone) enregistre les notes tout de suite.
document.addEventListener("visibilitychange", () => { if (document.hidden) saveNotesNow(); });
document.addEventListener("keydown", e => {
  if (e.key === "Enter" && e.target.id === "url") extract(e.target.value);
  if (e.key === "Enter" && e.target.id === "aurl") extractAdr(e.target.value);
  if (e.key === "Escape" && (form || popup)) closeForm();
});

// Le « gardien » hors ligne (service worker) : garde une copie des fichiers de l'appli,
// et installe tout seul les nouvelles versions.
if ("serviceWorker" in navigator && !MODE_DEMO && location.protocol === "https:") {
  // updateViaCache "none" : pour savoir s'il existe une nouvelle version, toujours
  // demander sw.js au serveur, jamais à la mémoire du téléphone
  navigator.serviceWorker.register("sw.js", { updateViaCache: "none" }).then(reg => {
    // Sur Android, l'appli installée reste souvent ouverte en arrière-plan :
    // à chaque retour dessus, on vérifie s'il y a une mise à jour
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") reg.update().catch(() => {});
    });
  }).catch(() => {}); // pas de gardien possible : l'appli marche quand même
  // Quand la nouvelle version prend le relais, on recharge la page une fois pour l'afficher
  // (sauf à la toute première installation). Fiche ouverte : on attend qu'elle soit fermée.
  const avaitUneVersion = !!navigator.serviceWorker.controller;
  let dejaRecharge = false;
  const recharger = () => {
    if (dejaRecharge) return;
    if (form || $("#modal").innerHTML) return setTimeout(recharger, 3000);
    dejaRecharge = true;
    saveNotesNow();
    setTimeout(() => location.reload(), 300);
  };
  navigator.serviceWorker.addEventListener("controllerchange", () => { if (avaitUneVersion) recharger(); });
}

if (MODE_DEMO) demarrerDemo(); else demarrerFirebase();
