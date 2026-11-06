# Journal des nouveautés — Quoi faire ?

## v1.2.0 — 2026-11-06
- Filtres repliés (Idées, À venir, Adresses) derrière une petite pastille qui résume les choix ; un point orange signale un filtre actif. On la touche pour déplier.
- Pastilles plus discrètes : fines, sans contour, grisées.
- Plus de noir pour ce qui est choisi : une couleur douce par ligne de filtres (pervenche, orange, vert, rose, ambre), et pervenche pour les onglets et les bascules.

## v1.1.0 — 2026-11-06
- Mises à jour automatiques fiables : le téléphone vérifie la nouvelle version à chaque retour dans l'appli et se recharge tout seul (comme Stuff to go).
- À venir : trois vues, Cartes / Liste / Calendrier. La liste montre une ligne par événement avec sa petite image.
- Cartes plus compactes : image plus basse, avec seulement les étiquettes en haut et les boutons en bas ; le titre passe dessous.
- Trois genres de dates, chacun avec son icône : un seul jour (concert…), sans fin connue (films : « À l'affiche », « depuis le… »), période (expos : « jusqu'au… », avec une barre de temps restant).
- Nouveau statut « Réservé », avec une heure facultative : contour vert, étiquette verte, remonte en tête. Bouton billet sur chaque carte pour le marquer d'un geste.
- Calendrier : vue Semaine par défaut (nom des jours, heure des réservés), et vue Mois au choix. Barre qui s'efface = sans fin connue, contour vert = réservé.
- Les films ne reçoivent plus de date de fin inventée à +1 mois (ils disparaissent quand même des « à venir » 1 mois après leur sortie).
- Paramètres : choix de la vue au démarrage pour Idées (Tirage / Liste) et À venir (Cartes / Liste / Calendrier), le même sur le téléphone et le PC.
- Excel : nouvelle colonne « Heure » dans l'onglet Événements.
- Les anciennes icônes sont rangées dans `icones/anciennes/`.

## v1.0.0 — 2026-11-06
Version stable : la migration depuis Google Sheet est terminée, toutes les données sont dans Firebase.
- Nouvelle icône : un billet de sortie avec un « ? » sur fond corail doux, signée du tampon rond commun à mes applis (versions normale et Android « maskable »).

## v0.2 — 2026-11-05
- Lecteur de liens branché : coller le lien d'un événement ou d'un lieu Google Maps remplit la fiche tout seul, et les adresses sont placées sur la carte.
- Correction : « nouveaux » au lieu de « nouveaus » dans le résumé d'import.

## v0.1 — 2026-11-05
Première version Firebase. Elle remplace l'ancienne version Google Sheet + Apps Script, gardée en archive.
- Interface identique à l'ancienne : Idées (tirage et liste, groupes, épingles), À venir (liste et calendrier), Adresses (liste et carte), Radar avec prompt IA, Notes.
- Données dans Firebase (compte Google), synchronisées entre téléphone et PC, utilisables hors ligne.
- Indicateur de synchro à côté de la date : synchro / en attente / hors ligne.
- Suppression immédiate avec « Annuler » pendant quelques secondes.
- Nouveau thème « Automatique » (par défaut) : Dure Vie le jour, Sombre la nuit, selon l'appareil.
- Export Excel au style commun (Idées, Événements, Adresses, Sources, Notes) ; import du même fichier ou de l'ancien Google Sheet.
- Lecture des liens (événements, Google Maps) par un mini Apps Script séparé, qui ne voit pas les données.
- Partage Android : un lien partagé vers l'appli ouvre directement la fiche à remplir.
- Version affichée dans Paramètres.
