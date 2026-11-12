# Mes Archives BdF : feuille de route

## Contexte

Le complément VSTO actuel expose un onglet « Mes archives » avec quatre boutons (2021 à 2024), « Modop » et « À propos ».
Un complément web ne peut pas recréer un onglet personnalisé complet, donc le ruban n'a qu'un bouton qui ouvre un popup.
Le popup reprend les quatre années, leur état (Indisponible / Disponible) et le code couleur rouge / vert.

## Choix d'architecture

| Sujet | Décision | Raison |
|---|---|---|
| Interface | 1 bouton de ruban + popup (`displayDialogAsync`) | Le ruban Outlook ne garantit pas l'affichage/masquage dynamique des boutons |
| Montage du PST | Agent local sur `127.0.0.1:45731` (COM Outlook, `AddStoreEx` / `RemoveStore`) | Seul un processus local peut monter un PST dans Outlook de bureau |
| Emplacement des PST | Source OneDrive synchronisée ; copie locale dans `%LOCALAPPDATA%\MesArchives\PST` | Outlook ouvre la copie, pas l'original synchronisé |
| État côté web (OWA) | Flux Power Automate HTTP, lu via `webStatusEndpoint` | Pas de clé ni de secret dans le complément |
| Icônes | Icônes rouge / vert existantes | Pas de nouveau visuel à produire |

## Déjà réalisé (dans `main`)

- [x] Analyse du dépôt, validation des manifestes et du JS
- [x] Correction du Runtime (`Commands.Url`) et du nom de fichier 2024 (`ADR.pst`)
- [x] Route `/state` dans l'agent local
- [x] Popup à quatre années, remplaçant les 8 boutons du ruban (v2.0.9.0)

## Étapes restantes

### Démo rapide : Outlook classique sur Windows

Cette démo utilise un PST dans OneDrive **synchronisé sur le poste** et l'agent
local. Elle ne monte pas de PST dans Outlook Web et n'importe rien dans l'archive
Exchange Online. Tester d'abord avec un PST de démonstration, pas avec une archive
de production.

1. Sur le PC Windows, ouvrir Outlook classique, vérifier que le complément est
   déployé et que le bouton « Mes Archives » apparaît lors de la lecture d'un mail.
2. Dans le dossier OneDrive synchronisé, placer un fichier `ADR.pst` pour 2024
   (ou `ADR_2021.pst`, `ADR_2022.pst`, `ADR_2023.pst`) et sélectionner « Toujours conserver
   sur cet appareil ». L'agent cherche par nom dans les dossiers OneDrive; éviter
   plusieurs fichiers portant le même nom.
3. Récupérer `tools/local-pst-agent.ps1` depuis la version déployée du projet.
   Dans PowerShell, avec l'accord et la politique d'exécution de la DSI, lancer :
   `& "C:\chemin\vers\local-pst-agent.ps1"`. Laisser la fenêtre ouverte. L'agent
   écoute sur `127.0.0.1:45731`; il ne constitue pas un service sécurisé de
   production et doit être utilisé uniquement dans un environnement de démo fiable.
4. Ouvrir le complément, attendre la vérification des états, puis cliquer sur
   « Charger » pour l'année choisie. Le premier chargement copie le PST dans
   `%LOCALAPPDATA%\MesArchives\PST` et peut prendre plusieurs minutes.
5. Vérifier l'apparition de l'archive dans le volet de dossiers d'Outlook. Cliquer
   sur « Décharger » pour la retirer d'Outlook; la copie locale n'est pas supprimée.
   Si l'état est inconnu, vérifier l'agent puis cliquer sur « Actualiser ».

Les modifications apportées à la copie locale ne sont pas synchronisées vers le
PST source OneDrive. Une copie existante est réutilisée et n'est pas remplacée par
une nouvelle version de la source. Ce scénario nécessite une recette sur le poste
Windows avant présentation; le code n'a pas été validé avec Outlook réel ici.

### Étape 1 : Déploiement (bloquant)
- [ ] Supprimer l'ancienne application « MesArchives SPOT » du centre d'administration
- [ ] Redéployer via `https://dellasiegaps.github.io/MesArchivesAddin/manifest.xml`
- [ ] Si échec : récupérer la réponse de l'onglet Réseau (la console ne contient pas la cause)
- [ ] Contrôle : le bouton « Mes archives » apparaît dans Outlook

### Étape 2 : Aligner le popup sur l'existant
- [ ] Quatre tuiles côte à côte : « 2021 – Indisponible » devient « Disponible » après le clic
- [ ] Icônes rouge / vert actuelles
- [ ] Infobulles et messages d'erreur identiques à l'existant (ex. « Pas de clé disponible dans le NAS »)
- [ ] Bouton « À propos » : logo SPRINT, version, équipe ADR
- [ ] Bouton « Modop » : **en attente du lien de la modop**

### Étape 3 : PST depuis OneDrive
- [ ] Confirmer les chemins : `%OneDriveCommercial%\BdF\PST\YYYY.pst` (2024 : `ADR.pst`)
- [ ] Vérifier que les fichiers sont bien en « toujours conserver sur cet appareil » (sinon le montage échoue)
- [ ] Au premier montage, l'agent copie le fichier dans `%LOCALAPPDATA%\MesArchives\PST` ; les montages suivants réutilisent cette copie sans l'écraser ni la synchroniser avec OneDrive
- [ ] Le PST monté doit apparaître dans le volet de dossiers d'Outlook

### Étape 4 : Agent local
- [ ] Script d'installation (démarrage automatique à l'ouverture de session)
- [ ] Test sur un poste réel : `/state`, `/mount-pst`, `/unmount-pst`
- [ ] Contrôle de la signature du script / politique d'exécution PowerShell côté DSI

### Étape 5 : Power Automate
- [ ] Flux HTTP : GET `?year=YYYY` renvoie `fileExists` / `mounted`, POST met à jour l'état
- [ ] Renseigner `webStatusEndpoint` dans `src/archives-config.js`
- [ ] Prompt prêt à l'emploi : `docs/POWER-AUTOMATE-PROMPT.md`
- [ ] Limite connue : Power Automate cloud ne monte pas un PST dans Outlook de bureau (il faudrait Power Automate Desktop)

### Étape 6 : Recette et démo
- [ ] Scénario : indisponible, clic, montage, disponible, vérification dans le volet de dossiers, démontage
- [ ] Cas d'erreur : PST absent, agent arrêté, OneDrive non synchronisé
- [ ] Nettoyage : archiver les anciens manifestes (`Manifest-Final-Link.xml`, `Manifest-TEST-Sideload.xml`, `manifests/*`)

## Points ouverts

1. Lien de la modop
2. Cohérence de la sécurité : l'agent n'accepte que l'origine `https://dellasiegaps.github.io`
3. Hébergement définitif : GitHub Pages convient pour la démo, un hébergement interne est à prévoir pour la production
4. Rien n'a encore été testé dans un vrai Outlook
