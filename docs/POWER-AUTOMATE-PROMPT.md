# Prompt Power Automate – Mes Archives (flip-flop Disponible / Indisponible)

## Prompt à coller dans Copilot de Power Automate

> Crée un flux cloud instantané « MesArchives-Statut-PST » déclenché par « Lorsqu'une requête HTTP est reçue »
> (méthodes GET et POST, authentification Microsoft Entra ID, jamais de clé dans le code de l'add-in).
>
> Schéma JSON du POST : `{ "annee": number, "Monte": boolean, "Couleur": "Vert"|"Rouge", "sharePointPath": string }`.
> GET : paramètre `year` (2021 à 2024).
>
> **GET** : action SharePoint « Lister les fichiers du dossier » sur `Shared Documents/PST`, cherche `<year>.pst`
> (pour 2024 : `ADR.pst`). Réponds 200 avec
> `{ "fileExists": true|false, "mounted": <colonne Monte>, "couleur": <colonne Couleur> }`.
>
> **POST** : valide `annee` entre 2021 et 2024, puis « Mettre à jour l'élément » dans la liste SharePoint
> « ArchivesPST » (colonnes : Annee, Monte [Oui/Non], Couleur [Rouge/Vert], DerniereMiseAJour = utcNow()).
> Réponds 200 `{ "ok": true, "annee": ..., "Monte": ..., "Couleur": ... }`, sinon 400/500 avec `{ "ok": false, "error": "..." }`.
> Ajoute l'en-tête CORS `Access-Control-Allow-Origin: https://dellasiegaps.github.io`.

## Règle flip-flop
- Disponible = `Monte = true` = **Vert** ; Indisponible = `Monte = false` = **Rouge**.
- Le bouton ne passe au vert qu'après montage réussi du PST (Outlook bureau : agent local `tools/local-pst-agent.ps1` ;
  Outlook Web : vérification `fileExists` puis POST).

## Après création
Copier l'URL HTTPS du flux dans `src/archives-config.js` → `webStatusEndpoint`.

## Comportement des boutons (v2.0.8.0)
Chaque année a deux boutons : `archiveAAAA_mount` (rouge, « Indisponible ») et `archiveAAAA_unmount` (vert, « Disponible »).
Un seul est visible à la fois (repli sur activé/désactivé si Outlook refuse `visible`).
Au démarrage, l'état est lu depuis la présence réelle du PST : agent `/state` (bureau) ou GET du flux (Web).
