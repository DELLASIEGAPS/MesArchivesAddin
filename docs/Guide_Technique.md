# Mes Archives — guide technique et démo

## 1. Le message à faire passer

> « J'ai cadré le besoin : retrouver les archives par année, sans perdre les
> habitudes des utilisateurs. J'ai retenu une interface web et un agent local.
> Aujourd'hui, l'agent utilise encore Outlook classique pour ouvrir les PST.
> Je distingue ce qui marche de la cible sans dépendance. »

Ce guide décrit la branche `fix/web-cloud-archive-actions` examinée le
26 novembre 2026. Les lignes sont des points de départ. Elles peuvent bouger
après une modification. La branche locale n'est pas forcément la version
publiée sur GitHub Pages ou installée dans Outlook.

### Ce qui marche aujourd'hui

- Le dialogue présente les années 2021 à 2024.
- Sur Outlook classique, l'agent peut charger et décharger une **copie** du PST.
- Le visualiseur local lit les dossiers, les messages en texte brut et les
  pièces jointes des archives montées.
- Le visualiseur peut être intégré au dialogue ou ouvert dans un onglet local.
- Le mode web du dialogue utilise un **marqueur cloud** distinct du montage.
  Le flux n'est pas configuré dans la version examinée.

### Ce qui ne marche pas encore

- Le visualiseur ne lit pas un PST sans Outlook classique ouvert sur le même PC.
- Outlook Web ne monte pas un fichier PST. Le marqueur web ne donne pas accès
  aux messages et ne prouve pas qu'un PST est monté.
- Le prototype d'import Graph n'est pas branché au complément et n'est pas prêt
  pour la production.
- Il n'y a pas d'exécutable autonome validé dans le dépôt. Un lanceur `.exe`
  ne retire pas la dépendance Outlook.
- Le CRM est un projet séparé. Il n'est pas couvert par ce guide.

## 2. Démo : le parcours qui existe

**Avant la réunion :** utiliser un PST de démonstration et un poste autorisé.
Ne pas exposer de vrais messages ou de pièces jointes confidentielles.

1. Vérifier la branche et le déploiement. Le manifeste pointe vers GitHub Pages ;
   il ne lit pas automatiquement la branche locale.
2. Ouvrir Outlook **classique** sur le PC. Vérifier le profil et le complément.
3. Vérifier les sources PST et les copies locales. Une copie déjà présente est
   réutilisée. Elle peut contenir des changements absents de la source.
4. Si aucun agent ne tourne, lancer Windows PowerShell 5.1, selon la politique
   de la DSI. Avec les sources dans `C:\DEV` :

   ```powershell
   powershell.exe -NoProfile -File 'C:\DEV\MesArchivesAddin\tools\local-pst-agent.ps1' -PstSourceDirectory 'C:\DEV'
   ```

   Garder cette fenêtre ouverte. Si le port est déjà pris, identifier l'agent
   existant avant tout arrêt. Ne pas lancer deux agents sur le même port.
5. Dans Outlook classique, ouvrir « Mes Archives ». Cliquer sur **Charger**.
   Vérifier l'archive dans le volet de dossiers Outlook.
6. Ouvrir le visualiseur dans le dialogue ou sur
   `http://127.0.0.1:45731/view`. Choisir l'année et un dossier, puis lire un
   message fictif. Tester une pièce jointe fictive.
7. Cliquer sur **Décharger** si le scénario le prévoit. Cela retire la copie du
   profil Outlook, sans supprimer la copie ni modifier la source.

Ne pas remplacer ou supprimer une copie existante sans vérifier son contenu et
fermer le PST dans Outlook. Le script
[`create-demo-pst.ps1`](../tools/create-demo-pst.ps1) crée un PST fictif dans un
répertoire choisi et refuse d'écraser un fichier existant. Il utilise aussi
Outlook classique. Son nom `ADR_2021.pst` peut correspondre à une copie de
travail existante : vérifier cette copie avant toute démo.

**Point de contrôle rapide :**

```powershell
Invoke-RestMethod -Uri 'http://127.0.0.1:45731/state' -Method Post `
  -ContentType 'application/json' -Body '{"fileName":"ADR_2021.pst"}'
```

`outlookRunning` indique si l'agent détecte Outlook. `mounted` indique le
montage de la copie dans ce profil. `local` indique une présence locale, pas
un montage. Si l'agent ne répond pas, vérifier la fenêtre PowerShell et le port.
Si `outlookRunning` vaut `false`, ouvrir Outlook classique puis actualiser.
La vérification ne remplace pas un test complet avec Outlook réel.

## 3. Comment supprimer la dépendance Outlook

### Le constat dans le code

Le dialogue appelle l'agent local. L'agent trouve le PST et crée une copie.
Mais il demande à Outlook de monter cette copie avec `AddStoreEx`. Il récupère
ensuite les dossiers et les messages via Outlook COM. Les points clés sont
[`Get-RunningOutlook`](../tools/local-pst-agent.ps1#L102),
[`Mount-Pst`](../tools/local-pst-agent.ps1#L127) et
[`Get-MountedStore`](../tools/local-pst-agent.ps1#L219). La recherche du
visualiseur vérifie aussi `outlookRunning` et `mounted`.

### Option A — lecteur PST local indépendant

**But :** lire un PST sur Windows sans lancer Outlook classique. Le résultat
serait un visualiseur séparé. Il ne ferait pas apparaître le PST dans le volet
de dossiers natif d'Outlook.

Travail à réaliser :

1. Choisir un moteur PST fiable, compatible Windows, avec licence et provenance
   vérifiées par la DSI. Le prototype `pypff` existant est expérimental.
2. Lire une copie ou un fichier de test. Gérer les fichiers verrouillés et les
   erreurs de lecture. Ne pas écraser les PST existants.
3. Définir un contrat clair pour les dossiers, messages et pièces jointes :
   identifiants propres à chaque archive, pagination et erreurs explicites.
4. Adapter `/state`, `/folder`, `/message`, `/attachment` et le visualiseur.
   Remplacer l'idée « monté dans Outlook » par un état de lecture indépendant.
   Garder séparé le mode actuel de montage Outlook.
5. Tester des PST fictifs variés : grands fichiers, pièces jointes, caractères
   accentués, champs absents, dossiers imbriqués et éléments non pris en charge.
6. Faire valider la sécurité locale, les droits sur les fichiers, les mises à
   jour et le packaging. Signer et autoriser l'exécutable selon la politique
   Windows. Tester l'exécutable sur le poste de démo.

**Pas encore développé ni validé ici.** Une stratégie Windows a bloqué le
lancement du lanceur compilé lors d'un essai précédent. Il faut une validation
administrateur ; ne pas contourner cette protection.

### Option B — migration Microsoft 365 / Exchange Online

**But :** importer les archives dans une boîte Exchange ou son archive, puis
les consulter dans Outlook Web. Microsoft propose le service d'import PST de
Purview. Il faut cadrer les droits, la licence, le stockage, la confidentialité,
la correspondance PST/boîte et les limites de taille avec la DSI.

Ce n'est pas un montage à la demande : l'import crée des données dans la boîte.
Il faut décider de leur durée de conservation et de leur suppression. Tester
d'abord sur une boîte et des PST de démonstration. Voir la
[présentation Microsoft de l'import PST](https://learn.microsoft.com/en-us/purview/pst-import-overview)
et la [FAQ Microsoft](https://learn.microsoft.com/en-us/purview/pst-import-faq).

Le dossier [`services/pst-import`](../services/pst-import/README.md) explore une
autre voie : conversion PST en MIME et appel Graph `importMime` **bêta**.
Il n'est pas raccordé au complément. Il ne gère pas encore la reprise fiable
d'un import partiel. Un déchargement confirmé peut supprimer un dossier marqué
et tout son contenu. Ne pas l'utiliser sur une boîte de production.

**Phrase de démo :** « Pour la démo, je montre le parcours actuel avec Outlook
classique. Pour supprimer cette dépendance, je propose soit un lecteur PST local
séparé, soit une migration vers Exchange. Ce sont deux usages différents. »

## 4. Architecture et circulation

```text
manifest.xml
  -> commands.html -> commands.js -> fenêtre archives-dialog.html
       -> mode classique : agent HTTP local -> copie PST -> Outlook COM
       -> mode web : marqueur Power Automate (non configuré ici)
       -> visualiseur local : dossiers / messages / pièces jointes

Prototype séparé, non raccordé :
  Azure Function -> OneDrive -> lecteur PST -> MIME -> Graph bêta -> Exchange
```

| Élément | Rôle |
|---|---|
| [`manifest.xml`](../manifest.xml#L106) | Déclare les commandes Outlook. `openArchives` est lié à la page de commandes. Les URL sont aux lignes [177](../manifest.xml#L177) et [290](../manifest.xml#L290). |
| [`commands.html`](../src/commands/commands.html#L4) | Charge Office.js, la configuration et le script de commande. |
| [`archives-config.js`](../src/archives-config.js#L1) | Adresse du lecteur local et adresse du flux web. `webStatusEndpoint` est vide. |
| [`commands.js`](../src/commands/commands.js#L1) | Ouvre le dialogue et reçoit ses messages. |
| [`archives-dialog.html`](../src/commands/archives-dialog.html#L73) | Affiche les années, les états, les boutons et le visualiseur intégré. |
| [`local-pst-agent.ps1`](../tools/local-pst-agent.ps1#L1) | Cherche la source, gère la copie, parle à Outlook et expose les routes locales. |
| [`local-pst-view.html`](../tools/local-pst-view.html#L152) | Affiche dossiers et messages. Lance la recherche et prépare les brouillons. |
| [`taskpane.html`](../src/taskpane/taskpane.html) et [`taskpane.js`](../src/taskpane/taskpane.js#L1) | Volet simple : liste les années et propose le lien du visualiseur. Le traitement au chargement commence à la ligne 4. |
| [`create-demo-pst.ps1`](../tools/create-demo-pst.ps1#L1) | Script de création d'un PST fictif. Il n'a pas de fonction nommée. |

### Routes de l'agent

Toutes les routes métier ci-dessous utilisent POST. `/view` sert la page en GET.
La boucle HTTP démarre [ligne 360](../tools/local-pst-agent.ps1#L360).

| Route | Point de départ | Rôle |
|---|---|---|
| `/mount-pst` | [396](../tools/local-pst-agent.ps1#L396) | Crée ou réutilise la copie, puis demande le montage. |
| `/unmount-pst` | [402](../tools/local-pst-agent.ps1#L402) | Retire le PST du profil Outlook, pas du disque. |
| `/state` | [411](../tools/local-pst-agent.ps1#L411) | Retourne présence locale, état Outlook et montage. |
| `/tree` | [412](../tools/local-pst-agent.ps1#L412) | Retourne une arborescence limitée à six niveaux. |
| `/folder` | [418](../tools/local-pst-agent.ps1#L418) | Retourne sous-dossiers et une page de 50 éléments. |
| `/message` | [424](../tools/local-pst-agent.ps1#L424) | Retourne le texte, les contacts et les pièces jointes. |
| `/attachment` | [428](../tools/local-pst-agent.ps1#L428) | Transmet une pièce jointe demandée. |
| `/status` | [434](../tools/local-pst-agent.ps1#L434) | Retourne l'état de l'agent et les racines OneDrive détectées. |

L'agent écoute en local, mais ce n'est pas un service de production sécurisé.
Le CORS ne constitue pas une authentification des logiciels locaux. Protéger le
poste et faire valider l'origine web, les accès locaux et les données par la DSI.

## 5. Fonctions du complément

### [`commands.js`](../src/commands/commands.js)

| Fonction | Ligne | Ce qu'elle fait |
|---|---:|---|
| `saveStates` | [6](../src/commands/commands.js#L6) | Enregistre les états dans les paramètres itinérants Office. |
| `onDialogMessage` | [15](../src/commands/commands.js#L15) | Reçoit les états ou la demande de fermeture du dialogue. |
| `openArchives` | [26](../src/commands/commands.js#L26) | Ouvre le dialogue Office et signale la fin de la commande. |

### [`archives-dialog.html`](../src/commands/archives-dialog.html)

| Fonction | Ligne | Ce qu'elle fait |
|---|---:|---|
| `isWeb` | [77](../src/commands/archives-dialog.html#L77) | Distingue Office Online du mode classique. |
| `beginProgress` | [100](../src/commands/archives-dialog.html#L100) | Démarre l'affichage des étapes. |
| `updateProgress` | [110](../src/commands/archives-dialog.html#L110) | Actualise les étapes et la barre. |
| `endProgress` | [120](../src/commands/archives-dialog.html#L120) | Termine la progression ou signale l'échec. |
| `setViewerExpanded` | [128](../src/commands/archives-dialog.html#L128) | Affiche ou masque le lecteur intégré. |
| `render` | [143](../src/commands/archives-dialog.html#L143) | Construit les quatre tuiles et adapte les boutons au mode. |
| `agent` | [177](../src/commands/archives-dialog.html#L177) | Envoie une demande JSON à l'agent local. |
| `agentState` | [190](../src/commands/archives-dialog.html#L190) | Contrôle et adapte l'état reçu de l'agent. |
| `webStatusUrl` | [205](../src/commands/archives-dialog.html#L205) | Construit l'URL du flux web avec l'année. |
| `webState` | [214](../src/commands/archives-dialog.html#L214) | Lit le marqueur web et contrôle sa réponse. |
| `setWebMarker` | [227](../src/commands/archives-dialog.html#L227) | Demande la mise à jour du marqueur cloud. |
| `readState` | [240](../src/commands/archives-dialog.html#L240) | Choisit l'état web ou l'état local. |
| `presenceDetail` | [242](../src/commands/archives-dialog.html#L242) | Prépare le texte adapté au mode et à l'état Outlook. |
| `presenceText` | [253](../src/commands/archives-dialog.html#L253) | Décrit la présence locale ou cloud connue. |
| `refreshStates` | [263](../src/commands/archives-dialog.html#L263) | Vérifie les années et montre les erreurs sans inventer d'état. |
| `toggle` | [302](../src/commands/archives-dialog.html#L302) | Vérifie, agit, puis confirme le résultat avant de l'afficher. |

Les autres blocs en bas du fichier raccordent les boutons, le retour au premier
plan et la fermeture. Ils ne sont pas des fonctions nommées.

## 6. Fonctions de l'agent local

### [`local-pst-agent.ps1`](../tools/local-pst-agent.ps1)

| Fonction | Ligne | Ce qu'elle fait |
|---|---:|---|
| `Get-OneDriveRoots` | [20](../tools/local-pst-agent.ps1#L20) | Liste les racines OneDrive détectées. |
| `Find-PstFile` | [28](../tools/local-pst-agent.ps1#L28) | Cherche un nom PST valide dans la source choisie ou OneDrive. |
| `Get-WorkingPstPath` | [45](../tools/local-pst-agent.ps1#L45) | Construit le chemin de copie dans `%LOCALAPPDATA%`. |
| `Get-OrCreateWorkingPst` | [58](../tools/local-pst-agent.ps1#L58) | Réutilise la copie ou copie la source via un fichier temporaire. |
| `Add-CorsHeaders` | [82](../tools/local-pst-agent.ps1#L82) | Ajoute les en-têtes pour l'origine web autorisée. |
| `Write-Json` | [89](../tools/local-pst-agent.ps1#L89) | Écrit le statut et la réponse JSON. |
| `Get-RunningOutlook` | [102](../tools/local-pst-agent.ps1#L102) | Trouve une instance Outlook classique déjà ouverte. |
| `Test-PstMounted` | [106](../tools/local-pst-agent.ps1#L106) | Cherche le chemin du PST dans les magasins Outlook. |
| `Set-PstDisplayName` | [113](../tools/local-pst-agent.ps1#L113) | Donne un nom visible à la racine du PST monté. |
| `Mount-Pst` | [127](../tools/local-pst-agent.ps1#L127) | Demande à Outlook de monter la copie si besoin. |
| `Unmount-Pst` | [137](../tools/local-pst-agent.ps1#L137) | Demande à Outlook de retirer la copie du profil. |
| `Test-CloudOnlyPlaceholder` | [153](../tools/local-pst-agent.ps1#L153) | Vérifie les attributs d'un fichier seulement en ligne. |
| `Test-WithinOneDrive` | [158](../tools/local-pst-agent.ps1#L158) | Vérifie si le chemin se trouve sous une racine OneDrive. |
| `Get-PstState` | [168](../tools/local-pst-agent.ps1#L168) | Réunit présence, copie, détection Outlook et montage. |
| `Get-PstTree` | [196](../tools/local-pst-agent.ps1#L196) | Demande l'arbre d'un PST déjà monté. |
| `ConvertTo-FolderNode` | [209](../tools/local-pst-agent.ps1#L209) | Convertit les dossiers Outlook en arbre, avec limite de profondeur. |
| `Get-MountedStore` | [219](../tools/local-pst-agent.ps1#L219) | Récupère le magasin Outlook correspondant à la copie. |
| `Get-ArchiveFolder` | [234](../tools/local-pst-agent.ps1#L234) | Retrouve un dossier et vérifie son appartenance à l'archive. |
| `Get-ArchiveMail` | [245](../tools/local-pst-agent.ps1#L245) | Retrouve un message et vérifie son type et son archive. |
| `Get-ArchiveFolderContents` | [254](../tools/local-pst-agent.ps1#L254) | Liste les sous-dossiers et une page de 50 éléments. |
| `Get-ArchiveAddress` | [283](../tools/local-pst-agent.ps1#L283) | Tente de résoudre une adresse SMTP ou Exchange. |
| `Get-ArchiveMessage` | [303](../tools/local-pst-agent.ps1#L303) | Prépare le détail d'un message, ses contacts et pièces jointes. |
| `Write-ArchiveAttachment` | [328](../tools/local-pst-agent.ps1#L328) | Transmet la pièce jointe via un fichier temporaire nettoyé ensuite. |
| `Test-ViewerRequest` | [352](../tools/local-pst-agent.ps1#L352) | Vérifie l'origine et le type JSON des routes du lecteur. |

## 7. Fonctions du visualiseur

### [`local-pst-view.html`](../tools/local-pst-view.html)

| Fonction | Ligne | Ce qu'elle fait |
|---|---:|---|
| `beginProgress` | [177](../tools/local-pst-view.html#L177) | Démarre la progression d'une opération. |
| `update` | [183](../tools/local-pst-view.html#L183) | Fonction interne : actualise les étapes de cette opération. |
| `end` | [199](../tools/local-pst-view.html#L199) | Méthode interne : termine la progression ou montre l'échec. |
| `renderContacts` | [208](../tools/local-pst-view.html#L208) | Affiche les contacts et crée les liens mail valides. |
| `request` | [230](../tools/local-pst-view.html#L230) | Appelle l'agent et contrôle la réponse ou l'erreur. |
| `link` | [250](../tools/local-pst-view.html#L250) | Crée un bouton dans une liste. |
| `clearReading` | [261](../tools/local-pst-view.html#L261) | Vide le message sélectionné. |
| `formatDate` | [267](../tools/local-pst-view.html#L267) | Formate la date pour l'affichage français. |
| `folderButton` | [275](../tools/local-pst-view.html#L275) | Crée un bouton de dossier avec son niveau et son compteur. |
| `showFolder` | [293](../tools/local-pst-view.html#L293) | Affiche le chemin, les sous-dossiers et la page de messages. |
| `showMessage` | [365](../tools/local-pst-view.html#L365) | Affiche le détail d'un message sélectionné. |
| `downloadAttachment` | [406](../tools/local-pst-view.html#L406) | Demande et remet une pièce jointe au navigateur. |
| `prepareDraft` | [428](../tools/local-pst-view.html#L428) | Prépare les champs d'une réponse ou d'un transfert. Aucun envoi automatique. |
| `add` | [436](../tools/local-pst-view.html#L436) | Fonction interne : valide et déduplique les adresses du brouillon. |
| `openArchive` | [500](../tools/local-pst-view.html#L500) | Ouvre la racine de l'archive choisie. |
| `searchArchives` | [504](../tools/local-pst-view.html#L504) | Parcourt les archives montées et remonte les résultats ou les erreurs. |
| `normalize` | [521](../tools/local-pst-view.html#L521) | Fonction interne : ignore casse et accents pour la recherche. |

La recherche ne lit pas le contenu des pièces jointes. L'ouverture du brouillon
dans Outlook Web peut placer son contenu dans une URL et des journaux. Vérifier
les règles de confidentialité avant la démo.

## 8. Fonctions du prototype d'import — non raccordé

### [`archive_service.py`](../services/pst-import/archive_service.py)

| Fonction ou classe | Ligne | Ce qu'elle fait |
|---|---:|---|
| `ArchiveError` | [24](../services/pst-import/archive_service.py#L24) | Type d'erreur du service. |
| `GraphRequestError.__init__` | [29](../services/pst-import/archive_service.py#L29) | Garde le code HTTP d'une erreur Graph. |
| `pst_name` | [34](../services/pst-import/archive_service.py#L34) | Valide l'année et donne le nom PST. |
| `folder_name` | [40](../services/pst-import/archive_service.py#L40) | Donne le nom du dossier de l'année. |
| `_graph` | [45](../services/pst-import/archive_service.py#L45) | Envoie une requête Graph ou télécharge un fichier. |
| `find_folder` | [65](../services/pst-import/archive_service.py#L65) | Cherche un dossier et refuse les noms en double. |
| `_is_owned_archive_folder` | [74](../services/pst-import/archive_service.py#L74) | Vérifie le marqueur de propriété et l'année. |
| `ensure_archive_folder` | [89](../services/pst-import/archive_service.py#L89) | Crée ou retrouve un dossier marqué MesArchives. |
| `ensure_folder` | [115](../services/pst-import/archive_service.py#L115) | Retrouve ou crée un sous-dossier. |
| `download_pst` | [123](../services/pst-import/archive_service.py#L123) | Télécharge le PST de OneDrive vers un chemin temporaire. |
| `load_archive` | [128](../services/pst-import/archive_service.py#L128) | Convertit et importe les messages, dossier par dossier. |
| `unload_archive` | [156](../services/pst-import/archive_service.py#L156) | Supprime le dossier marqué et son contenu. Action destructive. |

### [`pst_to_mime.py`](../services/pst-import/pst_to_mime.py)

| Fonction ou classe | Ligne | Ce qu'elle fait |
|---|---:|---|
| `PstReadError` | [16](../services/pst-import/pst_to_mime.py#L16) | Type d'erreur de lecture ou de conversion. |
| `_read_attachment` | [20](../services/pst-import/pst_to_mime.py#L20) | Lit une pièce jointe et vérifie sa taille. |
| `_message_to_mime` | [30](../services/pst-import/pst_to_mime.py#L30) | Construit un message MIME ; refuse certains champs incomplets. |
| `iter_pst_messages` | [93](../services/pst-import/pst_to_mime.py#L93) | Ouvre le PST et produit les messages convertis. |
| `walk` | [102](../services/pst-import/pst_to_mime.py#L102) | Fonction interne : parcourt récursivement les dossiers. |

### [`graph_import.py`](../services/pst-import/graph_import.py) et [`function_app.py`](../services/pst-import/function_app.py)

| Fonction ou classe | Ligne | Ce qu'elle fait |
|---|---:|---|
| `GraphImportError` | [16](../services/pst-import/graph_import.py#L16) | Type d'erreur d'import. |
| `import_mime` | [20](../services/pst-import/graph_import.py#L20) | Envoie un MIME à Graph bêta et vérifie la réponse. |
| `_reply` | [27](../services/pst-import/function_app.py#L27) | Prépare une réponse HTTP JSON. |
| `_claims` | [33](../services/pst-import/function_app.py#L33) | Vérifie le jeton et le scope attendu. |
| `archive` | [48](../services/pst-import/function_app.py#L48) | Contrôle la demande, le jeton et le consentement, puis lance l'action. |

## 9. Tests et fichiers secondaires

Les tests ne prouvent pas à eux seuls le fonctionnement dans un vrai Outlook.
Ils utilisent notamment des objets simulés. Pour les tests locaux sans toucher
aux PST réels :

```powershell
node --test tools\local-pst-agent.test.js tools\local-pst-view.test.js src\commands\archives-dialog.test.js
powershell.exe -NoProfile -File tools\local-pst-reader.test.ps1
```

| Fichier | Fonctions et cas repérés |
|---|---|
| [`local-pst-agent.test.js`](../tools/local-pst-agent.test.js#L11) | `unusedPort` [11](../tools/local-pst-agent.test.js#L11), `startAgent` [22](../tools/local-pst-agent.test.js#L22), `state` [51](../tools/local-pst-agent.test.js#L51) ; cas source locale [61](../tools/local-pst-agent.test.js#L61) et OneDrive [104](../tools/local-pst-agent.test.js#L104). |
| [`local-pst-reader.test.ps1`](../tools/local-pst-reader.test.ps1#L14) | `Assert` [14](../tools/local-pst-reader.test.ps1#L14), `Collection` [17](../tools/local-pst-reader.test.ps1#L17) ; vérifie pagination, message, adresse, pièce jointe et séparation des archives avec des objets simulés. |
| [`demo-pst.test.ps1`](../tools/demo-pst.test.ps1#L1) | Pas de fonction nommée. Monte un PST de démo donné en argument, vérifie son contenu puis le retire d'Outlook. Ne pas utiliser sur un PST réel. |
| [`archives-dialog.test.js`](../src/commands/archives-dialog.test.js#L8) | `createDialog` [8](../src/commands/archives-dialog.test.js#L8), `element` [10](../src/commands/archives-dialog.test.js#L10), `response` [200](../src/commands/archives-dialog.test.js#L200). Les cas couvrent ouverture, progression, séparation web/local, confirmation et erreurs. |
| [`local-pst-view.test.js`](../tools/local-pst-view.test.js#L7) | `createViewer` [7](../tools/local-pst-view.test.js#L7), `element` [9](../tools/local-pst-view.test.js#L9), `jsonResponse` [43](../tools/local-pst-view.test.js#L43). Les cas couvrent lecture, pagination, recherche, erreurs et brouillons. |
| [`test_archive_service.py`](../services/pst-import/tests/test_archive_service.py#L20) | Cas nom/année [20](../services/pst-import/tests/test_archive_service.py#L20), téléchargement [31](../services/pst-import/tests/test_archive_service.py#L31), marqueur et suppression [51](../services/pst-import/tests/test_archive_service.py#L51). |
| [`test_function_app.py`](../services/pst-import/tests/test_function_app.py#L9) | `FakeMSAL.acquire_token_on_behalf_of` [10](../services/pst-import/tests/test_function_app.py#L10), `request` [26](../services/pst-import/tests/test_function_app.py#L26) ; confirmation de suppression [37](../services/pst-import/tests/test_function_app.py#L37). |
| [`test_graph_import.py`](../services/pst-import/tests/test_graph_import.py#L12) | Classe simulée `Response` : `__init__` [15](../services/pst-import/tests/test_graph_import.py#L15), `__enter__` [18](../services/pst-import/tests/test_graph_import.py#L18), `__exit__` [21](../services/pst-import/tests/test_graph_import.py#L21), `read` [24](../services/pst-import/tests/test_graph_import.py#L24) ; cas import et erreurs [28](../services/pst-import/tests/test_graph_import.py#L28). |
| [`test_pst_to_mime.py`](../services/pst-import/tests/test_pst_to_mime.py#L10) | Cas conversion [10](../services/pst-import/tests/test_pst_to_mime.py#L10), pièce jointe tronquée [48](../services/pst-import/tests/test_pst_to_mime.py#L48), nom absent [71](../services/pst-import/tests/test_pst_to_mime.py#L71), adresses et RTF [85](../services/pst-import/tests/test_pst_to_mime.py#L85). |

### Détail des cas de test

Chaque lien mène au début du cas. Les fonctions de préparation sont dans le
tableau précédent.

**Dialogue :** [`archives-dialog.test.js`](../src/commands/archives-dialog.test.js).

| Ligne | Vérification |
|---:|---|
| [78](../src/commands/archives-dialog.test.js#L78) | Actualisation au retour sur la fenêtre visible. |
| [98](../src/commands/archives-dialog.test.js#L98) | Lecteur intégré et onglet séparé proposés. |
| [112](../src/commands/archives-dialog.test.js#L112) | Fermeture du dialogue avec lecteur ouvert. |
| [119](../src/commands/archives-dialog.test.js#L119) | Masquage demandé par le bon lecteur. |
| [132](../src/commands/archives-dialog.test.js#L132) | Message d'attente et étapes de progression. |
| [150](../src/commands/archives-dialog.test.js#L150) | Adresse du lecteur local par défaut. |
| [156](../src/commands/archives-dialog.test.js#L156) | Mode web limité au marqueur cloud. |
| [188](../src/commands/archives-dialog.test.js#L188) | Erreur explicite si le flux web manque. |
| [204](../src/commands/archives-dialog.test.js#L204) | Lecture de l'état à l'ouverture. |
| [214](../src/commands/archives-dialog.test.js#L214) | Confirmation du montage. |
| [234](../src/commands/archives-dialog.test.js#L234) | Confirmation du démontage. |
| [254](../src/commands/archives-dialog.test.js#L254) | Pas de succès sans confirmation. |
| [266](../src/commands/archives-dialog.test.js#L266) | État inconnu si confirmation impossible. |
| [278](../src/commands/archives-dialog.test.js#L278) | Pas d'action si l'état a changé entre-temps. |
| [293](../src/commands/archives-dialog.test.js#L293) | Erreur si l'agent est indisponible. |

**Visualiseur :** [`local-pst-view.test.js`](../tools/local-pst-view.test.js).

| Ligne | Vérification |
|---:|---|
| [47](../tools/local-pst-view.test.js#L47) | Dossiers, messages, détail et pagination. |
| [86](../tools/local-pst-view.test.js#L86) | Erreur de dossier sans ancien nom trompeur. |
| [99](../tools/local-pst-view.test.js#L99) | Pas de lien pour une adresse invalide. |
| [106](../tools/local-pst-view.test.js#L106) | Recherche dans les archives montées et leurs dossiers. |
| [134](../tools/local-pst-view.test.js#L134) | Recherche incomplète signalée en cas d'erreur. |
| [141](../tools/local-pst-view.test.js#L141) | Recherche sur contacts, dates et noms de pièces jointes. |
| [163](../tools/local-pst-view.test.js#L163) | Message clair pour une erreur non JSON. |
| [171](../tools/local-pst-view.test.js#L171) | Arrêt de la recherche. |
| [185](../tools/local-pst-view.test.js#L185) | Demande de masquage au dialogue parent. |
| [193](../tools/local-pst-view.test.js#L193) | Préparation du brouillon et des destinataires. |
| [214](../tools/local-pst-view.test.js#L214) | Lien Outlook Web sans envoi et limite de taille. |

**Prototype :** [`test_archive_service.py`](../services/pst-import/tests/test_archive_service.py),
[`test_function_app.py`](../services/pst-import/tests/test_function_app.py),
[`test_graph_import.py`](../services/pst-import/tests/test_graph_import.py) et
[`test_pst_to_mime.py`](../services/pst-import/tests/test_pst_to_mime.py).

| Fonction de test | Ligne | Vérification |
|---|---:|---|
| `test_pst_and_folder_names` | [20](../services/pst-import/tests/test_archive_service.py#L20) | Noms PST et dossiers. |
| `test_rejects_unknown_year` | [26](../services/pst-import/tests/test_archive_service.py#L26) | Refus d'une année inconnue. |
| `test_download_reads_while_response_is_open` | [31](../services/pst-import/tests/test_archive_service.py#L31) | Lecture pendant que la réponse est ouverte. |
| `test_download_reports_graph_http_error` | [44](../services/pst-import/tests/test_archive_service.py#L44) | Erreur HTTP au téléchargement. |
| `test_new_archive_folder_is_created_with_ownership_marker` | [51](../services/pst-import/tests/test_archive_service.py#L51) | Marqueur à la création. |
| `test_unmarked_archive_folder_cannot_be_deleted` | [70](../services/pst-import/tests/test_archive_service.py#L70) | Refus de supprimer un dossier non marqué. |
| `test_existing_unmarked_folder_is_not_reused_for_import` | [85](../services/pst-import/tests/test_archive_service.py#L85) | Refus de réutiliser un dossier non marqué. |
| `test_archive_folder_with_marker_for_another_year_is_not_owned` | [94](../services/pst-import/tests/test_archive_service.py#L94) | Marqueur d'une autre année refusé. |
| `test_missing_ownership_property_is_not_owned` | [102](../services/pst-import/tests/test_archive_service.py#L102) | Propriété de marqueur absente. |
| `test_ownership_check_propagates_non_404_graph_errors` | [110](../services/pst-import/tests/test_archive_service.py#L110) | Les autres erreurs Graph restent visibles. |
| `test_marked_archive_folder_can_be_deleted` | [119](../services/pst-import/tests/test_archive_service.py#L119) | Suppression d'un dossier marqué. |
| `test_missing_archive_folder_is_a_no_op` | [135](../services/pst-import/tests/test_archive_service.py#L135) | Aucun effet si le dossier est absent. |
| `test_unload_requires_explicit_confirmation` | [37](../services/pst-import/tests/test_function_app.py#L37) | Confirmation nécessaire avant suppression. |
| `test_confirmed_unload_reaches_archive_service` | [47](../services/pst-import/tests/test_function_app.py#L47) | Transmission de la demande confirmée. |
| `test_import_mime_encodes_payload_and_target` | [28](../services/pst-import/tests/test_graph_import.py#L28) | Encodage et destination de l'import. |
| `test_import_rejects_bad_response_and_ambiguous_network_failures` | [42](../services/pst-import/tests/test_graph_import.py#L42) | Refus des réponses et erreurs ambiguës. |
| `test_import_reports_http_error_without_exposing_response` | [52](../services/pst-import/tests/test_graph_import.py#L52) | Erreur HTTP sans exposer la réponse. |
| `test_converts_message_body_headers_and_attachment` | [10](../services/pst-import/tests/test_pst_to_mime.py#L10) | Conversion du corps, des en-têtes et de la pièce jointe. |
| `test_rejects_truncated_attachment` | [48](../services/pst-import/tests/test_pst_to_mime.py#L48) | Refus d'une pièce jointe incomplète. |
| `test_missing_attachment_name_uses_placeholder` | [71](../services/pst-import/tests/test_pst_to_mime.py#L71) | Nom de remplacement si absent. |
| `test_rejects_missing_addresses_and_rtf_only_messages` | [85](../services/pst-import/tests/test_pst_to_mime.py#L85) | Refus des adresses absentes et du corps RTF seul. |

Autres fichiers techniques :

- [`archive/archive.js`](../archive/archive.js#L8) et
  [`ui/archive.js`](../ui/archive.js#L8) : maquettes à états fixes.
  Chacun a `getStatusClass` ligne 8, `createRow` ligne 19 et `render` ligne 31.
  Leurs fichiers HTML et CSS gèrent l'affichage ; ce ne sont pas les routes PST.
- [`commandsbak.js`](../src/commands/commandsbak.js#L1) :
  ancienne fonction `openArchiveDialog`. Elle ouvre l'ancienne maquette.
- [`dialog2024.html`](../src/commands/dialog2024.html#L32) :
  ancien dialogue ; `render` ligne 32 et `callAgent` ligne 38.
  Son affichage change avant confirmation : ne pas le présenter comme la
  logique actuelle.
- [`taskpane.backup.js`](../src/taskpane/taskpane.backup.js#L17) :
  exemple ancien ; `run` affiche l'objet du message. Les fichiers `backup`
  ne sont pas le parcours actuel.
- Les fichiers CSS, images, anciens manifestes, JSON de configuration et
  workflows GitHub ne contiennent pas de fonctions métier de lecture PST.
  Ils servent au style, aux ressources, au déploiement ou aux essais.

## 10. Présenter son rôle et l'aide de l'IA

À adapter selon le travail réellement réalisé. Ne pas attribuer une ligne ou
une décision à une personne sans pouvoir le confirmer.

**Pour expliquer son rôle de responsable :**

> « J'ai porté le besoin métier : le choix des années, le parcours utilisateur,
> les contraintes de confidentialité et la démo. J'ai orienté les choix
> techniques, vérifié les limites et fixé la suite avec l'équipe. »

**Pour expliquer l'aide de l'IA :**

> « J'ai utilisé une IA comme assistant pour explorer le code, proposer des
> modifications, rédiger cette documentation et préparer des tests. Les
> résultats restent à relire et à valider sur le poste et avec la DSI. »

**Questions probables :**

- **Pourquoi Outlook classique ?** C'est lui qui monte et expose aujourd'hui le
  PST à l'agent via COM.
- **Le web monte-t-il le PST ?** Non. Le marqueur web n'est pas un montage.
- **La source est-elle écrasée ?** Non. Une copie locale est créée ou réutilisée.
- **La copie est-elle synchronisée ?** Non, pas automatiquement.
- **Le lecteur autonome est-il livré ?** Non. Il faut encore le développer,
  le sécuriser et le tester.
- **Un `.exe` règle-t-il le problème ?** Non. Il faut remplacer la lecture COM ;
  il faut aussi une autorisation Windows pour exécuter le programme.
