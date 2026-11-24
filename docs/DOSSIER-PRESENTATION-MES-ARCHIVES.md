# Mes Archives — dossier de présentation

Version de travail du 23 novembre 2026. Ce document décrit le prototype de la branche `client-lourd-seul`, pas une solution validée pour la production.

## 1. Synthèse pour les décideurs

**Besoin.** Permettre à un utilisateur de retrouver et consulter ses archives Outlook au format PST, par année, avec une interface familière et sans importer automatiquement les messages dans Exchange.

**Réponse proposée.** Un complément Outlook ouvre une fenêtre « Mes Archives ». Un agent exécuté sur le poste, avec Outlook classique, monte une copie locale du PST et fournit un visualiseur à trois volets : dossiers, liste des messages et lecture. Les pièces jointes sont téléchargeables à la demande.

**Décision demandée.** Valider un pilote restreint et supervisé, avec données de test puis données autorisées, après revue sécurité, conformité, poste de travail et exploitation. La généralisation ne doit pas être décidée sur la seule base de ce prototype.

**Valeur attendue.** Interface homogène, conservation des archives hors boîte Exchange, lecture rapide sur le poste. **Coûts et dépendances.** Outlook classique, agent local Windows, PST disponibles sur le poste, assistance et gestion des mises à jour.

## 2. Fonctionnement pour l'utilisateur

1. Démarrer Outlook classique et l'agent local sur le même poste.
2. Ouvrir « Mes Archives » dans Outlook. Pour chaque année, l'indicateur signale disponible, indisponible ou état inconnu. Si la vérification échoue, l'état reste inconnu : il ne faut pas le confondre avec un PST absent.
3. Charger une archive : l'agent recherche le fichier source, crée si nécessaire une copie locale de travail, la monte dans Outlook, puis vérifie l'état obtenu. Décharger démonte la copie de travail.
4. Afficher le visualiseur intégré ou ouvrir son onglet séparé. Sélectionner l'année, un dossier puis un message. Le corps apparaît en texte brut ; cliquer sur une pièce jointe demande son téléchargement au navigateur.
5. Si une adresse SMTP est résolue, cliquer sur l'expéditeur ou un destinataire prépare un nouveau message dans l'application de messagerie par défaut. Une adresse Exchange non résolue est signalée comme indisponible ; aucun message n'est envoyé automatiquement.
6. La progression présente les étapes réellement observables par l'interface. Après une seconde, elle demande de patienter. La barre représente les étapes terminées, **pas** le pourcentage d'octets copiés ni une durée estimée.

Le parcours détaillé et les commandes de démarrage figurent dans [VISUALISEUR-LOCAL.md](./VISUALISEUR-LOCAL.md).

### Limites fonctionnelles

- Messages uniquement : les rendez-vous, contacts et autres éléments ne sont pas affichés. Listes de 50 éléments du dossier par page ; les éléments non-messages comptent dans la pagination.
- Pas de rendu HTML ou d'images intégrées. La recherche parcourt les messages des archives montées et peut être longue ; les archives non montées sont exclues. Les actions de réponse et de transfert préparent un nouveau message Outlook Web via une URL contenant ses champs : ni le fil ni les pièces jointes ne sont repris automatiquement. Les données de l'URL peuvent apparaître dans l'historique ou des journaux.
- La résolution SMTP dépend des informations présentes dans Outlook et, pour certains correspondants Exchange, de l'annuaire accessible. Elle peut être indisponible ou déclencher des restrictions de sécurité Outlook.
- Le visualiseur est en lecture seule fonctionnelle. Outlook peut néanmoins modifier des données internes du PST monté. Il ne faut pas interpréter cela comme une garantie d'immuabilité.
- La copie locale de travail est réutilisée et n'est pas automatiquement mise à jour si la source change. Les autres applications locales et les téléchargements restent soumis aux règles de sécurité du poste.

## 3. Architecture technique

### Vue des composants et des flux

```mermaid
flowchart LR
    U[Utilisateur]
    GH[GitHub Pages<br/>Manifeste et code du complément]
    subgraph PC["Poste Windows de l'utilisateur"]
        OWA[Outlook Web ou classique<br/>Complément et dialogue]
        V[Visualiseur local<br/>Intégré ou onglet séparé]
        A[Agent local<br/>Windows PowerShell 5.1]
        OL[Outlook classique ouvert<br/>Object Model COM]
        S[(PST source local<br/>ou OneDrive synchronisé)]
        W[(Copie de travail PST<br/>LOCALAPPDATA / MesArchives / PST)]
        D[Téléchargements<br/>du navigateur]
    end
    U --> OWA
    GH -->|HTTPS : chargement du complément| OWA
    OWA -->|HTTP local : état, charger, décharger| A
    OWA -->|Affiche /view ou ouvre un onglet| V
    A -->|HTTP local : fournit /view| V
    V -->|HTTP local : dossiers, messages, pièces jointes| A
    A -->|Copie initiale si absente| W
    S -->|Lecture de la source| A
    A -->|COM : montage, démontage, lecture| OL
    OL <-->|Accès au PST monté| W
    V -->|Pièce jointe demandée| D
```

Le code du dialogue téléchargé depuis le site HTTPS s'exécute dans le navigateur et appelle l'agent sur le même poste via la boucle locale. Le schéma représente les dépendances, pas une publication de messages vers le site. Une source OneDrive doit être synchronisée et disponible localement ; la copie de travail existante n'est pas actualisée automatiquement.

### Parcours de chargement et de consultation

```mermaid
sequenceDiagram
    actor U as Utilisateur
    participant C as Dialogue Mes Archives
    participant A as Agent local
    participant S as PST source
    participant W as Copie de travail
    participant O as Outlook classique
    participant V as Visualiseur local

    U->>C: Choisit une année et Charger
    C->>A: État puis demande de montage
    A->>W: Recherche la copie existante
    alt Copie absente
        A->>S: Recherche le PST source disponible
        A->>W: Crée la copie locale
    else Copie présente
        Note over A,W: Réutilisation sans actualisation depuis la source
    end
    A->>O: Monte la copie via COM
    C->>A: Vérifie l'état
    U->>C: Ouvre le visualiseur intégré ou séparé
    C->>V: Affiche ou ouvre /view
    U->>V: Sélectionne dossier et message
    V->>A: Demande dossiers et message
    A->>O: Lit le PST monté via COM
    O-->>A: Données du dossier et du message
    A-->>V: Réponse locale ; corps en texte brut
    opt Pièce jointe demandée
        U->>V: Clique sur la pièce jointe
        V->>A: Demande le téléchargement
        A->>O: Récupère la pièce jointe
        A-->>V: Transmet le fichier
        Note over A,V: Fichier temporaire côté agent, puis téléchargement navigateur
    end
    opt Déchargement
        U->>C: Choisit Décharger
        C->>A: Demande de démontage
        A->>O: Démonte la copie via COM
        Note over W: La copie locale est conservée
    end
```

Le manifeste identifie le complément et les URL HTTPS ; le code du dialogue et des commandes est hébergé sur GitHub Pages. L'agent écoute par défaut sur `127.0.0.1:45731` ; cette adresse concerne le poste où le navigateur s'exécute. Le dialogue appelle l'agent pour l'état et le montage ; le visualiseur, servi par l'agent, utilise les routes `/folder`, `/message` et `/attachment`.

Le dialogue propose quatre noms de PST configurés ; l'agent vérifie le format du nom reçu, sans limiter les requêtes à ces quatre noms. Il utilise `%LOCALAPPDATA%\MesArchives\PST` pour la copie de travail. Outlook classique doit être déjà ouvert : l'agent s'y attache par COM avec Windows PowerShell 5.1. Une pièce jointe passe temporairement par un fichier local avant transmission puis suppression ; le navigateur choisit ensuite son dossier de téléchargement. Aucun flux d'import Exchange/Graph n'est déclenché par le visualiseur. Le PST source, le code du site, le navigateur, les journaux et les téléchargements peuvent néanmoins relever de politiques de synchronisation ou de conservation distinctes.

### Zoom sur le grand encadré : flux locaux et enjeux ARCOS

Le grand encadré comprend les échanges entre l'**agent local**, le **visualiseur**, **Outlook classique** et la **copie PST**, ainsi que la création initiale de cette copie. L'agent et les fichiers PST sont dessinés aux limites de l'encadré sur la capture, mais sont indispensables à l'explication des flux. Le visualiseur appelle l'agent HTTP local ; l'agent interroge Outlook classique par COM ; Outlook accède à la copie PST montée. Cette architecture ne permet pas de se passer d'Outlook classique.

```mermaid
flowchart TB
    S[(PST source disponible localement)]
    subgraph PC["Poste Windows : zoom sur les flux du grand encadré"]
        C[Dialogue Mes Archives]
        V[Visualiseur dans le navigateur<br/>intégré ou onglet séparé]
        A[Agent local PowerShell 5.1<br/>127.0.0.1:45731]
        O[Outlook classique ouvert<br/>Object Model COM]
        W[(Copie PST de travail<br/>LOCALAPPDATA)]
        T[Fichier temporaire<br/>pièce jointe]
        D[Téléchargement navigateur]
        C -->|HTTP local : état, montage, démontage| A
        A -->|HTTP local : page /view| V
        V -->|HTTP local : /folder, /message, /attachment| A
        A -->|HTTP local : listes, texte, pièce jointe| V
        A -->|COM : monter, démonter, lire| O
        O <-->|Accès au PST monté| W
        A -->|Crée la copie si absente| W
        O -->|Écrit à la demande| T
        T -->|Transmis puis suppression tentée| A
        V -->|Téléchargement sur clic| D
    end
    S -->|Lecture pour copie initiale| A
```

Le PST source est lui aussi local ou synchronisé sur ce poste ; il est représenté hors du cadre du zoom pour distinguer l'origine de la copie de travail. Le parcours des pièces jointes passe par un fichier temporaire avant la réponse de l'agent, puis par le navigateur. Les flèches représentent des échanges logiques observés dans le code, pas une preuve d'homologation ou de fonctionnement sur un poste Banque de France.

**Périmètre ARCOS à confirmer.** Le dépôt mentionne ARCOS et NEXT comme interlocuteurs pour l'acceptation de l'agent et de l'accès local, mais ne décrit ni leur rôle exact ni une interface technique avec ARCOS. Aucun échange de données avec un système ARCOS n'est établi par les éléments disponibles. Les validations ci-dessous sont des propositions à répartir avec les équipes concernées, pas des décisions déjà prises.

#### Flux de données observables dans le prototype

| Étape | Source → destination | Données et traitement | Implications pour la Banque de France |
|---|---|---|---|
| Mise à disposition | Site HTTPS → navigateur ; agent local → navigateur | Le navigateur charge le complément depuis GitHub Pages et le visualiseur `/view` depuis `127.0.0.1:45731`. | Valider l'hébergement du code, son intégrité, les versions compatibles, l'autorisation de l'iframe et l'accès du navigateur à la boucle locale. Le site et son code sont dans le périmètre de confiance. |
| Préparation | PST source local → agent → copie locale | Si la copie n'existe pas, l'agent copie le PST dans `%LOCALAPPDATA%\MesArchives\PST`. Une copie présente est réutilisée sans actualisation ; une source OneDrive doit être disponible sur le poste. | Une seconde copie potentiellement sensible persiste sur le poste, y compris après démontage. Définir classification, droits d'accès, chiffrement, espace disque, durée de conservation, sauvegarde et procédure contrôlée de renouvellement. |
| Montage | Dialogue → agent → Outlook classique → copie PST | Le dialogue transmet le nom du fichier et le nom d'affichage ; l'agent demande à Outlook, via COM, de monter ou démonter la copie. Il renvoie notamment un état et peut renvoyer le chemin local. | Autoriser et gérer l'agent et Outlook classique sur le même poste ; maîtriser les droits du compte Windows, les versions et les éventuelles alertes de sécurité Outlook. Le démontage ne supprime pas la copie. |
| Dossiers et liste | Visualiseur → agent → Outlook classique → visualiseur | Les requêtes `/folder` transportent le nom du PST, l'identifiant du dossier et la position de page. La réponse contient notamment noms et identifiants de dossiers, nombres d'éléments, objets, expéditeurs et dates des messages. | Les métadonnées peuvent elles-mêmes être sensibles. Elles traversent HTTP sur la boucle locale et sont affichées dans le navigateur ; évaluer le risque lié aux processus locaux et aux composants du navigateur. |
| Lecture | Visualiseur → agent → Outlook classique → visualiseur | `/message` utilise l'identifiant du message. La réponse comprend notamment objet, date, expéditeur et destinataires, adresses résolues, corps en texte brut et noms et tailles des pièces jointes. | Le contenu sort du PST vers la mémoire de l'agent et du navigateur. Évaluer confidentialité, exposition locale, captures d'écran, extensions et outils du poste. La résolution d'adresses Exchange peut aussi dépendre de l'annuaire accessible et des politiques Outlook. |
| Pièce jointe | Outlook classique → fichier temporaire → agent → navigateur | Sur clic, `/attachment` demande un message et un index. Outlook écrit un fichier temporaire ; l'agent le transmet puis tente de le supprimer. Le navigateur reçoit les octets et les enregistre selon ses paramètres. | Vérifier les droits et protections du dossier temporaire, la suppression en cas d'interruption, l'antivirus, les contrôles de téléchargement, la conservation et le risque d'un dossier de téléchargement synchronisé ou partagé. |

Les appels du visualiseur vers l'agent sont en HTTP sur `127.0.0.1` : ils ne constituent pas une liaison réseau vers GitHub Pages. **Cela ne garantit pas que toutes les données restent hors du cloud** : la source ou les téléchargements peuvent être synchronisés, Outlook peut contacter l'annuaire, et d'autres logiciels ou politiques du poste peuvent traiter les données. Le code du prototype ne déclenche pas d'import des messages dans Exchange ou Graph ; ne pas en déduire une garantie sur l'ensemble de l'environnement.

#### Décisions à instruire avec ARCOS et les équipes Banque de France

| Enjeu | Questions et actions proposées avant pilote |
|---|---|
| Sécurité des accès | Qui est autorisé à lancer et utiliser l'agent ? Examiner l'absence d'authentification des processus locaux, les routes de montage et de lecture, la portée des contrôles d'origine/CORS et les risques d'accès par d'autres logiciels du poste. Définir les protections complémentaires et les conditions d'homologation. |
| Poste et navigateur | ARCOS et les équipes poste/navigateur doivent confirmer les politiques applicables à Windows PowerShell 5.1, COM, l'écoute locale HTTP, l'accès réseau local depuis une page HTTPS, l'iframe, les extensions et les téléchargements. Tester la configuration réelle plutôt que présumer son autorisation. |
| Données et conformité | Identifier le propriétaire des PST, leur classification et leur durée de conservation ; définir les habilitations, le chiffrement, les règles de copie, de sauvegarde, d'effacement et de traitement des pièces jointes avec les équipes sécurité et conformité. |
| Exploitation et traçabilité | Désigner les responsables de distribution et de mise à jour de l'agent et du site, du support, des incidents et du retour arrière. Définir les événements à tracer et leur conservation sans journaliser inutilement le contenu des messages : le prototype ne fournit pas à lui seul une traçabilité d'audit complète. |
| Pérennité | Vérifier la disponibilité et le support d'Outlook classique et de Windows PowerShell 5.1 sur les postes cibles. Étudier séparément une solution ne dépendant pas de COM ; le visualiseur actuel ne répond pas à cet objectif. |
| Recette et arbitrage | Tester avec des données fictives puis autorisées : Outlook fermé, PST absent ou verrouillé, copie ancienne, politique navigateur restrictive, pièces jointes et arrêt brutal. Consigner les résultats et obtenir les accords requis avant toute généralisation. |

**Livrable attendu de la concertation :** un schéma de flux validé sur le poste cible, une matrice des responsabilités et des autorisations, les mesures de sécurité retenues et une décision explicite sur un pilote. Les propriétaires, destinations de synchronisation et règles effectivement appliquées restent à confirmer par la Banque de France.

### Frontières et sécurité

Le site autorisé et tous les logiciels locaux capables d'appeler l'agent sont dans le périmètre de confiance. CORS, contrôles d'origine sur les routes de lecture et CSP réduisent certaines expositions mais **ne constituent pas une authentification des processus locaux** ; les routes de montage et d'état demandent également une revue dédiée. La communication loopback HTTP, l'intégrité du code GitHub Pages, la distribution de l'agent, les droits du compte, les pièces jointes et les données temporaires doivent être validés par la sécurité avant production.

Le dépôt GitHub est public : aucun PST réel, message, pièce jointe, capture sensible, identifiant ou secret ne doit être publié. Des fichiers PST sont déjà suivis dans l'historique du dépôt : vérifier leur nature et, s'ils contiennent des données réelles, déclencher immédiatement la procédure interne de gestion d'incident ; une simple suppression dans un nouveau commit ne les retire pas de l'historique.

## 4. Enjeux et défis

| Enjeu | Défi / adaptation | Préconisation |
|---|---|---|
| Confidentialité | PST, pièces jointes et copies locales peuvent contenir des données sensibles. | Classification, chiffrement et contrôle d'accès du poste, antivirus, règles de téléchargement et revue de l'agent ; vérifier les journaux et le dépôt public. |
| Pérennité | Dépendance à Outlook classique, COM et Windows PowerShell 5.1. | Cartographier les versions prises en charge ; tester sur postes administrés et prévoir une trajectoire de remplacement. |
| Cohérence | La copie locale peut différer de la source, et un PST ouvert peut être verrouillé. | Procédure de synchronisation, sauvegarde et renouvellement maîtrisé ; ne jamais écraser une copie sans analyse et fermeture. |
| Expérience | Traitements longs et multiples zones de défilement. | Étapes et barre basées sur les événements connus, volet unique de défilement et actions accessibles ; vérifier en conditions réelles. |
| Déploiement | Code web, manifeste et agent local évoluent indépendamment. | Versionner les composants, publier une version testée sur un site contrôlé, distribuer l'agent signé et supervisé, prévoir retour arrière. |
| Migration | PST multiples et potentiels doublons. | Inventaire, volumétrie, conservation et pilote ; comparer maintien local et import gouverné en boîte Exchange. |

## 5. Recommandations et références Microsoft

- **API de dialogue Office.** Utiliser une URL de dialogue HTTPS sur le même domaine initial que le complément ; le dialogue communique avec son parent et c'est le parent qui appelle `Dialog.close()`. Une seule boîte par page hôte ; la fermeture est asynchrone. [Règles et bonnes pratiques de l'API de dialogue](https://learn.microsoft.com/fr-fr/office/dev/add-ins/develop/dialog-best-practices).
- **PST et OneDrive.** Microsoft décrit la copie d'un PST hors du dossier OneDrive et la reconnexion d'Outlook à cette copie. Éviter d'utiliser directement comme fichier actif un PST synchronisé sans validation de la politique de l'entreprise. [Déplacer un fichier de données Outlook hors de OneDrive](https://support.microsoft.com/en-us/onedrive/how-to-remove-an-outlook-pst-data-file-from-onedrive).
- **Nouvel Outlook.** Microsoft documente désormais l'ouverture, la lecture et la recherche d'e-mails PST avec des limites et prérequis variables. Cela ne rend pas cet agent COM compatible avec le nouvel Outlook : valider séparément la version et les scénarios. [Ouvrir des éléments PST](https://support.microsoft.com/en-gb/outlook/open-and-find-items-in-an-outlook-data-file-pst?nochrome=true).
- **Alternative d'entreprise.** Microsoft Purview propose l'import PST vers Microsoft 365 ; c'est une migration avec traitement cloud et gouvernance, non le fonctionnement de ce prototype. Évaluer quotas, sécurité, doublons et conservation avant décision. [Import de fichiers PST](https://learn.microsoft.com/en-us/purview/importing-pst-files-to-office-365) et [FAQ de l'import](https://learn.microsoft.com/en-us/purview/pst-import-faq).

## 6. Déploiement et points à valider

Les changements locaux sur `client-lourd-seul` ne sont pas automatiquement visibles dans la branche `main` ni sur GitHub Pages. Pousser une branche, publier le site et installer le manifeste sont des actions distinctes. Le site et l'agent doivent être issus de versions compatibles ; après mise à jour de l'agent ou du visualiseur local, arrêter et relancer l'agent.

Pour le pilote : tester les quatre années, présence/absence d'Outlook, PST absent ou verrouillé, droits restreints, erreur réseau locale, grands dossiers, pièces jointes, SMTP Exchange non résolu, fermeture du dialogue dans Outlook et dans le navigateur, accessibilité clavier, et comportement réel de GitHub Pages et des politiques entreprise. Le prototype ne permet pas de déduire une homologation Microsoft ou sécurité.

## 7. Archive de démonstration

Le script [create-demo-pst.ps1](../tools/create-demo-pst.ps1) crée un PST `ADR_2021.pst` séparé contenant uniquement un message fictif et une pièce jointe fictive. Outlook enregistre brièvement le message comme brouillon fictif avant son déplacement dans le PST ; aucun message n'est envoyé. Le script refuse de remplacer un fichier existant et nécessite Outlook classique ouvert. Exécuter dans un dossier de démonstration vide, jamais dans le dossier contenant l'archive réelle :

```powershell
powershell.exe -NoProfile -File 'C:\DEV\MesArchivesAddin\tools\create-demo-pst.ps1' -DestinationDirectory 'C:\chemin\demo-vide'
```

Pour l'utiliser via l'agent, lancer celui-ci avec ce dossier comme source **et un répertoire LOCALAPPDATA de test isolé** : sinon la copie de travail 2021 existante sera réutilisée. Exemple dans une nouvelle fenêtre PowerShell, en adaptant les deux chemins de démonstration :

```powershell
$env:LOCALAPPDATA = 'C:\chemin\demo-local-isole'
powershell.exe -NoProfile -File 'C:\DEV\MesArchivesAddin\tools\local-pst-agent.ps1' -PstSourceDirectory 'C:\chemin\demo-vide'
```

Cette variable est modifiée uniquement pour cette fenêtre et ses processus enfants. Ne pas lancer deux agents sur le même port. Outlook peut garder le nouveau PST verrouillé après démontage ; si la copie échoue, fermer normalement Outlook, attendre sa fermeture complète, puis le rouvrir avant de réessayer. Ne jamais supprimer ou remplacer la copie de travail réelle pour le test.
