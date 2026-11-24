# Visualiseur local Mes Archives

## Ce que fait cette version

Depuis la fenêtre « Mes Archives » dans Outlook Web, **Afficher le visualiseur ici** intègre le visualiseur dans le dialogue. **Ouvrir dans un onglet séparé** conserve l'accès direct à `http://127.0.0.1:45731/view`. Les deux modes utilisent le même visualiseur local. Il faut démarrer l'agent avant d'utiliser l'une ou l'autre option. Cette adresse désigne toujours le poste sur lequel le navigateur est ouvert : elle n'est pas accessible depuis un autre ordinateur.

Le visualiseur permet de naviguer dans les dossiers et sous-dossiers d'un PST monté, de lire les messages et de télécharger leurs pièces jointes. Les messages sont affichés en **texte brut** : la mise en forme HTML et les images intégrées ne sont pas restituées. Les éléments autres que les messages (par exemple les rendez-vous) ne sont pas affichés. Les listes sont paginées par 50 éléments du dossier ; une page peut être vide si elle ne contient aucun message. Cette version nécessite toujours Outlook **classique** ouvert sur le même PC ; Outlook Web ou le nouvel Outlook ne remplace pas Outlook classique pour le montage. Il ne s'agit pas d'un lecteur PST autonome ni d'un import dans Exchange.

La recherche parcourt les champs disponibles des messages : objet, expéditeur et adresses résolues, destinataires À, Cc et Cci lorsqu'ils sont disponibles, date, corps en texte brut et noms des pièces jointes. Elle ignore la casse et les accents, mais ne lit **pas le contenu des pièces jointes** ni les éléments autres que les messages. Elle explore les dossiers des quatre archives **montées** ; les années non montées sont indiquées et exclues. Elle peut être longue sur des PST volumineux ; **Arrêter** interrompt le parcours après la requête en cours et conserve des résultats partiels. Une erreur signale que les résultats sont incomplets.

Les boutons **Répondre**, **Répondre à tous** et **Transférer** préparent des champs modifiables (À, Cc, Cci, objet et texte) dans le visualiseur. **Ouvrir le brouillon dans Outlook Web** utilise une URL contenant ces champs : ils peuvent apparaître dans l'historique ou des journaux. Faire valider cet usage pour des données sensibles. Aucun envoi n'est automatique ; vérifier tous les champs dans Outlook Web avant l'envoi. Le préremplissage dépend du navigateur et d'Outlook Web ; si l'URL dépasse 2 000 caractères, ou si le préremplissage échoue, ouvrir Outlook Web sans préremplissage et copier les champs manuellement. Il ne s'agit pas d'une réponse liée au fil d'origine ; les pièces jointes doivent être téléchargées et ajoutées manuellement. La réponse utilise l'adresse de l'expéditeur, sans prendre en compte une éventuelle adresse « Répondre à » différente. Les adresses non résolues doivent être complétées et l'adresse de l'utilisateur peut figurer parmi les destinataires de « Répondre à tous ». Les Cci d'origine ne sont pas disponibles. Le brouillon préparé dans le visualiseur n'est pas enregistré de manière persistante.

La consultation ne modifie pas volontairement les messages, mais Outlook peut mettre à jour des données internes du PST monté. Les pièces jointes sont enregistrées temporairement sur le poste par l'agent puis supprimées après transmission ; le navigateur les enregistre dans son dossier de téléchargement selon ses paramètres. Appliquer les règles de confidentialité et de sécurité de l'organisation avant de télécharger ou d'ouvrir une pièce jointe. Les informations des archives restent sensibles : le site GitHub Pages autorisé et les logiciels locaux capables d'accéder à l'agent doivent être considérés comme de confiance.

L'agent utilise une copie de travail dans `%LOCALAPPDATA%\MesArchives\PST`. Si elle existe déjà, elle est réutilisée et n'est pas automatiquement actualisée depuis la source : ne pas la supprimer ou la remplacer sans vérifier son contenu et fermer le PST dans Outlook.

## Préparer les PST

Noms attendus : `ADR_2021.pst`, `ADR_2022.pst`, `ADR_2023.pst` et `ADR.pst` pour 2024. Placer les fichiers dans un même dossier local, par exemple `C:\DEV`, ou dans un dossier OneDrive synchronisé sur ce PC.

Pour OneDrive, rendre les fichiers disponibles sur le poste dans l'Explorateur Windows (clic droit, **Toujours conserver sur cet appareil**) et attendre la fin de la synchronisation. Une URL OneDrive Web ou SharePoint n'est pas un chemin local : l'agent ne télécharge pas un PST à partir d'une URL. La synchronisation OneDrive peut transférer les fichiers vers le cloud selon la configuration de l'organisation ; faire valider ce choix par l'administrateur.

## Démarrer l'agent

Ces commandes s'appliquent à la branche `client-lourd-seul` contenant ce visualiseur, sur le PC de l'utilisateur. Elles lancent explicitement **Windows PowerShell 5.1**, nécessaire pour détecter Outlook classique. Ne pas lancer le script directement dans PowerShell 7 (`pwsh`). Ouvrir Outlook classique avant de charger une archive.

Pour les fichiers dans `C:\DEV` :

```powershell
powershell.exe -NoProfile -File 'C:\DEV\MesArchivesAddin\tools\local-pst-agent.ps1' -PstSourceDirectory 'C:\DEV'
```

Pour le dossier synchronisé `Personnel` de OneDrive professionnel, sans inscrire le nom Windows en dur :

```powershell
powershell.exe -NoProfile -File 'C:\DEV\MesArchivesAddin\tools\local-pst-agent.ps1' -PstSourceDirectory (Join-Path $env:OneDriveCommercial 'Personnel')
```

Si `$env:OneDriveCommercial` n'est pas défini ou si le dossier réel est différent, utiliser le chemin local indiqué dans l'Explorateur. Avec `-PstSourceDirectory`, l'agent cherche uniquement dans ce dossier, sans parcourir ses sous-dossiers. Sans ce paramètre, il recherche dans les racines OneDrive détectées, y compris leurs sous-dossiers ; en présence de plusieurs copies du même nom, définir explicitement le dossier pour choisir la bonne.

Laisser la fenêtre PowerShell ouverte. Pour arrêter l'agent, y appuyer sur **Ctrl+C**. Si le port est déjà utilisé, arrêter seulement l'agent existant identifié avant de relancer.

## Charger et voir une archive

1. Dans la fenêtre « Mes Archives », sélectionner **Charger** pour l'année voulue. L'agent crée une copie locale de travail si nécessaire puis la monte dans Outlook classique.
2. Dans le dialogue, cliquer sur **Afficher le visualiseur ici**, choisir l'année, puis **Afficher**.
3. Le visualiseur présente les dossiers à gauche, la liste des messages au centre et le volet de lecture à droite, comme dans Outlook classique. Cliquer sur un dossier, puis sur un message pour le lire sans masquer la liste. Utiliser le chemin des dossiers ou **Dossier parent**, et les boutons de pagination pour naviguer. Sur une fenêtre étroite, le volet de lecture passe sous les deux listes. Les adresses SMTP résolues sont cliquables et préparent un nouveau message ; sinon le nom est affiché avec « adresse indisponible ». Une pièce jointe est téléchargée uniquement après un clic sur son nom.
4. Pour utiliser la page complète du navigateur, cliquer sur **Ouvrir dans un onglet séparé**. Si le complément affiché n'inclut pas encore ces commandes, ouvrir manuellement `http://127.0.0.1:45731/view` dans un nouvel onglet sur le même PC. Ne pas ouvrir directement le fichier HTML en `file://` : les appels à l'agent ne fonctionneront pas correctement.

Les modifications du complément sur la branche de développement ne deviennent visibles dans Outlook Web qu'après déploiement de cette version. Après toute modification du script ou de la page du visualiseur, arrêter l'ancienne instance de l'agent puis relancer la commande ci-dessus ; garder les fichiers `tools\local-pst-agent.ps1` et `tools\local-pst-view.html` ensemble. La branche de développement n'est pas automatiquement publiée sur GitHub Pages.

Si l'agent est injoignable, vérifier qu'il est démarré, que le navigateur est sur le même PC et que les règles du navigateur/de l'entreprise autorisent l'accès local. Si le PST est indiqué comme non monté, vérifier Outlook classique puis le charger. Si la source est « uniquement en ligne », terminer d'abord la synchronisation OneDrive.

La progression affiche les étapes observées par le complément et demande de patienter après une seconde. Elle ne représente pas un pourcentage de copie ni une estimation de durée. Le bouton **Fermer** demande au complément de fermer sa fenêtre. **Fermer le visualiseur** demande de masquer la vue intégrée ; si elle reste visible, utiliser **Masquer le visualiseur** dans la fenêtre Mes Archives. Dans l'onglet séparé, le navigateur peut refuser la fermeture automatique : utiliser alors la croix de l'onglet. La fermeture ne démonte pas l'archive ; utiliser **Décharger** pour cela.

L'état des archives est vérifié à l'ouverture du dialogue et lorsqu'il redevient visible, dans Outlook Web comme dans Outlook classique ; le bouton **Actualiser** reste disponible. Un retour pendant une opération déclenche une vérification après sa fin. Le bouton **Masquer le visualiseur** replie le visualiseur intégré sans fermer le dialogue.

Le visualiseur ne copie pas les messages vers Exchange ou Graph. Le PST d'origine peut cependant être synchronisé par OneDrive et les téléchargements peuvent être placés dans un dossier synchronisé ; ces opérations sont indépendantes de ce visualiseur.
