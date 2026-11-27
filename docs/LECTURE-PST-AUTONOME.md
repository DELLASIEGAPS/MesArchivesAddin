# Lecteur PST autonome — prototype local

## Principe

Le script `tools\local-pst-lecture-autonome.ps1` lit directement les fichiers
PST avec XstReader.Api, sans utiliser Outlook COM. Il sert le visualiseur sur
`http://127.0.0.1:45732/view`. Il ne monte pas les PST dans Outlook.
L'agent classique reste séparé, sur le port 45731.

Le lecteur ouvre les fichiers dans le dossier passé en paramètre. Il ne crée ni
ne remplace de copie. Utiliser des PST fictifs pour les essais. Ne pas ouvrir un
PST qui est monté ou en cours de modification dans Outlook : la bibliothèque
peut le trouver verrouillé. Ne pas remplacer les copies existantes.

## Préparation

Ce prototype utilise PowerShell 7 et une version corrigée de XstReader.Api.
Le paquet NuGet 1.0.6 ne remonte pas les erreurs de lecture ; il est refusé.
La révision amont est `da80989623f60685beaf5132948d32a3da48fd60`.
Projet : <https://github.com/iluvadev/XstReader>. Licence : MS-PL.
Le script de construction applique deux corrections : nettoyage synchrone
avant fermeture du fichier et passage de `System.Security.Cryptography.Pkcs`
de 7.0.1 à 8.0.2. Il ajoute un marqueur vérifié au démarrage.
Faire valider ces changements par la DSI. Le SDK .NET et l'accès au flux NuGet
doivent respecter les règles de l'organisation.

Le dépôt XstReader doit être propre et placé sur la révision indiquée. Puis :

```powershell
pwsh -NoProfile -File C:\DEV\MesArchivesAddin\tools\build-pst-moteur-autonome.ps1 `
  -SourceDirectory C:\chemin\XstReader
```

Le chemin de la DLL est donné explicitement au démarrage :

```powershell
pwsh -NoProfile -File C:\DEV\MesArchivesAddin\tools\local-pst-lecture-autonome.ps1 `
  -PstSourceDirectory C:\chemin\pst-fictifs `
  -LibraryPath C:\chemin\XstReader\src\XstReader.Api\bin\Release\netstandard2.0\XstReader.Api.dll
```

Garder la fenêtre ouverte. Ouvrir `http://127.0.0.1:45732/view` sur ce PC.
Arrêter avec Ctrl+C dans cette fenêtre. Les noms attendus restent
`ADR_2021.pst`, `ADR_2022.pst`, `ADR_2023.pst` et `ADR.pst`.

Dans la version de développement du complément, choisir **Lecteur PST autonome
(expérimental)** dans le dialogue (sélectionné par défaut), puis afficher le
visualiseur ici ou dans un onglet séparé. Le volet propose aussi un lien vers
l'onglet. Dans ce mode, **Charger** vérifie que le PST peut être ouvert et
autorise sa lecture ; **Décharger** retire cette autorisation sans supprimer ou
modifier le fichier. L'état est partagé par les fenêtres connectées au même
agent et remis à zéro à son redémarrage. Le PST n'apparaît pas dans les dossiers
natifs d'Outlook. Le mode classique reste séparé et conserve ses anciennes actions.
Dans l'onglet séparé, les boutons du visualiseur permettent aussi de charger ou
décharger une archive.
Décharger n'efface pas les pièces jointes déjà téléchargées, les brouillons
préparés ou les données déjà affichées ou copiées ; ce n'est pas un mécanisme
de sécurité contre les autres logiciels du poste.

Le dialogue communique avec le lecteur via le visualiseur local intégré.
Si le navigateur ou Outlook bloque cet affichage, les boutons du dialogue ne
fonctionnent pas : démarrer l'agent puis utiliser les boutons de l'onglet local
séparé. Le dialogue et le lecteur doivent tourner sur le même PC ; cette solution
ne donne pas accès aux PST depuis un autre appareil. Le dialogue publié doit
également être déployé avant que ces changements soient visibles.
L'affichage intégré doit encore être validé dans chaque client Outlook, après
autorisation de la DLL et déploiement du complément. Si l'intégration est bloquée
par le client, utiliser l'onglet séparé sur le même PC.

## Limites importantes

- **Prototype non prêt pour la production.** Le raccordement aux boutons
  Charger/Décharger et l'intégration au complément restent à valider en conditions réelles.
- Le PST est lu directement et doit rester stable pendant la lecture. Les
  identifiants sont invalidés quand sa taille ou sa date de modification change.
- La lecture du corps est limitée au texte brut. Un corps HTML ou RTF est signalé
  comme non pris en charge. Les adresses peuvent être incomplètes ; les messages
  joints à d'autres messages ne sont pas téléchargeables comme fichiers.
- Le moteur peut aussi retourner des éléments qui ne sont pas des courriels.
  Leur restitution n'est pas garantie.
- Les pièces jointes sont chargées en mémoire. Les tailles annoncées par le
  moteur ne sont pas assez fiables : le visualiseur indique « taille non vérifiée ».
  Les gros PST et grosses pièces jointes n'ont pas été qualifiés.
- La recherche peut être longue. Elle porte sur les archives chargées et présentes localement.
- L'accès HTTP est local et exige un jeton et une origine locale pour les appels
  de lecture. Ce n'est pas une authentification des logiciels présents sur le
  poste. L'intégration autorise aussi l'origine `https://dellasiegaps.github.io`
  à afficher la page locale. Elle doit donc être considérée comme de confiance.
  Cette page relaie uniquement les demandes de statut, de chargement et de
  déchargement du dialogue de cette origine ; le jeton reste dans la page locale.
  La sécurité, les licences et la bibliothèque doivent être validées avant
  utilisation avec des archives sensibles.
- Il n'y a pas d'exécutable livré. La stratégie Windows qui a bloqué un essai
  précédent doit être traitée avec l'administrateur. Elle a également bloqué
  le chargement de la DLL corrigée sur le poste de développement (0x800711C7).
  Le test complet après correction n'a donc pas pu être réalisé sur ce poste.
- La dépendance cryptographique amont 7.0.1 signale une vulnérabilité connue.
  Le script utilise 8.0.2 lors de la construction. Le poste de test charge une
  version fournie par PowerShell 7. Cela ne constitue pas une validation de
  sécurité : vérifier toute la chaîne avant usage avec des données sensibles.

## Test reproductible

Les tests utilisent un PST synthétique fourni séparément ; ne pas enregistrer
de PST métier dans le dépôt. Le scénario vérifie le parcours de dossiers, un
message, une pièce jointe et le refus de requêtes non autorisées.

```powershell
$env:MESARCHIVES_TEST_PST = 'C:\chemin\alpha-beta-gamma-delta.pst'
$env:MESARCHIVES_XST_DLL = 'C:\chemin\XstReader.Api.dll'
node --test tools\local-pst-lecture-autonome.test.js tools\local-pst-view.test.js
```

Le fichier de test utilisé pendant le développement provient du jeu synthétique
[pst-viewer](https://github.com/bod09/pst-viewer/tree/f9c27d12209f09afba7e41f1426bca871b9c7540/samples/pst-extractor),
sous la licence indiquée dans ce dossier. Le test vérifie le message `Alpha`.
Si les deux variables ne sont pas définies, le test d'intégration est ignoré.
