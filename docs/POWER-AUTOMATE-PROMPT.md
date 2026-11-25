# Flux Power Automate « MesArchives-Statut-PST » (Outlook Web)

Le flux stocke l'état web sous forme de **fichier marqueur** OneDrive : présent = vert, absent = rouge.
Aucune liste SharePoint n'est nécessaire. Outlook Web ne peut pas monter un PST : le vert signifie uniquement « marqueur posé ». Le client lourd continue d'utiliser l'agent local et le vrai montage PST.

## Déclencheur
« Lorsqu'une requête HTTP est reçue » : Who can trigger = **Anyone**, méthode **vide** (GET et POST), schéma vide.
Régénérer l'URL signée si elle a été exposée, puis la reporter dans `src/archives-config.js` (`webStatusEndpoint`).

## Variables (Initialiser une variable)
| Nom | Type | Valeur |
|---|---|---|
| varBody | Objet | `if(empty(string(triggerBody())), json('{}'), json(string(triggerBody())))` |
| varYear | Chaîne | `coalesce(triggerOutputs()?['queries']?['year'], string(variables('varBody')?['annee']))` |
| varFileName | Chaîne | `if(equals(variables('varYear'),'2024'),'ADR.pst',concat('ADR_',variables('varYear'),'.pst'))` |
| varIsRead | Booléen | `not(empty(triggerOutputs()?['queries']?['year']))` |

## Condition « Annee valide »
`and(greaterOrEquals(int(variables('varYear')),2021), lessOrEquals(int(variables('varYear')),2024))` doit être vrai.
Si non : Réponse 400, corps `{"ok":false,"error":"annee invalide"}`.
Supprimer toute condition « Action valide » (l'add-in n'envoie jamais `action`).

## Branche Oui
1. OneDrive « Lister les fichiers du dossier » (dossier `/Personnel`) ; noms d'action sans accents.
2. « Filtrer le tableau » **Filtrer_PST** sur la sortie de la liste : `item()?['Name']` est égal à `variables('varFileName')`.
3. « Filtrer le tableau » **Filtrer_marqueur** : `item()?['Name']` est égal à `concat('monte_', variables('varYear'), '.flag')`.
4. Condition `variables('varIsRead')` est égal à `true`. Les requêtes POST du complément envoient le JSON avec `Content-Type: text/plain;charset=UTF-8` pour éviter un préflight CORS ; le corps reste du JSON et `varBody` le convertit avec l'expression ci-dessus.

### Oui (GET)
Réponse 200, corps :
`json(concat('{"fileExists":', string(greater(length(body('Filtrer_PST')), 0)), ',"mounted":', string(greater(length(body('Filtrer_marqueur')), 0)), '}'))`

### Non (POST) : condition `variables('varBody')?['Monte']` est égal à `true`
- Oui : si `length(body('Filtrer_marqueur'))` est égal à 0, OneDrive « Créer un fichier » : dossier `/Personnel`, nom `concat('monte_', variables('varYear'), '.flag')`, contenu `1`.
- Non : « Appliquer à chacun » sur `body('Filtrer_marqueur')` avec « Supprimer le fichier » sur `items('Appliquer_a_chacun')?['Id']`.
- Réponse 200, corps `{"ok":true}`.

## En-têtes de **toutes** les réponses (200 et 400)
`Content-Type: application/json` et `Access-Control-Allow-Origin: https://dellasiegaps.github.io`.
Autoriser les méthodes GET et POST et l'en-tête `Content-Type` dans la configuration CORS du flux.
Les booléens s'écrivent avec `string(...)`, jamais `toLower`.

Renseigner l'URL signée dans `src/archives-config.js` sous `webStatusEndpoint`. Cette URL est nécessaire au complément web et sera visible dans le navigateur : limiter les droits du flux aux marqueurs non sensibles et régénérer l'URL si elle est exposée. Ne jamais y intégrer un secret Microsoft 365 ou Azure.

## Test
Ouvrir `<URL du flux>&year=2024` : attendre `{"fileExists":true,"mounted":false}`.
