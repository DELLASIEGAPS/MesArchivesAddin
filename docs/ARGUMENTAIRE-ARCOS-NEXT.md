# Argumentaire pour Pascal : accepter l'agent PST local (ARCOS / NEXT)

**Objectif** : que la vue « Mes Archives » soit identique dans Outlook classique et Outlook Web (outlook.cloud.microsoft) sur les postes de la Banque de France.

## Constat
- Sur les postes BdF, la vue Outlook Web est grisée (« État inconnu / Vérification impossible ») : la chaîne OneDrive cloud → Outlook Web → Entra (applications d'entreprise) → Azure Functions / Power Automate est lente, voire en échec.
- Le montage d'un PST est une opération **locale** à Outlook ; le faire transiter par le cloud ajoute latence, dépendances et points de rupture.

## Arguments de scalabilité
1. **Moins de dépendances** : l'agent PST (127.0.0.1) ne dépend ni d'Entra, ni d'Azure Functions, ni de Power Automate pour lire/monter/démonter un PST.
2. **Performance** : réponse locale en millisecondes, contre plusieurs appels HTTP distants par archive (4 années = 4+ aller-retours cloud à chaque actualisation).
3. **Charge maîtrisée** : chaque poste traite sa propre demande ; aucun goulot d'étranglement central (quotas Power Automate/Functions, jetons Entra) quand le nombre d'utilisateurs augmente.
4. **Expérience identique** : même code, mêmes états (Rouge/Vert, « Charger/Décharger ») sur Outlook classique et Web ; le service cloud ne sert plus que de **secours** et pour vérifier la présence du PST dans OneDrive.
5. **Sécurité** : agent en écoute sur boucle locale uniquement, aucun secret Azure dans le code client.

## Demande
Valider côté **ARCOS** et **NEXT** : déploiement de l'agent PST et autorisation de l'accès à `http://127.0.0.1:45731` depuis la page du complément (Edge/Chrome, accès au réseau local), en attendant l'optimisation du traitement via OneDrive cloud / Outlook Web / Azure.
