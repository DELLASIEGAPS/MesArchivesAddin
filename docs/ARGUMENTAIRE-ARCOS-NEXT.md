# Argumentaire pour Pascal : accepter l'agent PST local (ARCOS / NEXT)

**Objectif** : que l'interface « Mes Archives » reste utilisable dans Outlook classique et Outlook Web (outlook.cloud.microsoft), avec une action adaptée à chaque plateforme.

## Constat
- Sur les postes BdF, la vue Outlook Web est grisée (« État inconnu / Vérification impossible ») : la chaîne OneDrive cloud → Outlook Web → Entra (applications d'entreprise) → Azure Functions / Power Automate est lente, voire en échec.
- Le montage d'un PST est une opération **locale** à Outlook classique. Outlook Web ne peut ni monter un PST ni lire ses messages via ce flux.

## Arguments de scalabilité
1. **Moins de dépendances** : l'agent PST (127.0.0.1) ne dépend ni d'Entra, ni d'Azure Functions, ni de Power Automate pour lire/monter/démonter un PST.
2. **Performance** : réponse locale en millisecondes, contre plusieurs appels HTTP distants par archive (4 années = 4+ aller-retours cloud à chaque actualisation).
3. **Charge maîtrisée** : chaque poste traite sa propre demande ; aucun goulot d'étranglement central (quotas Power Automate/Functions, jetons Entra) quand le nombre d'utilisateurs augmente.
4. **Actions distinctes et claires** : le client lourd conserve « Charger/Décharger » via l'agent local ; le web utilise « Activer côté web / Retirer le marqueur web » via Power Automate. Le vert côté web indique un marqueur cloud, jamais un PST monté.
5. **Sécurité** : agent en écoute sur boucle locale uniquement, aucun secret Azure dans le code client.

## Demande
Valider côté **ARCOS** et **NEXT** : déploiement de l'agent PST et autorisation de l'accès à `http://127.0.0.1:45731` depuis le complément Outlook classique. Les actions Outlook Web utilisent séparément le flux cloud et ne doivent pas dépendre de cet agent.
