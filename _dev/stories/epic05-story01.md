# Story 1: Exceptions COM pour la gestion des erreurs Excel

**Epic:** Epic 5 - Gestion du cycle de vie Excel
**Priorité:** Haute
**Statut:** À faire

## Description
Créer les exceptions spécifiques pour la gestion des erreurs COM liées à Excel. Ces exceptions doivent être ajoutées dans le fichier `src/xlmanage/exceptions.py` et exportées dans `__init__.py`.

## Critères d'acceptation
1. Les exceptions suivantes doivent être implémentées :
   - `ExcelConnectionError` : Erreur de connexion COM (Excel non installé, serveur COM indisponible)
   - `ExcelInstanceNotFoundError` : Instance Excel demandée introuvable
   - `ExcelRPCError` : Erreur RPC (serveur COM déconnecté ou indisponible)

2. Chaque exception doit inclure des attributs pertinents pour le diagnostic :
   - `ExcelConnectionError` : `hresult` et `message`
   - `ExcelInstanceNotFoundError` : `instance_id` et `message`
   - `ExcelRPCError` : `hresult` et `message`

3. Les exceptions doivent être ajoutées à la liste `__all__` dans `src/xlmanage/__init__.py`.

## Tâches
- [ ] Implémenter les exceptions dans `src/xlmanage/exceptions.py`
- [ ] Ajouter les exceptions à `__all__` dans `src/xlmanage/__init__.py`
- [ ] Vérifier que les exceptions sont correctement importées et utilisables

## Dépendances
Aucune dépendance externe. Cette story est la première de l'Epic 5 et doit être complétée avant les autres stories.

## Notes
Les exceptions doivent suivre le format des exceptions existantes dans le projet et inclure des messages d'erreur clairs et informatifs.