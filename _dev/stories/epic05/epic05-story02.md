# Story 2: Implémentation du gestionnaire de cycle de vie Excel

**Epic:** Epic 5 - Gestion du cycle de vie Excel
**Priorité:** Haute
**Statut:** Terminé ✅

**Date de complétion:** 2026-02-03
**Version:** 1.0

## Description
Créer le fichier `src/xlmanage/excel_manager.py` qui gérera le cycle de vie des instances Excel. Ce gestionnaire doit permettre de démarrer, arrêter, et lister les instances Excel en cours d'exécution.

## Critères d'acceptation
1. Implémenter la dataclass `InstanceInfo` avec les attributs suivants :
   - `pid` : Process ID du processus EXCEL.EXE
   - `visible` : Indique si l'instance est visible à l'écran
   - `workbooks_count` : Nombre de classeurs ouverts
   - `hwnd` : Handle de fenêtre Windows pour identification unique

2. Implémenter la classe `ExcelManager` avec les méthodes suivantes :
   - `__init__(self, visible: bool = False)` : Initialise le gestionnaire
   - `start(self, new: bool = False) -> InstanceInfo` : Démarre ou se connecte à une instance Excel
   - `get_running_instance(self) -> InstanceInfo | None` : Récupère l'instance Excel active
   - `get_instance_info(self, app: CDispatch) -> InstanceInfo` : Lit les informations d'une instance Excel
   - `list_running_instances(self) -> list[InstanceInfo]` : Énumère toutes les instances Excel actives

3. Implémenter les méthodes de gestion des erreurs :
   - Lever `ExcelConnectionError` si la connexion COM échoue
   - Lever `ExcelInstanceNotFoundError` si une instance demandée n'est pas trouvée

4. Implémenter les fonctions utilitaires pour l'énumération des instances :
   - `enumerate_excel_instances()` : Énumération via le Running Object Table (ROT)
   - `enumerate_excel_pids()` : Fallback pour l'énumération des PIDs via tasklist
   - `connect_by_hwnd(hwnd: int) -> CDispatch | None` : Connexion à une instance Excel par son handle de fenêtre

## Tâches
- [x] Créer le fichier `src/xlmanage/excel_manager.py`
- [x] Implémenter la dataclass `InstanceInfo`
- [x] Implémenter la classe `ExcelManager` avec toutes les méthodes requises
- [x] Implémenter les fonctions utilitaires pour l'énumération des instances
- [x] Tester les fonctionnalités de base (démarrage, connexion, énumération)
- [x] Créer des tests unitaires complets avec mocks COM
- [x] Vérifier la couverture de code (73% pour excel_manager.py)

## Dépendances
- Story 1: Exceptions COM pour la gestion des erreurs Excel (doit être complétée avant cette story)
- Dépendances externes : `pywin32`, `pythoncom`, `ctypes`, `subprocess`, `gc`

## Notes
- Le gestionnaire doit implémenter le pattern RAII via context manager (`__enter__` et `__exit__`)
- Ne jamais appeler `app.Quit()` directement pour éviter les erreurs RPC
- Utiliser `Dispatch()` pour réutiliser une instance via ROT et `DispatchEx()` pour créer un processus isolé

## Résultats
✅ **Statut final :** Terminé et amélioré
✅ **Date de complétion initiale :** 2026-02-03
✅ **Date d'amélioration :** 2026-02-04
✅ **Tous les critères d'acceptation validés**
✅ **Tests unitaires complets créés** (66 tests passés, 0 skipped)
✅ **Intégration réussie** dans le module principal
✅ **Couverture de code :** 94% pour excel_manager.py, 100% pour exceptions.py, 90% global

## Validation
- Tests unitaires : 66/66 passés ✅ (amélioration de +48 tests)
- Couverture de code : 94% pour excel_manager.py ✅ (+21% d'amélioration)
- Couverture de code : 100% pour exceptions.py ✅ (+35% d'amélioration)
- Couverture globale : 90.05% ✅ (objectif atteint)
- Intégration : Méthodes exportées et utilisables ✅
- Documentation : Docstrings complètes pour chaque méthode ✅
- Conformité architecture : Respecte les spécifications définies ✅

## Améliorations apportées (2026-02-04)
- ✅ Test du context manager implémenté (était skippé)
- ✅ Tests complets pour les cas d'erreur dans `stop()`
- ✅ Tests pour les erreurs COM avec et sans HRESULT
- ✅ Tests pour les cas d'erreur dans `list_running_instances()`
- ✅ Tests pour les fonctions utilitaires (enumerate, connect)
- ✅ Tests supplémentaires pour les exceptions
