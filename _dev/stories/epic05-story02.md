# Story 2: Implémentation du gestionnaire de cycle de vie Excel

**Epic:** Epic 5 - Gestion du cycle de vie Excel
**Priorité:** Haute
**Statut:** À faire

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
- [ ] Créer le fichier `src/xlmanage/excel_manager.py`
- [ ] Implémenter la dataclass `InstanceInfo`
- [ ] Implémenter la classe `ExcelManager` avec toutes les méthodes requises
- [ ] Implémenter les fonctions utilitaires pour l'énumération des instances
- [ ] Tester les fonctionnalités de base (démarrage, connexion, énumération)

## Dépendances
- Story 1: Exceptions COM pour la gestion des erreurs Excel (doit être complétée avant cette story)
- Dépendances externes : `pywin32`, `pythoncom`, `ctypes`, `subprocess`, `gc`

## Notes
- Le gestionnaire doit implémenter le pattern RAII via context manager (`__enter__` et `__exit__`)
- Ne jamais appeler `app.Quit()` directement pour éviter les erreurs RPC
- Utiliser `Dispatch()` pour réutiliser une instance via ROT et `DispatchEx()` pour créer un processus isolé