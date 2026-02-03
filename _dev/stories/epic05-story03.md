# Story 3: Intégration CLI pour la gestion du cycle de vie Excel

**Epic:** Epic 5 - Gestion du cycle de vie Excel
**Priorité:** Moyenne
**Statut:** Terminé ✅

**Date de complétion:** 2026-02-04
**Version:** 1.0

## Description
Intégrer les fonctionnalités du gestionnaire de cycle de vie Excel dans l'interface en ligne de commande (CLI). Cela inclut les commandes pour démarrer, arrêter, et lister les instances Excel.

## Critères d'acceptation
1. Implémenter les commandes CLI suivantes dans `src/xlmanage/cli.py` :
   - `start` : Démarre une nouvelle instance Excel ou se connecte à une instance existante
   - `stop` : Arrête une instance Excel spécifique ou toutes les instances
   - `status` : Affiche le statut des instances Excel en cours d'exécution

2. Les commandes doivent accepter les options suivantes :
   - `start` : `--visible` (pour démarrer une instance visible), `--new` (pour forcer une nouvelle instance)
   - `stop` : `--all` (pour arrêter toutes les instances), `--force` (pour forcer l'arrêt), `--no-save` (pour ne pas sauvegarder les classeurs)
   - `status` : Aucune option requise

3. Les commandes doivent afficher des messages clairs et informatifs en utilisant Rich pour le formatage.

4. Les commandes doivent gérer les erreurs de manière appropriée et afficher des messages d'erreur clairs.

## Tâches
- [x] Implémenter la commande `start` dans `src/xlmanage/cli.py`
- [x] Implémenter la commande `stop` dans `src/xlmanage/cli.py`
- [x] Implémenter la commande `status` dans `src/xlmanage/cli.py`
- [x] Tester les commandes CLI pour s'assurer qu'elles fonctionnent correctement
- [x] Vérifier que les messages d'erreur sont clairs et informatifs
- [x] Créer des tests unitaires complets avec mocks
- [x] Vérifier la couverture de code (98% pour cli.py)

## Dépendances
- Story 1: Exceptions COM pour la gestion des erreurs Excel (doit être complétée avant cette story)
- Story 2: Implémentation du gestionnaire de cycle de vie Excel (doit être complétée avant cette story)

## Notes
- Les commandes CLI doivent être minces et déléguer la logique métier au gestionnaire de cycle de vie Excel
- Utiliser Rich pour le formatage des messages de sortie
- Gérer les erreurs de manière appropriée et afficher des messages d'erreur clairs

## Résultats
✅ **Statut final :** Terminé
✅ **Date de complétion :** 2026-02-04
✅ **Tous les critères d'acceptation validés**
✅ **Tests unitaires complets créés** (26 tests passés)
✅ **Intégration réussie** des commandes CLI
✅ **Couverture de code :** 98% pour cli.py

## Validation
- Commandes CLI : 3/3 implémentées ✅ (start, stop, status)
- Tests unitaires : 26/26 passés ✅
- Couverture de code : 98% pour cli.py ✅
- Intégration : Commandes fonctionnelles avec Rich ✅
- Gestion des erreurs : Messages clairs et informatifs ✅
- Conformité architecture : Délégation au ExcelManager ✅

## Fonctionnalités implémentées
### Commande `start`
- Options : `--visible` (afficher Excel), `--new` (forcer nouvelle instance)
- Affiche : PID, HWND, nombre de workbooks, mode, visibilité
- Gestion d'erreurs : ExcelConnectionError, ExcelManageError, Exception générique

### Commande `stop`
- Options : `--all` (toutes instances), `--force` (sans confirmation), `--no-save` (ne pas sauvegarder)
- Confirmations utilisateur pour sécurité
- Gestion des échecs partiels en mode --all
- Gestion d'erreurs complète

### Commande `status`
- Affiche tableau avec : PID, HWND, Visibilité (✓/✗), Nombre de workbooks
- Message informatif si aucune instance
- Gestion d'erreurs complète

### Commande `version`
- Affiche version formatée avec Rich (existait déjà, améliorée)
