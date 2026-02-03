# Story 3: Intégration CLI pour la gestion du cycle de vie Excel

**Epic:** Epic 5 - Gestion du cycle de vie Excel
**Priorité:** Moyenne
**Statut:** À faire

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
- [ ] Implémenter la commande `start` dans `src/xlmanage/cli.py`
- [ ] Implémenter la commande `stop` dans `src/xlmanage/cli.py`
- [ ] Implémenter la commande `status` dans `src/xlmanage/cli.py`
- [ ] Tester les commandes CLI pour s'assurer qu'elles fonctionnent correctement
- [ ] Vérifier que les messages d'erreur sont clairs et informatifs

## Dépendances
- Story 1: Exceptions COM pour la gestion des erreurs Excel (doit être complétée avant cette story)
- Story 2: Implémentation du gestionnaire de cycle de vie Excel (doit être complétée avant cette story)

## Notes
- Les commandes CLI doivent être minces et déléguer la logique métier au gestionnaire de cycle de vie Excel
- Utiliser Rich pour le formatage des messages de sortie
- Gérer les erreurs de manière appropriée et afficher des messages d'erreur clairs