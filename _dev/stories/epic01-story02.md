# Story 2: Configuration de l'environnement de test avec pytest

**Epic:** Epic 01 - Mise en place de l'environnement de développement
**Priorité:** Haute
**Statut:** Terminé

## Description
Configurer l'environnement de test pour le projet xlManage en utilisant pytest et les plugins associés. Cela inclut la configuration des tests unitaires, des tests d'intégration, et des outils de couverture de code.

## Résultats
✅ **Statut final :** Terminé
✅ **Date de complétion :** 2026-02-03
✅ **Couverture de code atteinte :** 91% (seuil de 90% dépassé)
✅ **Tous les critères d'acceptation validés**

## Critères d'acceptation
1. Installer les dépendances de test nécessaires :
   - `pytest` : Framework de test principal
   - `pytest-cov` : Couverture de code (seuil minimum : 90%)
   - `pytest-mock` : Injection de mocks via fixture `mocker`
   - `pytest-timeout` : Timeout 60s par test (prévenir les blocages COM)
   - `unittest.mock.Mock` : Mock des objets COM
   - `typer.testing.CliRunner` : Test des commandes CLI en isolation

2. Créer le fichier de configuration `pytest.ini` à la racine du projet avec les options suivantes :
   - `testpaths = tests`
   - `python_files = test_*.py`
   - `python_functions = test_*`
   - `addopts = --cov=src/ --cov-report=html --cov-report=term --cov-fail-under=90`

3. Créer le fichier `tests/conftest.py` pour les fixtures et hooks globaux pytest :
   - Fixtures partagées : Excel app, workbooks, workbooks temporaires, etc.
   - Hooks pytest : Nettoyage automatique des ressources COM
   - Configuration timeout : Timeout par défaut pour les tests (recommandé : 60s)
   - Configuration markers : Markers pour catégoriser les tests (ex: @pytest.mark.com pour tests COM)

4. Vérifier que les tests peuvent être exécutés avec la commande `pytest --cov=src/ --cov-report=html --cov-report=term --cov-fail-under=90`

## Tâches
- [x] Installer les dépendances de test avec Poetry
- [x] Créer le fichier `pytest.ini` avec la configuration requise
- [x] Créer le fichier `tests/conftest.py` avec les fixtures et hooks globaux
- [x] Tester la configuration avec un test simple
- [x] Vérifier que la couverture de code est mesurée correctement

## Dépendances
- Story 1: Création de la structure de répertoires du projet (doit être complétée avant cette story)
- Dépendances externes : Poetry pour la gestion des dépendances

## Notes
- Utiliser Poetry pour ajouter les dépendances de test : `poetry add --group dev <package>`
- Suivre les bonnes pratiques de test pour les projets Python
- S'assurer que les tests sont isolés et ne dépendent pas de l'état global
- Utiliser des mocks pour simuler les objets COM et éviter les tests lents ou instables