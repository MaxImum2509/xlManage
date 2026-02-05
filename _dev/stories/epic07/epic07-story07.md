# Epic 7 - Story 7: Intégration CLI complète

**Statut** : ✅ Terminé

**En tant que** utilisateur
**Je veux** utiliser toutes les commandes worksheet via CLI
**Afin de** manipuler les feuilles depuis la ligne de commande

## Critères d'acceptation

1. ✅ Commandes CLI implémentées : `worksheet create/delete/list/copy`
2. ✅ Options CLI complètes
3. ✅ Affichage Rich (panels, tables)
4. ✅ Gestion d'erreur complète
5. ✅ Tests CLI complets

## Tâches techniques

### Tâche 7.1 : Ajouter les commandes CLI

**Fichier** : `src/xlmanage/cli.py`

Après le groupe `workbook`, ajouter un groupe `worksheet` avec 4 commandes :

1. **worksheet create** : Créer une feuille
   - Argument : `name` (nom de la nouvelle feuille)
   - Option : `--workbook` / `-w` (classeur cible, optional)

2. **worksheet delete** : Supprimer une feuille
   - Argument : `name` (nom de la feuille à supprimer)
   - Option : `--workbook` / `-w` (classeur cible, optional)
   - Option : `--force` / `-f` (forcer sans confirmation)

3. **worksheet list** : Lister les feuilles
   - Option : `--workbook` / `-w` (classeur cible, optional)

4. **worksheet copy** : Copier une feuille
   - Argument 1 : `source` (feuille source)
   - Argument 2 : `destination` (nom de la copie)
   - Option : `--workbook` / `-w` (classeur cible, optional)

### Tâche 7.2 : Ajouter les imports dans cli.py

En haut de `cli.py`, ajouter :

```python
from .exceptions import (
    # ... existing ...
    WorksheetNotFoundError,
    WorksheetAlreadyExistsError,
    WorksheetDeleteError,
    WorksheetNameError,
)
```

### Tâche 7.3 : Gestion d'erreur et affichage Rich

Chaque commande doit :
1. Afficher un Panel de succès (vert) avec les détails
2. Afficher un Panel d'erreur (rouge) en cas de problème
3. Pour `list` : afficher une Table avec Position, Nom, Visible, Lignes, Colonnes

### Tâche 7.4 : Confirmation pour delete

Avant de supprimer, demander la confirmation à l'utilisateur (sauf avec `--force`)

## Dépendances

- Story 1-6 (Toutes les stories précédentes)
- Groupe workbook (de Epic 6)

## Définition of Done

- [x] 4 commandes CLI implémentées
- [x] Affichage Rich avec panels et tables
- [x] Gestion d'erreur complète
- [x] Confirmation pour delete
- [x] Tous les tests CLI passent (minimum 4 tests) - 17 tests passent
- [x] Messages utilisateur en français
- [x] Options complètes (--workbook sur chaque commande)

## Livrables

### Commandes disponibles

```bash
# Créer une feuille
xlmanage worksheet create "NomFeuille"
xlmanage worksheet create "NomFeuille" --workbook /path/to/file.xlsx

# Supprimer une feuille
xlmanage worksheet delete "NomFeuille"
xlmanage worksheet delete "NomFeuille" --force

# Lister les feuilles
xlmanage worksheet list
xlmanage worksheet list --workbook /path/to/file.xlsx

# Copier une feuille
xlmanage worksheet copy "Source" "Destination"
xlmanage worksheet copy "Source" "Destination" --workbook /path/to/file.xlsx
```

## Architecture

```
cli.py
├── worksheet_app = typer.Typer()
│   ├── @worksheet_app.command("create")
│   │   └── worksheet_create()
│   ├── @worksheet_app.command("delete")
│   │   └── worksheet_delete()
│   ├── @worksheet_app.command("list")
│   │   └── worksheet_list()
│   └── @worksheet_app.command("copy")
│       └── worksheet_copy()
└── app.add_typer(worksheet_app, name="worksheet")
```

Chaque commande :
1. Crée un ExcelManager context manager
2. Crée un WorksheetManager
3. Appelle la méthode correspondante
4. Affiche le résultat ou gère l'erreur
