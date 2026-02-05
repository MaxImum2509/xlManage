# Epic 8 - Story 6: Intégration CLI complète

**Statut** : ✅ À implémenter

**En tant que** utilisateur
**Je veux** utiliser toutes les commandes table via CLI
**Afin de** manipuler les tables depuis la ligne de commande

## Critères d'acceptation

1. ✅ Commandes CLI implémentées : `table create/delete/list`
2. ✅ Options CLI complètes
3. ✅ Affichage Rich (panels, tables)
4. ✅ Gestion d'erreur complète
5. ✅ Tests CLI complets

## Tâches techniques

### Tâche 6.1 : Ajouter les commandes CLI

**Fichier** : `src/xlmanage/cli.py`

Après le groupe `worksheet`, ajouter un groupe `table` avec 3 commandes :

1. **table create** : Créer une table
   - Argument 1 : `name` (nom de la table)
   - Argument 2 : `range_ref` (plage Excel, ex: "A1:D100")
   - Option : `--worksheet` / `-ws` (feuille cible, optional)
   - Option : `--workbook` / `-w` (classeur cible, optional)

2. **table delete** : Supprimer une table
   - Argument : `name` (nom de la table)
   - Option : `--worksheet` / `-ws` (feuille contenant la table, optional)
   - Option : `--workbook` / `-w` (classeur cible, optional)
   - Option : `--force` / `-f` (supprimer sans confirmation)

3. **table list** : Lister les tables
   - Option : `--worksheet` / `-ws` (feuille cible, optional; si non fourni, lister tout le classeur)
   - Option : `--workbook` / `-w` (classeur cible, optional)

### Tâche 6.2 : Ajouter les imports dans cli.py

En haut de `cli.py`, ajouter :

```python
from .exceptions import (
    # ... existing ...
    TableNotFoundError,
    TableAlreadyExistsError,
    TableRangeError,
    TableNameError,
)
```

### Tâche 6.3 : Gestion d'erreur et affichage Rich

Chaque commande doit :
1. Afficher un Panel de succès (vert) avec les détails
2. Afficher un Panel d'erreur (rouge) en cas de problème
3. Pour `list` : afficher une Table avec Nom, Feuille, Plage, Lignes

### Tâche 6.4 : Confirmation pour delete

Avant de supprimer, demander la confirmation à l'utilisateur (sauf avec `--force`)

## Dépendances

- Story 1-5 (Toutes les stories précédentes)
- Groupes workbook et worksheet (Epic 6 et 7)

## Définition of Done

- [ ] 3 commandes CLI implémentées
- [ ] Affichage Rich avec panels et tables
- [ ] Gestion d'erreur complète
- [ ] Confirmation pour delete
- [ ] Tous les tests CLI passent (minimum 3 tests)
- [ ] Messages utilisateur en français
- [ ] Options complètes (--worksheet et --workbook sur chaque commande)

## Livrables

### Commandes disponibles

```bash
# Créer une table
xlmanage table create "tbl_Sales" "A1:D100"
xlmanage table create "tbl_Sales" "A1:D100" --worksheet "Data"

# Supprimer une table
xlmanage table delete "tbl_Sales"
xlmanage table delete "tbl_Sales" --force

# Lister les tables
xlmanage table list
xlmanage table list --worksheet "Data"
xlmanage table list --workbook /path/to/file.xlsx
```

### Architecture

```
cli.py
├── table_app = typer.Typer()
│   ├── @table_app.command("create")
│   │   └── table_create()
│   ├── @table_app.command("delete")
│   │   └── table_delete()
│   └── @table_app.command("list")
│       └── table_list()
└── app.add_typer(table_app, name="table")
```

Chaque commande :
1. Crée un ExcelManager context manager
2. Crée un TableManager
3. Appelle la méthode correspondante
4. Affiche le résultat ou gère l'erreur
