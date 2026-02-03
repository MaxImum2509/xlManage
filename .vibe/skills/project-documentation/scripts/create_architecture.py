"""
Script pour créer un document d'architecture initial.

This file is part of xlManage.

xlManage is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xlManage is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with xlManage.  If not, see <https://www.gnu.org/licenses/>.
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path


def create_architecture(
    output_path: Path,
    product_name: str,
    author: str,
) -> Path:
    """Crée un nouveau fichier architecture.md."""
    today = datetime.now().strftime("%Y-%m-%d")

    content = f"""# Architecture - {product_name}

**Version** : 1.0
**Date de création** : {today}
**Dernière mise à jour** : {today}
**Auteur** : {author}

## Vue d'ensemble

[Description de l'architecture et des objectifs architecturaux]

### Objectifs architecturaux
- [Objectif 1]
- [Objectif 2]
- [Objectif 3]

## Architecture de haut niveau

### Couches

```
┌─────────────────────────────────────┐
│         Couche Application          │
│  (CLI, Scripts utilisateur)         │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│      Couche API ({product_name})    │
│  ┌─────────────────────────────┐    │
│  │      [Manager principal]    │    │
│  │  (Facade principale)        │    │
│  └─────────────────────────────┘    │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│      Couche d'abstraction           │
│  ┌─────────────────────────────┐    │
│  │  [Composants métier]        │    │
│  └─────────────────────────────┘    │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│      Couche d'accès                  │
│  (API externe / Bibliothèque)        │
└──────────────────────────────────────┘
```

### Flux de données
1. [Description du flux utilisateur vers la couche d'accès]
2. [Description du flux retour]

## Composants

### [Composant 1]
**Responsabilité** : [Description]

**Attributs** :
- `attr1` : [Description]
- `attr2` : [Description]

**Méthodes principales** :
```python
def method1(self) -> ReturnType:
    \"\"\"Description de la méthode.\"\"\"
    ...
```

**Responsabilités** :
- [Responsabilité 1]
- [Responsabilité 2]

### [Composant 2]
[...]

## Patterns de design

### [Pattern 1]
**Où** : [Composant]

**Pourquoi** : [Raison]

**Exemple** :
```python
# Exemple d'utilisation du pattern
...
```

### [Pattern 2]
[...]

## Gestion des erreurs

### Exceptions personnalisées
```python
class BaseError(Exception):
    \"\"\"Exception de base pour les erreurs.\"\"\"
    pass

class SpecificError(BaseError):
    \"\"\"Exception specifique.\"\"\"
    pass
```

### Stratégie de propagation
- [Description de la stratégie]

## Performance et optimisation

### Optimisations
1. [Optimisation 1]
2. [Optimisation 2]

### Benchmarks
| Opération | Temps | Notes |
|-----------|-------|-------|
| [Opération] | [Temps] | [Notes] |

## Sécurité

### Validation des entrées
- [Description]

### Protection
- [Description]

## Tests

### Tests unitaires
- [Description]

### Tests d'intégration
- [Description]

### Coverage
- **Objectif global** : > X%
- **Code critique** : > X%

## Dépendances

### Dépendances principales
- **[Dépendance 1]** : [Description]
  - Version : [Version]
  - Raison : [Raison]

### Dépendances de développement
- **[Dépendance 2]** : [Description]

## Décisions techniques

### ADR-XXX : [Titre de l'ADR]
- **Status** : [Status]
- **Date** : [Date]
- **Voir** : [ADR-XXX-titre-aaaammjj.md](adr/ADR-XXX-titre-aaaammjj.md)
- **Résumé** : [Résumé de la décision]

## Évolution et extensibilité

### Extensibilité prévue
- [Extensibilité 1]
- [Extensibilité 2]

### Futurs composants
- [Futur composant 1]
- [Futur composant 2]
"""

    output_path.write_text(content, encoding="utf-8")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="Créer un nouveau document d'architecture"
    )
    parser.add_argument(
        "product_name",
        help="Nom du produit",
    )
    parser.add_argument(
        "--author",
        required=True,
        help="Auteur du document",
    )
    parser.add_argument(
        "--output",
        help="Chemin du fichier de sortie (défaut: _dev/architecture.md)",
    )

    args = parser.parse_args()

    output_path = Path(args.output) if args.output else Path("_dev/architecture.md")

    if output_path.exists():
        print(f"Attention : {output_path} existe déjà et sera écrasé.")
        response = input("Continuer ? (y/N) : ")
        if response.lower() != "y":
            print("Abandon.")
            return 1

    create_architecture(output_path, args.product_name, args.author)

    print(f"✓ Architecture document créé : {output_path}")
    print(f"  Produit : {args.product_name}")
    print(f"  Auteur : {args.author}")
    print(f"\nN'oubliez pas de remplir les sections avec des informations réelles !")

    return 0


if __name__ == "__main__":
    sys.exit(main())
