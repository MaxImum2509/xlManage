"""
Script pour créer un Product Brief initial.

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


def create_product_brief(
    output_path: Path,
    product_name: str,
    author: str,
) -> Path:
    """Crée un nouveau fichier product-brief.md."""
    today = datetime.now().strftime("%Y-%m-%d")

    content = f"""# Product Brief - {product_name}

**Date de création** : {today}
**Version** : 1.0
**Auteur** : {author}

## Résumé exécutif

[Paragraphe décrivant rapidement le produit, son objectif et son public cible]

## Problème

### Description
[Description détaillée du problème]

### Problème actuel
- [Problème 1]
- [Problème 2]

### Impact
- [Impact business]
- [Impact technique]
- [Impact utilisateur]

## Solution proposée

### Concept
[Description du concept de solution]

### Avantages
- [Avantage 1]
- [Avantage 2]

### Différenciation
[Comment se différencie de solutions existantes]

## Public cible

### Utilisateurs principaux
- [Type d'utilisateur 1] : [Description]
- [Type d'utilisateur 2] : [Description]

### Personas
#### Persona 1 : [Nom]
- **Rôle** : [Rôle]
- **Objectifs** : [Liste d'objectifs]
- **Douleurs** : [Liste de problèmes]
- **Motivations** : [Liste de motivations]

#### Persona 2 : [Nom]
[...]

## Cas d'usage

### Cas d'usage 1 : [Titre]
- **Qui** : [Type d'utilisateur]
- **Contexte** : [Contexte]
- **Objectif** : [Objectif]
- **Scénario** :
  1. [Étape 1]
  2. [Étape 2]
  3. [Étape 3]

### Cas d'usage 2 : [Titre]
[...]

## Opportunité de marché

### Taille du marché
- [Estimation de la taille du marché]
- [Segmentation]

### Tendances
- [Tendance 1]
- [Tendance 2]

### Concurrents
| Concurrent | Forces | Faiblesses |
|-----------|--------|------------|
| [Nom] | [...] | [...] |

## Hypothèses clés

### Hypothèses de produit
- [Hypothèse 1] : Comment valider
- [Hypothèse 2] : Comment valider

### Hypothèses techniques
- [Hypothèse technique 1]
- [Hypothèse technique 2]

## Mesures de succès

### KPIs
- [KPI 1] : [Objectif]
- [KPI 2] : [Objectif]

### Indicateurs de succès
- [Indicateur 1]
- [Indicateur 2]

## Risques et atténuations

| Risque | Probabilité | Impact | Atténuation |
|--------|-------------|--------|-------------|
| [Risque 1] | [Haute/Moyenne/Basse] | [Impact] | [Atténuation] |
| [Risque 2] | [...] | [...] | [...] |

## Prochaines étapes

### Immédiat
- [ ] [Action 1]
- [ ] [Action 2]

### Court terme
- [ ] [Action 3]
- [ ] [Action 4]

### Long terme
- [ ] [Action 5]
"""

    output_path.write_text(content, encoding="utf-8")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Créer un nouveau Product Brief")
    parser.add_argument(
        "product_name",
        help="Nom du produit",
    )
    parser.add_argument(
        "--author",
        required=True,
        help="Auteur du Product Brief",
    )
    parser.add_argument(
        "--output",
        help="Chemin du fichier de sortie (défaut: _dev/product-brief.md)",
    )

    args = parser.parse_args()

    output_path = Path(args.output) if args.output else Path("_dev/product-brief.md")

    if output_path.exists():
        print(f"Attention : {output_path} existe déjà et sera écrasé.")
        response = input("Continuer ? (y/N) : ")
        if response.lower() != "y":
            print("Abandon.")
            return 1

    create_product_brief(output_path, args.product_name, args.author)

    print(f"✓ Product Brief créé : {output_path}")
    print(f"  Produit : {args.product_name}")
    print(f"  Auteur : {args.author}")
    print("\nN'oubliez pas de remplir les sections avec des informations réelles !")

    return 0


if __name__ == "__main__":
    sys.exit(main())
