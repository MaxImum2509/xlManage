"""
Script pour créer un PRD initial.

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


def create_prd(
    output_path: Path,
    product_name: str,
    author: str,
    version: str = "1.0",
) -> Path:
    """Crée un nouveau fichier prd.md."""
    today = datetime.now().strftime("%Y-%m-%d")

    content = f"""# Product Requirements Document (PRD) - {product_name}

**Version** : {version}
**Date de création** : {today}
**Dernière mise à jour** : {today}
**Statut** : Brouillon
**Auteur** : {author}
**Approuvé par** :

## Vue d'ensemble

### Description
[Description détaillée du produit]

### Objectifs
- [Objectif 1]
- [Objectif 2]

### Portée (In Scope)
- [Fonctionnalité 1]
- [Fonctionnalité 2]

### Hors portée (Out of Scope)
- [Fonctionnalité non inclue 1]
- [Fonctionnalité non inclue 2]

## Personas utilisateurs

### Persona 1 : [Nom]
- **Rôle** : [Rôle]
- **Expérience technique** : [Niveau]
- **Objectifs** :
  - [Objectif 1]
  - [Objectif 2]
- **Douleurs** :
  - [Douleur 1]
  - [Douleur 2]
- **Fréquence d'utilisation** : [Quotidienne/Hebdomadaire/Mensuelle]

### Persona 2 : [Nom]
[...]

## Exigences fonctionnelles

### Epic 1 : [Nom de l'epic]

#### FR-1.1 : [Titre de la fonctionnalité]
- **Description** : [Description détaillée]
- **Priorité** : [Must have/Should have/Could have/Won't have]
- **User Story** : En tant que [persona], je veux [action], afin que [bénéfice]
- **Critères d'acceptation** :
  - [ ] [Critère 1]
  - [ ] [Critère 2]
- **Dépendances** : [Autres fonctionnalités ou éléments]

#### FR-1.2 : [Titre de la fonctionnalité]
[...]

### Epic 2 : [Nom de l'epic]
[...]

## Exigences non-fonctionnelles

### Performance
- **NFR-P1** : [Exigence de performance]
- **NFR-P2** : [Exigence de performance]

### Fiabilité
- **NFR-R1** : [Exigence de fiabilité]
- **NFR-R2** : [Exigence de fiabilité]

### Sécurité
- **NFR-S1** : [Exigence de sécurité]
- **NFR-S2** : [Exigence de sécurité]

### Compatibilité
- **NFR-C1** : [Exigence de compatibilité]
- **NFR-C2** : [Exigence de compatibilité]

### Maintenance
- **NFR-M1** : [Exigence de maintenance]
- **NFR-M2** : [Exigence de maintenance]

## Architecture

Vue d'ensemble de l'architecture. Voir [architecture.md](architecture.md) pour les détails.

### Composants
- **Composant 1** : [Description]
- **Composant 2** : [Description]

## Tests

### Types de tests
- **Tests unitaires** : [Description]
- **Tests d'intégration** : [Description]
- **Tests de bout en bout** : [Description]

### Coverage
- **Objectif global** : > X%
- **Code critique** : > X%

## Roadmap et livrables

### Version 0.1.0 (MVP)
**Date cible** : YYYY-MM-DD

**Fonctionnalités** :
- [ ] [Fonctionnalité 1]
- [ ] [Fonctionnalité 2]

**Livrables** :
- [ ] Code source
- [ ] Documentation
- [ ] Tests

### Version 0.2.0
**Date cible** : YYYY-MM-DD

**Fonctionnalités** :
- [ ] [Fonctionnalité 3]
- [ ] [Fonctionnalité 4]

## Critères de succès

### Critères produit
- [ ] Toutes les fonctionnalités MVP implémentées
- [ ] Tests avec coverage > X%

### Critères qualité
- [ ] Linter et type checking passent
- [ ] Pas de bugs connus bloquants

## Historique des versions

### Version {version} ({today})
- Création initiale du PRD
"""

    output_path.write_text(content, encoding="utf-8")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Créer un nouveau PRD")
    parser.add_argument(
        "product_name",
        help="Nom du produit",
    )
    parser.add_argument(
        "--author",
        required=True,
        help="Auteur du PRD",
    )
    parser.add_argument(
        "--version",
        default="1.0",
        help="Version du PRD (défaut: 1.0)",
    )
    parser.add_argument(
        "--output",
        help="Chemin du fichier de sortie (défaut: _dev/prd.md)",
    )

    args = parser.parse_args()

    output_path = Path(args.output) if args.output else Path("_dev/prd.md")

    if output_path.exists():
        print(f"Attention : {output_path} existe déjà et sera écrasé.")
        response = input("Continuer ? (y/N) : ")
        if response.lower() != "y":
            print("Abandon.")
            return 1

    create_prd(output_path, args.product_name, args.author, args.version)

    print(f"✓ PRD créé : {output_path}")
    print(f"  Produit : {args.product_name}")
    print(f"  Auteur : {args.author}")
    print(f"  Version : {args.version}")
    print(f"\nN'oubliez pas de remplir les sections avec des informations réelles !")

    return 0


if __name__ == "__main__":
    sys.exit(main())
