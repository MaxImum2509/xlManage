"""
Script pour créer un fichier d'epics initial.

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


def create_epics(
    output_path: Path,
    product_name: str,
    first_version: str = "0.1.0",
) -> Path:
    """Crée un nouveau fichier epics.md."""
    today = datetime.now().strftime("%Y-%m-%d")

    content = f"""# Epics - {product_name}

**Version** : 1.0
**Date de création** : {today}
**Dernière mise à jour** : {today}

## Vue d'ensemble

Cette document décrit les epics planifiés pour {product_name}. Les epics sont regroupés
par version et incluent les user stories correspondantes.

### Liste des epics
1. [Epic 1] - [Version]
2. [Epic 2] - [Version]
3. [Epic 3] - [Version]

## Epic 1 : [Nom de l'epic]

**Version cible** : {first_version}
**Priorité** : [Haute/Moyenne/Basse]
**Statut** : [Non demarre/En cours/Termine]
**Proprietaire** : [Nom]
**ADR associé** : [ADR-XXX-titre-aaaammjj.md](adr/ADR-XXX-titre-aaaammjj.md)

### Description
[Description detaillee de l'epic]

### Objectifs
- [Objectif 1]
- [Objectif 2]

### Stories

#### Story 1.1 : [Titre de la story]
- **User Story** : En tant que [persona], je veux [action], afin que [benefice]
- **Priorité** : [Must have/Should have/Could have]
- **Complexite** : [Petite/Moyenne/Grande]
- **Estimation** : X jours/heures
- **Dependances** : [Story ou fonctionnalité]
- **Criteres d'acceptation** :
  - [ ] [Critere 1]
  - [ ] [Critere 2]
- **Tests** :
  - [ ] Test unitaire 1
  - [ ] Test d'intégration 1
- **Statut** : [Non demarre/En cours/Termine]
- **Assigné à** : [Nom]

#### Story 1.2 : [Titre de la story]
[...]

### Risques
- [Risque 1] : [Mitigation]
- [Risque 2] : [Mitigation]

### Metriques de succes
- [Metrique 1] : [Objectif]
- [Metrique 2] : [Objectif]

## Epic 2 : [Nom de l'epic]

**Version cible** : [Version]
**Priorité** : [Haute/Moyenne/Basse]
**Statut** : [Non demarre/En cours/Termine]
**Proprietaire** : [Nom]
**ADR associé** : [ADR-XXX-titre-aaaammjj.md](adr/ADR-XXX-titre-aaaammjj.md)

### Description
[...]

### Objectifs
[...]

### Stories
[...]

### Risques
[...]

### Metriques de succes
[...]

## Epic 3 : [Nom de l'epic]
[...]

## Historique des versions

### Version 1.0 ({today})
- Création initiale du document
- Définition des epics pour la version {first_version}
"""

    output_path.write_text(content, encoding="utf-8")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Créer un nouveau fichier d'epics")
    parser.add_argument(
        "product_name",
        help="Nom du produit",
    )
    parser.add_argument(
        "--first-version",
        default="0.1.0",
        help="Première version (défaut: 0.1.0)",
    )
    parser.add_argument(
        "--output",
        help="Chemin du fichier de sortie (défaut: _dev/epics.md)",
    )

    args = parser.parse_args()

    output_path = Path(args.output) if args.output else Path("_dev/epics.md")

    if output_path.exists():
        print(f"Attention : {output_path} existe déjà et sera écrasé.")
        response = input("Continuer ? (y/N) : ")
        if response.lower() != "y":
            print("Abandon.")
            return 1

    create_epics(output_path, args.product_name, args.first_version)

    print(f"✓ Epics créé : {output_path}")
    print(f"  Produit : {args.product_name}")
    print(f"  Première version : {args.first_version}")
    print("\nN'oubliez pas de remplir les sections avec des informations réelles !")

    return 0


if __name__ == "__main__":
    sys.exit(main())
