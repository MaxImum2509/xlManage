"""
Script pour créer un nouvel Architecture Decision Record (ADR).

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
from datetime import datetime
from pathlib import Path


def get_next_adr_number(adr_dir: Path) -> int:
    """Trouve le prochain numéro ADR disponible."""
    if not adr_dir.exists():
        return 1

    adrs = list(adr_dir.glob("ADR-*.md"))
    if not adrs:
        return 1

    max_num = 0
    for adr in adrs:
        try:
            num = int(adr.name.split("-")[1])
            max_num = max(max_num, num)
        except (ValueError, IndexError):
            continue

    return max_num + 1


def create_adr(
    adr_dir: Path,
    title: str,
    status: str = "proposé",
    decision_makers: str = "",
    technical_stories: str = "",
) -> Path:
    """Crée un nouveau fichier ADR."""
    adr_dir.mkdir(parents=True, exist_ok=True)

    num = get_next_adr_number(adr_dir)
    date = datetime.now().strftime("%Y%m%d")
    kebab_title = title.lower().replace(" ", "-").replace("_", "-")

    filename = f"ADR-{num:03d}-{kebab_title}-{date}.md"
    filepath = adr_dir / filename

    content = f"""---
title: ADR-{num:03d}: {title}
status: {status}
date: {date}
decision-makers: {decision_makers}
technical-stories: {technical_stories}
---

# Contexte et description du problème

[Description du problème à résoudre, du contexte et de la motivation pour cette décision.]

# Les options alternatives

## Option 1: [Titre de l'option]
[Description de l'option]
- **Avantages** : Liste des avantages
- **Désavantages** : Liste des désavantages

## Option 2: [Titre de l'option]
[Description de l'option]
- **Avantages** : Liste des avantages
- **Désavantages** : Liste des désavantages

# La décision

[Description de la décision choisie et sa justification.]

# Status des conséquences

[Description des conséquences positives et négatives de cette décision, et ce qui doit être fait ensuite.]

## Avantages
- [Liste des avantages]

## Désavantages
- [Liste des désavantages]

## ADRs concernés
- [ADR-XXX-titre-aaaammjj.md](ADR-XXX-titre-aaaammjj.md)
"""

    filepath.write_text(content, encoding="utf-8")
    return filepath


def main():
    parser = argparse.ArgumentParser(
        description="Créer un nouveau Architecture Decision Record (ADR)"
    )
    parser.add_argument("title", help="Titre de la décision")
    parser.add_argument(
        "--status",
        choices=["proposé", "accepté", "rejeté", "obsolète", "remplacé"],
        default="proposé",
        help="Status de la décision (défaut: proposé)",
    )
    parser.add_argument(
        "--decision-makers", default="", help="Personnes ayant pris la décision"
    )
    parser.add_argument(
        "--technical-stories", default="", help="Tickets/liens concernés"
    )
    parser.add_argument(
        "--adr-dir",
        help="Répertoire des ADR (défaut: _dev/adr)",
    )

    args = parser.parse_args()

    adr_dir = Path(args.adr_dir) if args.adr_dir else Path("_dev/adr")
    filepath = create_adr(
        adr_dir=adr_dir,
        title=args.title,
        status=args.status,
        decision_makers=args.decision_makers,
        technical_stories=args.technical_stories,
    )

    print(f"✓ ADR créé : {filepath}")
    print(f"  Numéro : {filepath.name.split('-')[1]}")
    print(f"  Status : {args.status}")


if __name__ == "__main__":
    main()
