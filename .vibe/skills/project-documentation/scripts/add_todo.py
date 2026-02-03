"""
Script pour ajouter des entrées dans TODO.md.

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
from pathlib import Path


def update_header(content: str) -> str:
    """Met à jour la date de dernière mise à jour dans l'en-tête."""
    from datetime import datetime

    today = datetime.now().strftime("%Y-%m-%d")
    lines = content.split("\n")
    for i, line in enumerate(lines):
        if "**Dernière mise à jour**:" in line:
            lines[i] = f"**Dernière mise à jour**: {today}"
            break
    return "\n".join(lines)


def add_entry(
    content: str,
    section: str,
    title: str,
    entries: dict[str, str],
) -> str:
    """Ajoute une entrée dans une section spécifique."""
    lines = content.split("\n")

    entry_lines = [f"### {title}"]
    for key, value in entries.items():
        if value:
            entry_lines.append(f"- **{key}** : {value}")
        else:
            entry_lines.append(f"- **{key}** : N/A")

    insert_pos = None
    for i, line in enumerate(lines):
        if line.strip() == f"## {section}":
            insert_pos = i + 2  # Skip header and blank line
            break

    if insert_pos is None:
        raise ValueError(f"Section '{section}' non trouvée")

    lines.insert(insert_pos, "\n".join(entry_lines))
    lines.insert(insert_pos, "")

    return "\n".join(lines)


def add_feature(
    content: str,
    section: str,
    title: str,
    description: str,
    priority: str,
    adr_required: str = "non",
    estimation: str = "",
    dependencies: str = "",
) -> str:
    """Ajoute une feature."""
    entries = {
        "Priorité": priority,
        "Description": description,
        "ADR requis": adr_required,
        "Estimation": estimation,
        "Dépendances": dependencies,
    }
    return add_entry(content, section, title, entries)


def add_bug(
    content: str,
    title: str,
    module: str,
    description: str,
    severity: str,
    steps: list[str],
) -> str:
    """Ajoute un bug."""
    entries = {
        "Module affecté": module,
        "Description": description,
        "Sévérité": severity,
    }

    lines = content.split("\n")
    entry_lines = [f"### {title}"]
    for key, value in entries.items():
        if value:
            entry_lines.append(f"- **{key}** : {value}")

    if steps:
        entry_lines.append("- **Étapes pour reproduire** :")
        for i, step in enumerate(steps, 1):
            entry_lines.append(f"  {i}. {step}")

    insert_pos = None
    for i, line in enumerate(lines):
        if line.strip() == "## Bugs à Corriger":
            insert_pos = i + 2
            break

    if insert_pos is None:
        raise ValueError("Section 'Bugs à Corriger' non trouvée")

    lines.insert(insert_pos, "\n".join(entry_lines))
    lines.insert(insert_pos, "")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Ajouter des entrées dans TODO.md")
    parser.add_argument(
        "--todo-file",
        default="TODO.md",
        help="Chemin vers TODO.md (defaut: TODO.md a la racine)",
    )
    parser.add_argument(
        "--update-header",
        action="store_true",
        help="Mettre a jour la date dans l'en-tete",
    )

    subparsers = parser.add_subparsers(dest="command", help="Commande a executer")

    parser_feature = subparsers.add_parser("feature", help="Ajouter une feature")
    parser_feature.add_argument(
        "section",
        choices=["Features Prioritaires", "Futures Features"],
        help="Section cible",
    )
    parser_feature.add_argument("title", help="Titre de la feature")
    parser_feature.add_argument("description", help="Description de la feature")
    parser_feature.add_argument(
        "priority", choices=["Haute", "Moyenne", "Basse"], help="Priorite"
    )
    parser_feature.add_argument(
        "--adr", choices=["oui", "non"], default="non", help="ADR requis (defaut: non)"
    )
    parser_feature.add_argument("--estimation", help="Estimation (optionnel)")
    parser_feature.add_argument("--dependencies", help="Dependances (optionnel)")

    parser_bug = subparsers.add_parser("bug", help="Ajouter un bug")
    parser_bug.add_argument("title", help="Titre du bug")
    parser_bug.add_argument("module", help="Module affecte")
    parser_bug.add_argument("description", help="Description du bug")
    parser_bug.add_argument(
        "severity", choices=["Haute", "Moyenne", "Basse"], help="Severite"
    )
    parser_bug.add_argument("steps", nargs="+", help="Etapes pour reproduire")

    parser_generic = subparsers.add_parser(
        "generic", help="Ajouter une entree generique"
    )
    parser_generic.add_argument("section", help="Section cible")
    parser_generic.add_argument("title", help="Titre")
    parser_generic.add_argument("--desc", help="Description", required=True)
    parser_generic.add_argument("--module", help="Module affecte")
    parser_generic.add_argument("--priority", help="Priorite")
    parser_generic.add_argument("--reason", help="Raison (pour refactoring)")
    parser_generic.add_argument(
        "--approach", help="Approche suggeree (pour refactoring)"
    )
    parser_generic.add_argument("--advantages", help="Avantages (pour ameliorations)")
    parser_generic.add_argument(
        "--resources", help="Ressources (pour idees de recherche)"
    )

    args = parser.parse_args()

    todo_file = Path(args.todo_file)
    if not todo_file.exists():
        print(f"Erreur : {todo_file} n'existe pas")
        return 1

    content = todo_file.read_text(encoding="utf-8")

    try:
        if args.update_header:
            content = update_header(content)
            print("✓ En-tête mis à jour")

        if args.command == "feature":
            content = add_feature(
                content,
                args.section,
                args.title,
                args.description,
                args.priority,
                args.adr,
                args.estimation or "",
                args.dependencies or "",
            )
            print(f"✓ Feature ajoutée dans '{args.section}' : {args.title}")

        elif args.command == "bug":
            content = add_bug(
                content,
                args.title,
                args.module,
                args.description,
                args.severity,
                args.steps,
            )
            print(f"✓ Bug ajouté : {args.title}")

        elif args.command == "generic":
            entries = {"Description": args.desc}
            if args.module:
                entries["Module affecté"] = args.module
            if args.priority:
                entries["Priorité"] = args.priority
            if args.reason:
                entries["Raison"] = args.reason
            if args.approach:
                entries["Approche suggérée"] = args.approach
            if args.advantages:
                entries["Avantages"] = args.advantages
            if args.resources:
                entries["Ressources"] = args.resources

            content = add_entry(content, args.section, args.title, entries)
            print(f"✓ Entrée ajoutée dans '{args.section}' : {args.title}")

        if not any([args.update_header, args.command]):
            print("Aucune commande spécifiée. Utiliser --help pour les options.")
            return 0

        todo_file.write_text(content, encoding="utf-8")
        print(f"\n✓ {todo_file} mis à jour avec succès")
        return 0

    except ValueError as e:
        print(f"Erreur : {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
