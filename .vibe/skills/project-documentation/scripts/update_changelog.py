"""
Script pour mettre à jour CHANGELOG.md avec les changements d'une nouvelle version.

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
import re
import sys
from datetime import datetime
from pathlib import Path


def add_version(
    content: str,
    version: str,
    date: str,
    added: list[str] | None = None,
    changed: list[str] | None = None,
    deprecated: list[str] | None = None,
    removed: list[str] | None = None,
    fixed: list[str] | None = None,
    security: list[str] | None = None,
) -> str:
    """Ajoute une nouvelle version dans le changelog."""
    lines = content.split("\n")

    version_section = f"## [{version}] - {date}"
    version_content = [version_section, ""]

    if added:
        version_content.append("### Ajouté")
        for item in added:
            version_content.append(f"- {item}")
        version_content.append("")

    if changed:
        version_content.append("### Changé")
        for item in changed:
            version_content.append(f"- {item}")
        version_content.append("")

    if deprecated:
        version_content.append("### Déprécié")
        for item in deprecated:
            version_content.append(f"- {item}")
        version_content.append("")

    if removed:
        version_content.append("### Supprimé")
        for item in removed:
            version_content.append(f"- {item}")
        version_content.append("")

    if fixed:
        version_content.append("### Corrigé")
        for item in fixed:
            version_content.append(f"- {item}")
        version_content.append("")

    if security:
        version_content.append("### Sécurité")
        for item in security:
            version_content.append(f"- {item}")
        version_content.append("")

    insert_pos = None
    for i, line in enumerate(lines):
        if re.match(r"^## \[", line):
            insert_pos = i
            break

    if insert_pos is None:
        insert_pos = len(lines)

    for i, line in enumerate(reversed(version_content)):
        lines.insert(insert_pos, line)

    return "\n".join(lines)


def add_unreleased(
    content: str,
    section: str,
    items: list[str],
) -> str:
    """Ajoute des changements dans la section Unreleased."""
    lines = content.split("\n")

    section_map = {
        "ajouté": "### Ajouté",
        "changé": "### Changé",
        "déprécié": "### Déprécié",
        "supprimé": "### Supprimé",
        "corrigé": "### Corrigé",
        "sécurité": "### Sécurité",
    }

    section_header = section_map.get(section.lower())
    if not section_header:
        raise ValueError(
            f"Section invalide : {section}. Options : {', '.join(section_map.keys())}"
        )

    unreleased_pos = None
    for i, line in enumerate(lines):
        if line.strip() == "## [Unreleased]":
            unreleased_pos = i
            break

    if unreleased_pos is None:
        raise ValueError("Section '[Unreleased]' non trouvée")

    section_pos = None
    for i in range(unreleased_pos + 1, len(lines)):
        if lines[i].strip() == section_header:
            section_pos = i
            break

    if section_pos is None:
        insert_pos = unreleased_pos + 2
        lines.insert(insert_pos, "")
        lines.insert(insert_pos, section_header)
        lines.insert(insert_pos + 2, "")
        section_pos = insert_pos + 2
    else:
        insert_pos = section_pos + 1
        if lines[insert_pos].strip() != "":
            lines.insert(insert_pos, "")

    for i, item in enumerate(items):
        lines.insert(insert_pos + i, f"- {item}")

    return "\n".join(lines)


def release_version(content: str, version: str, date: str | None = None) -> str:
    """Convertit la section [Unreleased] en une version spécifique."""
    if not date:
        date = datetime.now().strftime("%Y-%m-%d")

    lines = content.split("\n")

    unreleased_pos = None
    for i, line in enumerate(lines):
        if line.strip() == "## [Unreleased]":
            unreleased_pos = i
            break

    if unreleased_pos is None:
        raise ValueError("Section '[Unreleased]' non trouvée")

    lines[unreleased_pos] = f"## [{version}] - {date}"

    next_version_pos = None
    for i in range(unreleased_pos + 1, len(lines)):
        if re.match(r"^## \[", lines[i]):
            next_version_pos = i
            break

    if next_version_pos is None:
        lines.append("")
        lines.append("## [Unreleased]")

    unreleased_section = []
    for i in range(unreleased_pos + 1, next_version_pos or len(lines)):
        unreleased_section.append(lines[i])

    if next_version_pos:
        for i, line in enumerate(reversed(unreleased_section)):
            lines.insert(next_version_pos, line)

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Mettre à jour CHANGELOG.md")
    parser.add_argument(
        "--changelog-file",
        default="CHANGELOG.md",
        help="Chemin vers CHANGELOG.md (défaut: CHANGELOG.md à la racine)",
    )
    parser.add_argument(
        "--add-version",
        metavar="VERSION",
        help="Ajouter une nouvelle version",
    )
    parser.add_argument(
        "--date",
        help="Date de la version (défaut: aujourd'hui, format YYYY-MM-DD)",
    )
    parser.add_argument(
        "--added",
        nargs="*",
        help="Liste des ajouts",
    )
    parser.add_argument(
        "--changed",
        nargs="*",
        help="Liste des modifications",
    )
    parser.add_argument(
        "--deprecated",
        nargs="*",
        help="Liste des dépréciations",
    )
    parser.add_argument(
        "--removed",
        nargs="*",
        help="Liste des suppressions",
    )
    parser.add_argument(
        "--fixed",
        nargs="*",
        help="Liste des corrections",
    )
    parser.add_argument(
        "--security",
        nargs="*",
        help="Liste des corrections de sécurité",
    )
    parser.add_argument(
        "--add-unreleased",
        metavar="SECTION",
        help="Ajouter des changements dans la section Unreleased (ajouté/changé/déprécié/supprimé/corrigé/sécurité)",
    )
    parser.add_argument(
        "--items",
        nargs="*",
        help="Éléments à ajouter (à utiliser avec --add-unreleased)",
    )
    parser.add_argument(
        "--release",
        metavar="VERSION",
        help="Convertir Unreleased en version spécifique",
    )

    args = parser.parse_args()

    changelog_file = Path(args.changelog_file)
    if not changelog_file.exists():
        print(f"Erreur : {changelog_file} n'existe pas")
        return 1

    content = changelog_file.read_text(encoding="utf-8")

    try:
        if args.add_version:
            date = args.date or datetime.now().strftime("%Y-%m-%d")
            content = add_version(
                content,
                args.add_version,
                date,
                args.added,
                args.changed,
                args.deprecated,
                args.removed,
                args.fixed,
                args.security,
            )
            print(f"✓ Version {args.add_version} ajoutée pour le {date}")

        elif args.add_unreleased:
            if not args.items:
                print("Erreur : --items est requis avec --add-unreleased")
                return 1

            content = add_unreleased(content, args.add_unreleased, args.items)
            print(
                f"✓ {len(args.items)} élément(s) ajouté(s) dans la section '{args.add_unreleased}'"
            )

        elif args.release:
            content = release_version(content, args.release, args.date)
            print(f"✓ Version {args.release} publiée")

        else:
            print("Aucune action spécifiée. Utiliser --help pour les options.")
            return 0

        changelog_file.write_text(content, encoding="utf-8")
        print(f"\n✓ {changelog_file} mis à jour avec succès")
        return 0

    except ValueError as e:
        print(f"Erreur : {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
