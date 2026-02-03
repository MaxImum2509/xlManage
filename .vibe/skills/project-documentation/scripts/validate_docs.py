"""
Script pour valider la documentation du projet.

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
from pathlib import Path


def validate_adr_filename(filename: str) -> tuple[bool, str]:
    """Valide le format du nom de fichier ADR."""
    pattern = r"^ADR-\d{3}-.+-\d{8}\.md$"
    if not re.match(pattern, filename):
        return (
            False,
            f"Nom de fichier ADR invalide : {filename}. Format attendu : ADR-NNN-titre-aaaammjj.md",
        )
    return True, ""


def validate_adr_content(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu d'un fichier ADR."""
    errors = []

    frontmatter_pattern = r"^---\s*\n.+?\n---\s*\n"
    if not re.match(frontmatter_pattern, content, re.DOTALL):
        errors.append("Frontmatter YAML manquant ou invalide")

    required_sections = [
        "Contexte et description du problème",
        "Les options alternatives",
        "La décision",
        "Status des conséquences",
    ]

    for section in required_sections:
        if f"# {section}" not in content:
            errors.append(f"Section requise manquante : #{section}")

    return len(errors) == 0, errors


def validate_product_brief_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de product-brief.md."""
    errors = []

    if "# Product Brief" not in content:
        errors.append("Titre principal manquant : # Product Brief")

    required_sections = [
        "Résumé exécutif",
        "Problème",
        "Solution proposée",
        "Public cible",
        "Cas d'usage",
    ]

    for section in required_sections:
        if f"## {section}" not in content:
            errors.append(f"Section requise manquante : ## {section}")

    return len(errors) == 0, errors


def validate_prd_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de prd.md."""
    errors = []

    if "# Product Requirements Document" not in content:
        errors.append("Titre principal manquant : # Product Requirements Document")

    required_sections = [
        "Vue d'ensemble",
        "Personas utilisateurs",
        "Exigences fonctionnelles",
        "Exigences non-fonctionnelles",
    ]

    for section in required_sections:
        if f"## {section}" not in content:
            errors.append(f"Section requise manquante : ## {section}")

    return len(errors) == 0, errors


def validate_architecture_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de architecture.md."""
    errors = []

    if "# Architecture" not in content:
        errors.append("Titre principal manquant : # Architecture")

    required_sections = [
        "Vue d'ensemble",
        "Architecture de haut niveau",
        "Composants",
        "Patterns de design",
    ]

    for section in required_sections:
        if f"## {section}" not in content:
            errors.append(f"Section requise manquante : ## {section}")

    return len(errors) == 0, errors


def validate_epics_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de epics.md."""
    errors = []

    if "# Epics" not in content:
        errors.append("Titre principal manquant : # Epics")

    if "## Epic" not in content:
        errors.append(
            "Aucun epic trouvé. Le fichier doit contenir au moins un epic (## Epic X : ...)"
        )

    return len(errors) == 0, errors


def validate_progress_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de PROGRESS.md."""
    errors = []

    if "# Avancement du Projet" not in content:
        errors.append("Titre principal manquant : # Avancement du Projet")

    required_sections = [
        "Features Terminées",
        "En Cours d'Implémentation",
        "Tests et Qualité",
    ]

    for section in required_sections:
        if f"## {section}" not in content:
            errors.append(f"Section requise manquante : ## {section}")

    return len(errors) == 0, errors


def validate_todo_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de TODO.md."""
    errors = []

    if "# TODO" not in content:
        errors.append("Titre principal manquant : # TODO")

    required_sections = [
        "Features Prioritaires",
        "Futures Features",
        "Bugs à Corriger",
    ]

    for section in required_sections:
        if f"## {section}" not in content:
            errors.append(f"Section requise manquante : ## {section}")

    return len(errors) == 0, errors


def validate_changelog_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de CHANGELOG.md."""
    errors = []

    if "# Changelog" not in content:
        errors.append("Titre principal manquant : # Changelog")

    if "## [" not in content:
        errors.append(
            "Aucune version trouvée. Format attendu : ## [X.Y.Z] - YYYY-MM-DD"
        )

    return len(errors) == 0, errors


def validate_readme_md(content: str) -> tuple[bool, list[str]]:
    """Valide le contenu de README.md."""
    errors = []

    if not content.strip().startswith("# xlManage") and not content.strip().startswith(
        "# xlManage"
    ):
        errors.append("Titre principal manquant : # xlManage")

    required_sections = [
        "## Qu'est-ce que xlManage",
        "## Installation",
    ]

    for section in required_sections:
        if section not in content:
            errors.append(f"Section requise manquante : {section}")

    return len(errors) == 0, errors


def validate_dev_dir(dev_dir: Path) -> tuple[bool, list[str]]:
    """Valide tous les fichiers de documentation dans _dev/."""
    all_errors = []
    all_valid = True

    print(f"\nValidation du répertoire : {dev_dir}")
    print("=" * 60)

    # Valider README.md à la racine
    root_readme = dev_dir.parent / "README.md"
    if root_readme.exists():
        content = root_readme.read_text(encoding="utf-8")
        valid, errors = validate_readme_md(content)
        if not valid:
            all_errors.extend([f"README.md : {e}" for e in errors])
            all_valid = False
        else:
            print("  ✓ README.md")
    else:
        all_errors.append("README.md manquant à la racine du projet")
        all_valid = False

    # Fichiers à la racine
    root_files = [
        "PROGRESS.md",
        "TODO.md",
        "CHANGELOG.md",
        "CONTRIBUTING.md",
    ]

    for filename in root_files:
        filepath = dev_dir.parent / filename
        if not filepath.exists():
            all_errors.append(f"Fichier requis manquant à la racine : {filename}")
            all_valid = False
            continue

        content = filepath.read_text(encoding="utf-8")
        if filename == "PROGRESS.md":
            valid, errors = validate_progress_md(content)
            if not valid:
                all_errors.extend([f"{filename} : {e}" for e in errors])
                all_valid = False
            else:
                print(f"  ✓ {filename}")

        elif filename == "TODO.md":
            valid, errors = validate_todo_md(content)
            if not valid:
                all_errors.extend([f"{filename} : {e}" for e in errors])
                all_valid = False
            else:
                print(f"  ✓ {filename}")

        elif filename == "CHANGELOG.md":
            valid, errors = validate_changelog_md(content)
            if not valid:
                all_errors.extend([f"{filename} : {e}" for e in errors])
                all_valid = False
            else:
                print(f"  ✓ {filename}")

        elif filename == "CONTRIBUTING.md":
            print(f"  ✓ {filename}")

    required_files = [
        "product-brief.md",
        "prd.md",
        "architecture.md",
        "epics.md",
    ]

    for filename in required_files:
        filepath = dev_dir / filename
        if not filepath.exists():
            all_errors.append(f"Fichier requis manquant dans _dev/ : {filename}")
            all_valid = False
            continue

        content = filepath.read_text(encoding="utf-8")
        if filename == "product-brief.md":
            valid, errors = validate_product_brief_md(content)
            if not valid:
                all_errors.extend([f"{filename} : {e}" for e in errors])
                all_valid = False
            else:
                print(f"  ✓ {filename}")

        elif filename == "prd.md":
            valid, errors = validate_prd_md(content)
            if not valid:
                all_errors.extend([f"{filename} : {e}" for e in errors])
                all_valid = False
            else:
                print(f"  ✓ {filename}")

        elif filename == "architecture.md":
            valid, errors = validate_architecture_md(content)
            if not valid:
                all_errors.extend([f"{filename} : {e}" for e in errors])
                all_valid = False
            else:
                print(f"  ✓ {filename}")

        elif filename == "epics.md":
            valid, errors = validate_epics_md(content)
            if not valid:
                all_errors.extend([f"{filename} : {e}" for e in errors])
                all_valid = False
            else:
                print(f"  ✓ {filename}")

    adr_dir = dev_dir / "adr"
    if adr_dir.exists():
        print(f"\n  Validation des ADR ({adr_dir})")
        adr_files = list(adr_dir.glob("*.md"))
        if adr_files:
            for adr_file in sorted(adr_files):
                valid_filename, filename_error = validate_adr_filename(adr_file.name)
                if not valid_filename:
                    all_errors.append(filename_error)
                    all_valid = False
                    continue

                content = adr_file.read_text(encoding="utf-8")
                valid_content, content_errors = validate_adr_content(content)
                if not valid_content:
                    for error in content_errors:
                        all_errors.append(f"{adr_file.name} : {error}")
                    all_valid = False
                else:
                    print(f"    ✓ {adr_file.name}")
        else:
            all_errors.append("Répertoire adr/ vide")
            all_valid = False
    else:
        all_errors.append("Répertoire adr/ manquant")
        all_valid = False

    return all_valid, all_errors


def main():
    parser = argparse.ArgumentParser(description="Valider la documentation du projet")
    parser.add_argument(
        "--dev-dir",
        default="_dev",
        help="Répertoire de développement (défaut: _dev)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Échouer dès la première erreur",
    )

    args = parser.parse_args()

    dev_dir = Path(args.dev_dir)
    valid, errors = validate_dev_dir(dev_dir)

    if valid:
        print("\n" + "=" * 60)
        print("✓ Toute la documentation est valide !")
        return 0
    else:
        print("\n" + "=" * 60)
        print("✗ Erreurs trouvées :")
        for error in errors:
            print(f"  • {error}")
        print("=" * 60)
        return 1


if __name__ == "__main__":
    sys.exit(main())
