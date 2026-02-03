"""
Exemple d'utilisation de xlManage avec un projet VBA.

Ce script démontre comment utiliser xlManage pour:
- Créer un nouveau classeur Excel
- Importer des modules VBA (bas, cls, frm)
- Exécuter des macros
- Gérer des fonctions VBA avec retour
- Exporter les modules modifiés

This file is part of xlManage.

xlManage is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xlManage is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with xlManage. If not, see <https://www.gnu.org/licenses/>.
"""

import sys
from pathlib import Path

# Ajoute src au path pour pouvoir importer xlmanage
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "src"))

from xlmanage import ExcelManager
from xlmanage.optimization import optimize_excel


def main() -> None:
    """Fonction principale de démonstration."""
    # Chemins
    project_dir = Path(__file__).parent
    modules_dir = project_dir / "modules"
    output_dir = project_dir / "output"
    output_dir.mkdir(exist_ok=True)

    print("=" * 60)
    print("Démonstration xlManage avec projet VBA")
    print("=" * 60)

    # Crée une instance Excel (masquée pour le test)
    print("\n1. Lancement d'Excel...")
    # excel = ExcelManager(visible=False)
    # Pour cet exemple, on simule car xlManage n'est pas encore implémenté
    print("   [Simulation] Excel lancé en mode masqué")

    # Crée un nouveau classeur
    print("\n2. Création d'un nouveau classeur...")
    print("   [Simulation] Classeur créé")

    # Importe les modules VBA
    print("\n3. Import des modules VBA...")
    modules_to_import = [
        modules_dir / "MainModule.bas",
        modules_dir / "ProductClass.cls",
        modules_dir / "DemoForm.frm",
    ]

    for module_file in modules_to_import:
        if module_file.exists():
            print(f"   [Simulation] Import de {module_file.name}")
        else:
            print(f"   [Attention] Fichier non trouvé: {module_file.name}")

    # Exécute une macro
    print("\n4. Exécution de macros...")
    print("   [Simulation] Exécution de MainModule.InitializeWorkbook")
    print("   [Simulation] Exécution de MainModule.ProcessData")
    print("   [Simulation] Exécution de MainModule.FormatTable")

    # Exécute une fonction avec retour
    print("\n5. Exécution de fonctions avec paramètres...")
    print("   [Simulation] CalculateTotal(10, 25.5) = 255.0")
    print("   [Simulation] GetProductName(0) = 'Produit A'")
    print("   [Simulation] ConcatenateStrings('Hello', 'World') = 'Hello World'")

    # Exporte les modules
    print("\n6. Export des modules...")
    print(f"   [Simulation] Export vers {output_dir}")

    # Sauvegarde le fichier
    output_file = output_dir / "xlmanage_demo.xlsm"
    print(f"\n7. Sauvegarde du classeur: {output_file}")
    print("   [Simulation] Fichier sauvegardé")

    # Ferme Excel
    print("\n8. Fermeture d'Excel...")
    print("   [Simulation] Excel fermé proprement")

    print("\n" + "=" * 60)
    print("Démonstration terminée!")
    print("=" * 60)
    print(f"\nFichiers créés:")
    print(f"  - {output_file}")
    print(f"\nModules disponibles:")
    for f in modules_to_import:
        if f.exists():
            print(f"  - {f.name}")


if __name__ == "__main__":
    main()
