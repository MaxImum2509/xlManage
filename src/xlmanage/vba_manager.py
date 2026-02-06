"""
VBA Manager for manipulating VBA projects in Excel workbooks.

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

import re
from pathlib import Path

import pywintypes
from win32com.client import CDispatch

from .exceptions import VBAImportError, VBAProjectAccessError, VBAWorkbookFormatError

# Extension to module type mapping
EXTENSION_TO_TYPE: dict[str, str] = {
    ".bas": "standard",
    ".cls": "class",
    ".frm": "userform",
}


def _get_vba_project(wb: CDispatch) -> CDispatch:
    """Accède au VBProject avec gestion d'erreur.

    Args:
        wb: Objet COM Workbook

    Returns:
        CDispatch: Objet COM VBProject

    Raises:
        VBAProjectAccessError: Si Trust Center bloque l'accès
        VBAWorkbookFormatError: Si le classeur est en .xlsx
    """
    # Vérifier d'abord le format du classeur
    workbook_name = wb.Name
    if workbook_name.endswith(".xlsx"):
        raise VBAWorkbookFormatError(workbook_name)

    try:
        vb_project = wb.VBProject
        return vb_project
    except pywintypes.com_error as e:
        # HRESULT 0x800A03EC = Trust Center bloque l'accès
        if e.hresult == -2146827284:  # 0x800A03EC en signé
            raise VBAProjectAccessError(workbook_name) from e
        # Autre erreur COM inattendue
        raise


def _find_component(vb_project: CDispatch, name: str) -> CDispatch | None:
    """Recherche un composant VBA par nom.

    Args:
        vb_project: Objet COM VBProject
        name: Nom du module à chercher

    Returns:
        CDispatch | None: Composant VBA trouvé, ou None si absent
    """
    try:
        # Itérer sur VBComponents
        for component in vb_project.VBComponents:
            if component.Name == name:
                return component
        return None
    except pywintypes.com_error:
        # En cas d'erreur COM, retourner None
        return None


def _detect_module_type(path: Path) -> str:
    """Détecte le type de module depuis l'extension.

    Args:
        path: Chemin du fichier module VBA

    Returns:
        str: Type du module ("standard", "class", "userform")

    Raises:
        VBAImportError: Si l'extension n'est pas reconnue
    """
    extension = path.suffix.lower()

    if extension not in EXTENSION_TO_TYPE:
        raise VBAImportError(
            str(path),
            f"Extension '{extension}' non reconnue. "
            f"Extensions valides : {', '.join(EXTENSION_TO_TYPE.keys())}",
        )

    return EXTENSION_TO_TYPE[extension]


def _parse_class_module(file_path: Path) -> tuple[str, bool, str]:
    """Parse un fichier .cls pour extraire les métadonnées.

    Les fichiers .cls commencent par des lignes "Attribute VB_Name" qu'il
    faut parser séparément avant d'importer le code.

    Args:
        file_path: Chemin du fichier .cls

    Returns:
        tuple[str, bool, str]: (module_name, predeclared_id, code_content)
            - module_name: Nom du module extrait de VB_Name
            - predeclared_id: True si VB_PredeclaredId = True
            - code_content: Code source sans les attributs d'en-tête

    Raises:
        VBAImportError: Si le fichier est invalide ou mal encodé
    """
    try:
        # Lire le fichier avec l'encodage VBA (OBLIGATOIRE)
        content = file_path.read_text(encoding="windows-1252")
    except UnicodeDecodeError as e:
        raise VBAImportError(
            str(file_path),
            f"Encodage invalide. Les fichiers VBA doivent être en windows-1252 : {e}",
        ) from e

    # Extraire VB_Name
    name_match = re.search(r'Attribute VB_Name = "([^"]+)"', content)
    if not name_match:
        raise VBAImportError(
            str(file_path), "Attribut VB_Name manquant dans le fichier .cls"
        )
    module_name = name_match.group(1)

    # Extraire VB_PredeclaredId (False par défaut)
    predeclared_match = re.search(r"Attribute VB_PredeclaredId = (True|False)", content)
    predeclared_id = (
        predeclared_match.group(1) == "True" if predeclared_match else False
    )

    # Extraire le code (tout après la dernière ligne Attribute)
    # On cherche "Option Explicit" ou la première ligne de code
    code_start = content.find("Option Explicit")
    if code_start == -1:
        # Pas de Option Explicit, chercher la première ligne non-Attribute
        lines = content.splitlines()
        for i, line in enumerate(lines):
            if not line.startswith("VERSION") and not line.startswith("Attribute"):
                code_start = sum(len(line) + 2 for line in lines[:i])  # +2 pour \r\n
                break

    if code_start == -1:
        code_content = ""
    else:
        code_content = content[code_start:].strip()

    return module_name, predeclared_id, code_content
