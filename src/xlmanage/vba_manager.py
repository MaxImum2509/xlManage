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
from dataclasses import dataclass
from pathlib import Path

import pywintypes
from win32com.client import CDispatch

from .excel_manager import ExcelManager
from .exceptions import (
    VBAImportError,
    VBAModuleAlreadyExistsError,
    VBAProjectAccessError,
    VBAWorkbookFormatError,
)


@dataclass
class VBAModuleInfo:
    """Informations sur un module VBA.

    Attributes:
        name: Nom du module (ex: "Module1", "MyClass")
        module_type: Type du module ("standard", "class", "userform", "document")
        lines_count: Nombre de lignes de code dans le module
        has_predeclared_id: True si PredeclaredId activé (classes uniquement)
    """

    name: str
    module_type: str
    lines_count: int
    has_predeclared_id: bool = False


# Types de composants VBA (constantes Excel)
VBEXT_CT_STD_MODULE: int = 1  # Module standard (.bas)
VBEXT_CT_CLASS_MODULE: int = 2  # Module de classe (.cls)
VBEXT_CT_MS_FORM: int = 3  # UserForm (.frm + .frx)
VBEXT_CT_DOCUMENT: int = 100  # Module de document (ThisWorkbook, Sheet1)

# Mapping type code -> nom lisible
VBA_TYPE_NAMES: dict[int, str] = {
    1: "standard",
    2: "class",
    3: "userform",
    100: "document",
}

# Extension to module type mapping
EXTENSION_TO_TYPE: dict[str, str] = {
    ".bas": "standard",
    ".cls": "class",
    ".frm": "userform",
}

# Encodage obligatoire pour les fichiers VBA
VBA_ENCODING: str = "windows-1252"


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


class VBAManager:
    """Gestionnaire des modules VBA.

    Permet d'importer, exporter, lister et supprimer des modules VBA
    dans les classeurs Excel. Nécessite que le Trust Center autorise
    l'accès programmatique aux projets VBA.

    Important:
        - Le classeur doit être au format .xlsm pour supporter les macros
        - L'option "Trust access to the VBA project object model" doit
          être activée dans Excel Trust Center

    Example:
        >>> with ExcelManager() as excel_mgr:
        ...     excel_mgr.start()
        ...     vba_mgr = VBAManager(excel_mgr)
        ...     modules = vba_mgr.list_modules()
        ...     for module in modules:
        ...         print(f"{module.name}: {module.module_type}")
    """

    def __init__(self, excel_manager: ExcelManager):
        """Initialize VBA manager.

        Args:
            excel_manager: Instance d'ExcelManager déjà démarrée.
                Utilisé pour accéder à l'objet COM Application.

        Example:
            >>> excel_mgr = ExcelManager()
            >>> excel_mgr.start()
            >>> vba_mgr = VBAManager(excel_mgr)
        """
        self._mgr = excel_manager

    @property
    def app(self) -> CDispatch:
        """Objet COM Excel.Application.

        Returns:
            CDispatch: Application Excel active

        Raises:
            RuntimeError: Si Excel n'est pas démarré
        """
        return self._mgr.app

    def import_module(
        self,
        module_file: Path,
        module_type: str | None = None,
        workbook: Path | None = None,
        overwrite: bool = False,
    ) -> VBAModuleInfo:
        """Importe un module VBA depuis un fichier.

        Supporte les modules standard (.bas), de classe (.cls) et UserForms (.frm).
        Les modules de classe nécessitent un traitement spécial pour extraire
        les attributs VB_Name et VB_PredeclaredId.

        Args:
            module_file: Chemin du fichier .bas, .cls ou .frm à importer
            module_type: Type forcé du module. Si None, auto-détecté
            workbook: Classeur cible. Si None, utilise le classeur actif
            overwrite: Si True, supprime le module existant avant import

        Returns:
            VBAModuleInfo: Informations sur le module importé

        Raises:
            VBAImportError: Fichier invalide, encodage incorrect, ou .frx manquant
            VBAModuleAlreadyExistsError: Module existe et overwrite=False
            VBAProjectAccessError: Trust Center refuse l'accès au VBProject
            VBAWorkbookFormatError: Classeur au format .xlsx
            WorkbookNotFoundError: Classeur non ouvert

        Example:
            >>> vba_mgr.import_module(Path("Module1.bas"))
            VBAModuleInfo(name='Module1', module_type='standard', ...)

            >>> # Importer avec remplacement
            >>> vba_mgr.import_module(Path("MyClass.cls"), overwrite=True)
        """
        # Vérifier que le fichier existe
        if not module_file.exists():
            raise VBAImportError(str(module_file), "Fichier introuvable")

        # Détection automatique du type si non fourni
        if module_type is None:
            module_type = _detect_module_type(module_file)

        # Résoudre le classeur cible
        from .worksheet_manager import _resolve_workbook

        wb = _resolve_workbook(self.app, workbook)

        # Accéder au VBProject (raise si Trust Center bloque)
        vb_project = _get_vba_project(wb)

        # Router vers la méthode appropriée selon le type
        if module_type == "standard":
            return self._import_standard_module(vb_project, module_file, overwrite)
        elif module_type == "class":
            return self._import_class_module(vb_project, module_file, overwrite)
        elif module_type == "userform":
            return self._import_userform_module(vb_project, module_file, overwrite)
        else:
            raise VBAImportError(
                str(module_file), f"Type de module '{module_type}' non supporté"
            )

    def _import_standard_module(
        self, vb_project: CDispatch, module_file: Path, overwrite: bool
    ) -> VBAModuleInfo:
        """Importe un module standard (.bas).

        Args:
            vb_project: Objet COM VBProject
            module_file: Chemin du fichier .bas
            overwrite: Si True, remplace le module existant

        Returns:
            VBAModuleInfo du module importé

        Raises:
            VBAModuleAlreadyExistsError: Si overwrite=False et module existe
            VBAImportError: Si l'import COM échoue
        """
        try:
            # Import direct via VBComponents.Import()
            component = vb_project.VBComponents.Import(str(module_file.resolve()))

            # Récupérer le nom du module importé
            module_name = component.Name

            # Vérifier si un module avec ce nom existe déjà (sauf si overwrite)
            if not overwrite:
                existing = _find_component(vb_project, module_name)
                if existing is not None and existing != component:
                    # Annuler l'import (supprimer le module importé)
                    vb_project.VBComponents.Remove(component)
                    raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)

            # Si overwrite et module existait, l'import l'a écrasé automatiquement

            # Construire VBAModuleInfo
            lines_count = component.CodeModule.CountOfLines
            return VBAModuleInfo(
                name=module_name,
                module_type="standard",
                lines_count=lines_count,
                has_predeclared_id=False,
            )

        except pywintypes.com_error as e:
            raise VBAImportError(str(module_file), f"Erreur COM: {e}") from e

    def _import_class_module(
        self, vb_project: CDispatch, module_file: Path, overwrite: bool
    ) -> VBAModuleInfo:
        """Importe un module de classe (.cls) avec parsing des attributs.

        Les modules .cls contiennent des attributs (VB_Name, VB_PredeclaredId)
        qu'il faut extraire manuellement car Import() ne les gère pas correctement.

        Args:
            vb_project: Objet COM VBProject
            module_file: Chemin du fichier .cls
            overwrite: Si True, remplace le module existant

        Returns:
            VBAModuleInfo du module importé

        Raises:
            VBAModuleAlreadyExistsError: Si overwrite=False et module existe
            VBAImportError: Si le parsing échoue
        """
        # Parser le fichier .cls pour extraire les métadonnées
        module_name, predeclared_id, code_content = _parse_class_module(module_file)

        # Vérifier si le module existe déjà
        existing = _find_component(vb_project, module_name)
        if existing is not None:
            if not overwrite:
                raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)
            # Supprimer l'ancien module
            vb_project.VBComponents.Remove(existing)
            del existing

        try:
            # Créer un nouveau module de classe (type 2)
            component = vb_project.VBComponents.Add(VBEXT_CT_CLASS_MODULE)

            # Définir le nom
            component.Name = module_name

            # Définir PredeclaredId si nécessaire
            if predeclared_id:
                component.Properties("PredeclaredId").Value = True

            # Ajouter le code source
            if code_content:
                component.CodeModule.AddFromString(code_content)

            # Construire VBAModuleInfo
            lines_count = component.CodeModule.CountOfLines
            return VBAModuleInfo(
                name=module_name,
                module_type="class",
                lines_count=lines_count,
                has_predeclared_id=predeclared_id,
            )

        except pywintypes.com_error as e:
            raise VBAImportError(str(module_file), f"Erreur COM: {e}") from e

    def _import_userform_module(
        self, vb_project: CDispatch, module_file: Path, overwrite: bool
    ) -> VBAModuleInfo:
        """Importe un UserForm (.frm + .frx).

        Args:
            vb_project: Objet COM VBProject
            module_file: Chemin du fichier .frm
            overwrite: Si True, remplace le UserForm existant

        Returns:
            VBAModuleInfo du UserForm importé

        Raises:
            VBAModuleAlreadyExistsError: Si overwrite=False et UserForm existe
            VBAImportError: Si le fichier .frx est manquant ou l'import échoue
        """
        # Vérifier que le fichier .frx existe (obligatoire pour les UserForms)
        frx_file = module_file.with_suffix(".frx")
        if not frx_file.exists():
            raise VBAImportError(
                str(module_file), f"Fichier .frx manquant : {frx_file}"
            )

        try:
            # Import direct via VBComponents.Import()
            component = vb_project.VBComponents.Import(str(module_file.resolve()))

            # Récupérer le nom du UserForm
            module_name = component.Name

            # Vérifier si un UserForm avec ce nom existe déjà
            if not overwrite:
                existing = _find_component(vb_project, module_name)
                if existing is not None and existing != component:
                    # Annuler l'import
                    vb_project.VBComponents.Remove(component)
                    raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)

            # Construire VBAModuleInfo
            lines_count = component.CodeModule.CountOfLines
            return VBAModuleInfo(
                name=module_name,
                module_type="userform",
                lines_count=lines_count,
                has_predeclared_id=True,  # UserForms ont toujours PredeclaredId=True
            )

        except pywintypes.com_error as e:
            raise VBAImportError(str(module_file), f"Erreur COM: {e}") from e
