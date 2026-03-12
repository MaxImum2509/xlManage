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

import logging
import re
import shutil
import tempfile
import time
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

logger = logging.getLogger(__name__)


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

# BOM UTF-8
_UTF8_BOM: bytes = b"\xef\xbb\xbf"


@dataclass
class EncodingCheckResult:
    """Result of VBA file encoding check.

    Attributes:
        original_path: Path to the source file.
        effective_path: Path to use for import (may be a temp file).
        was_converted: True if the file was converted.
        source_encoding: Detected source encoding.
        had_wrong_line_endings: True if line endings were not CRLF.
    """

    original_path: Path
    effective_path: Path
    was_converted: bool
    source_encoding: str
    had_wrong_line_endings: bool


def _detect_file_encoding(raw: bytes) -> str:
    """Detect encoding of raw VBA file bytes.

    Strategy (in order):
      1. UTF-8 BOM present -> "utf-8-sig"
      2. Decodable as UTF-8 **and** contains bytes > 127 -> "utf-8"
      3. Otherwise -> "windows-1252" (already compliant)

    Args:
        raw: Raw file bytes.

    Returns:
        Detected encoding name usable with ``open(encoding=...)``.
    """
    if raw.startswith(_UTF8_BOM):
        return "utf-8-sig"

    # Check if the content is valid UTF-8 with high bytes
    has_high_bytes = any(b > 127 for b in raw)
    if has_high_bytes:
        try:
            raw.decode("utf-8")
            return "utf-8"
        except UnicodeDecodeError:
            pass

    return "windows-1252"


def _has_wrong_line_endings(raw: bytes) -> bool:
    """Check whether raw bytes contain LF without preceding CR.

    Args:
        raw: Raw file bytes.

    Returns:
        True if at least one bare ``\\n`` (not preceded by ``\\r``) is found.
    """
    i = raw.find(b"\n")
    while i != -1:
        if i == 0 or raw[i - 1 : i] != b"\r":
            return True
        i = raw.find(b"\n", i + 1)
    return False


def _ensure_vba_encoding(module_file: Path) -> EncodingCheckResult:
    """Ensure a VBA source file is Windows-1252 with CRLF line endings.

    If the file is already compliant, returns the original path.
    Otherwise, reads with the detected encoding, converts to
    Windows-1252 + CRLF, writes a temporary file, and returns
    that path.

    Args:
        module_file: Path to the VBA source file (.bas, .cls, .frm).

    Returns:
        EncodingCheckResult with the effective path to use for import.

    Raises:
        VBAImportError: If the file contains characters that cannot be
            represented in Windows-1252.
    """
    raw = module_file.read_bytes()
    detected = _detect_file_encoding(raw)
    wrong_endings = _has_wrong_line_endings(raw)

    # Already compliant?
    if detected == "windows-1252" and not wrong_endings:
        return EncodingCheckResult(
            original_path=module_file,
            effective_path=module_file,
            was_converted=False,
            source_encoding=detected,
            had_wrong_line_endings=False,
        )

    # Need conversion
    try:
        text = raw.decode(detected)
    except UnicodeDecodeError as e:
        raise VBAImportError(
            str(module_file),
            f"Impossible de decoder le fichier en {detected}: {e}",
        ) from e

    try:
        encoded = text.encode("windows-1252")
    except UnicodeEncodeError as e:
        raise VBAImportError(
            str(module_file),
            f"Le fichier contient des caracteres non representables "
            f"en Windows-1252 (position {e.start}): {e.reason}",
        ) from e

    # Normalize line endings to CRLF
    # First remove any existing \r to avoid \r\r\n, then replace \n with \r\n
    normalized = (
        encoded.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
    )

    # Write to a temp file in the same directory (same extension required)
    suffix = module_file.suffix
    tmp = tempfile.NamedTemporaryFile(
        suffix=suffix,
        prefix=f"{module_file.stem}_",
        dir=module_file.parent,
        delete=False,
    )
    tmp.write(normalized)
    tmp.close()

    tmp_path = Path(tmp.name)

    # For UserForms (.frm), copy the companion .frx alongside the temp .frm
    if suffix.lower() == ".frm":
        frx_source = module_file.with_suffix(".frx")
        if not frx_source.exists():
            raise VBAImportError(
                str(module_file), f"Fichier .frx manquant : {frx_source}"
            )
        frx_dest = tmp_path.with_suffix(".frx")
        shutil.copy2(frx_source, frx_dest)

    logger.info(
        "Converted %s from %s to windows-1252/CRLF -> %s",
        module_file.name,
        detected,
        tmp.name,
    )

    return EncodingCheckResult(
        original_path=module_file,
        effective_path=tmp_path,
        was_converted=True,
        source_encoding=detected,
        had_wrong_line_endings=wrong_endings,
    )


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
    """Détecte le type de module depuis l'extension et le contenu.

    Pour les fichiers .cls, distingue un module de classe d'un module de
    document (ThisWorkbook, Sheet) en inspectant les attributs
    ``VB_PredeclaredId`` et ``VB_Exposed``.

    Args:
        path: Chemin du fichier module VBA

    Returns:
        str: Type du module ("standard", "class", "document", "userform")

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

    base_type = EXTENSION_TO_TYPE[extension]

    # Pour les .cls, distinguer class vs document
    if base_type == "class" and _is_document_module(path):
        return "document"

    return base_type


def _is_document_module(file_path: Path) -> bool:
    """Détermine si un fichier .cls est un module de document.

    Un module de document (ThisWorkbook, Sheet) possède à la fois
    ``VB_PredeclaredId = True`` et ``VB_Exposed = True``, contrairement
    aux modules de classe ordinaires.

    Args:
        file_path: Chemin du fichier .cls

    Returns:
        bool: True si le fichier décrit un module de document
    """
    try:
        content = file_path.read_text(encoding=VBA_ENCODING)
    except UnicodeDecodeError:
        return False

    has_predeclared = bool(re.search(r"Attribute VB_PredeclaredId\s*=\s*True", content))
    has_exposed = bool(re.search(r"Attribute VB_Exposed\s*=\s*True", content))
    return has_predeclared and has_exposed


def _parse_standard_module_name(file_path: Path) -> str:
    """Parse un fichier .bas pour extraire le nom du module (VB_Name).

    Si ``Attribute VB_Name`` est absent, utilise le stem du fichier.

    Args:
        file_path: Chemin du fichier .bas

    Returns:
        str: Nom du module extrait de VB_Name, ou stem du fichier.

    Raises:
        VBAImportError: Si le fichier est illisible (encodage invalide)
    """
    try:
        content = file_path.read_text(encoding="windows-1252")
    except UnicodeDecodeError as e:
        raise VBAImportError(
            str(file_path),
            f"Encodage invalide. Les fichiers VBA doivent être en windows-1252 : {e}",
        ) from e

    name_match = re.search(r'Attribute VB_Name = "([^"]+)"', content)
    if name_match:
        return name_match.group(1)

    logger.warning(
        "Attribut VB_Name absent dans '%s', utilisation du nom de fichier : '%s'",
        file_path.name,
        file_path.stem,
    )
    return file_path.stem


def _parse_userform_name(file_path: Path) -> str:
    """Parse un fichier .frm pour extraire le nom du UserForm.

    Cherche d'abord ``Attribute VB_Name``, puis la ligne
    ``Begin {CLSID} FormName`` si l'attribut est absent.

    Args:
        file_path: Chemin du fichier .frm

    Returns:
        str: Nom du UserForm

    Raises:
        VBAImportError: Si le nom ne peut pas être extrait
    """
    try:
        content = file_path.read_text(encoding="windows-1252")
    except UnicodeDecodeError as e:
        raise VBAImportError(
            str(file_path),
            f"Encodage invalide. Les fichiers VBA doivent être en windows-1252 : {e}",
        ) from e

    # Try Attribute VB_Name first
    name_match = re.search(r'Attribute VB_Name = "([^"]+)"', content)
    if name_match:
        return name_match.group(1)

    # Fallback: parse Begin {CLSID} FormName header
    begin_match = re.search(r"^Begin\s+\{[^}]+\}\s+(\w+)", content, re.MULTILINE)
    if begin_match:
        return begin_match.group(1)

    raise VBAImportError(
        str(file_path),
        "Impossible d'extraire le nom du UserForm (ni VB_Name ni Begin header)",
    )


def _parse_class_module(file_path: Path) -> tuple[str, bool, str]:
    """Parse un fichier .cls pour extraire les mtadonnes.

    Les fichiers .cls commencent par des lignes "Attribute VB_Name" qu'il
    faut parser sparment avant d'importer le code.

    Args:
        file_path: Chemin du fichier .cls

    Returns:
        tuple[str, bool, str]: (module_name, predeclared_id, code_content)
            - module_name: Nom du module extrait de VB_Name
            - predeclared_id: True si VB_PredeclaredId = True
            - code_content: Code source sans les attributs d'en-tte

    Raises:
        VBAImportError: Si le fichier est invalide ou mal encod
    """
    try:
        # Lire le fichier avec l'encodage VBA (OBLIGATOIRE)
        content = file_path.read_text(encoding="windows-1252")
    except UnicodeDecodeError as e:
        raise VBAImportError(
            str(file_path),
            f"Encodage invalide. Les fichiers VBA doivent être en windows-1252 : {e}",
        ) from e

    # Extraire VB_Name — fallback sur le stem du fichier si absent
    name_match = re.search(r'Attribute VB_Name = "([^"]+)"', content)
    if name_match:
        module_name = name_match.group(1)
    else:
        module_name = file_path.stem
        logger.warning(
            "Attribut VB_Name absent dans '%s', utilisation du nom de fichier : '%s'",
            file_path.name,
            module_name,
        )

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


def _parse_document_module(file_path: Path) -> tuple[str, str]:
    """Parse un fichier .cls de module document pour extraire le nom et le code.

    Retire les en-têtes physiques (VERSION, BEGIN/END, Attribute) pour ne
    conserver que le code VBA injectable via ``CodeModule.AddFromString``.

    Args:
        file_path: Chemin du fichier .cls (module document)

    Returns:
        tuple[str, str]: (module_name, code_content)
            - module_name: Nom du module extrait de VB_Name
            - code_content: Code source sans les attributs d'en-tête

    Raises:
        VBAImportError: Si le fichier est invalide ou mal encodé
    """
    try:
        content = file_path.read_text(encoding=VBA_ENCODING)
    except UnicodeDecodeError as e:
        raise VBAImportError(
            str(file_path),
            f"Encodage invalide. Les fichiers VBA doivent être en windows-1252 : {e}",
        ) from e

    # Extraire VB_Name — fallback sur le stem du fichier si absent
    name_match = re.search(r'Attribute VB_Name = "([^"]+)"', content)
    if name_match:
        module_name = name_match.group(1)
    else:
        module_name = file_path.stem
        logger.warning(
            "Attribut VB_Name absent dans '%s', utilisation du nom de fichier : '%s'",
            file_path.name,
            module_name,
        )

    # Filtrer les en-têtes physiques pour ne garder que le code
    lines = content.splitlines()
    code_lines: list[str] = []
    for line in lines:
        stripped = line.strip()
        if re.match(r"^(VERSION\s|BEGIN$|END$|Attribute\s+VB_)", stripped):
            continue
        if "MultiUse" in stripped:
            continue
        code_lines.append(line)

    code_content = "\r\n".join(code_lines)
    return module_name, code_content


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

        Supporte les modules standard (.bas), de classe (.cls), UserForms (.frm)
        et les modules de document (.cls avec PredeclaredId+Exposed, ex:
        ThisWorkbook, Sheet1). Les modules de document sont détectés
        automatiquement et leur code est injecté via CodeModule.

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

        # Vérifier et convertir l'encodage si nécessaire
        encoding_result = _ensure_vba_encoding(module_file)
        effective_file = encoding_result.effective_path
        self._last_encoding_result: EncodingCheckResult | None = encoding_result

        try:
            # Détection automatique du type si non fourni
            if module_type is None:
                module_type = _detect_module_type(effective_file)

            # Résoudre le classeur cible
            from .worksheet_manager import _resolve_workbook

            wb = _resolve_workbook(self.app, workbook)

            # Accéder au VBProject (raise si Trust Center bloque)
            vb_project = _get_vba_project(wb)

            # Router vers la méthode appropriée selon le type
            if module_type == "standard":
                return self._import_standard_module(
                    vb_project, effective_file, overwrite
                )
            elif module_type == "class":
                return self._import_class_module(vb_project, effective_file, overwrite)
            elif module_type == "userform":
                return self._import_userform_module(
                    vb_project, effective_file, overwrite
                )
            elif module_type == "document":
                return self._import_document_module(vb_project, effective_file)
            else:
                raise VBAImportError(
                    str(module_file),
                    f"Type de module '{module_type}' non supporté",
                )
        finally:
            # Nettoyer le fichier temporaire si une conversion a eu lieu
            if encoding_result.was_converted:
                try:
                    encoding_result.effective_path.unlink(missing_ok=True)
                    if encoding_result.effective_path.suffix.lower() == ".frm":
                        frx_tmp = encoding_result.effective_path.with_suffix(".frx")
                        frx_tmp.unlink(missing_ok=True)
                except OSError:
                    pass

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
            # Parser le fichier .bas pour extraire le nom (VB_Name)
            module_name = _parse_standard_module_name(module_file)

            # Vérifier si un module avec ce nom existe déjà
            existing = _find_component(vb_project, module_name)
            if existing is not None:
                if not overwrite:
                    raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)
                # Supprimer l'ancien module
                vb_project.VBComponents.Remove(existing)
                del existing

            # Import direct via VBComponents.Import()
            component = vb_project.VBComponents.Import(str(module_file.resolve()))

            # Le nom devrait être le même, mais on le récupère quand même
            imported_name = component.Name

            # Construire VBAModuleInfo
            lines_count = component.CodeModule.CountOfLines
            return VBAModuleInfo(
                name=imported_name,
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

            # Effacer le contenu par défaut ("Option Explicit") et ajouter le code
            if code_content:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
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
            # Parser le nom du UserForm depuis le fichier .frm
            module_name = _parse_userform_name(module_file)

            # Vérifier si un UserForm avec ce nom existe déjà
            existing = _find_component(vb_project, module_name)
            if existing is not None:
                if not overwrite:
                    raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)
                # Supprimer l'ancien UserForm avant import
                vb_project.VBComponents.Remove(existing)
                del existing
                # Laisser Excel finaliser la suppression du UserForm
                time.sleep(0.5)

            # Import via VBComponents.Import()
            component = vb_project.VBComponents.Import(str(module_file.resolve()))

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

    def _import_document_module(
        self, vb_project: CDispatch, module_file: Path
    ) -> VBAModuleInfo:
        """Importe un module de document (.cls avec PredeclaredId+Exposed).

        Les modules de document (ThisWorkbook, Sheet1, etc.) sont intégrés
        au classeur et ne peuvent pas être supprimés ni recréés. Le code
        est injecté en remplaçant le contenu du CodeModule existant.

        Le paramètre ``overwrite`` n'est pas nécessaire : le remplacement
        du code est toujours le comportement attendu pour un module document.

        Args:
            vb_project: Objet COM VBProject
            module_file: Chemin du fichier .cls (module document)

        Returns:
            VBAModuleInfo du module mis à jour

        Raises:
            VBAImportError: Si le module cible est introuvable dans le projet
                ou si le parsing du fichier échoue
        """
        module_name, code_content = _parse_document_module(module_file)

        # Trouver le composant document existant dans le projet
        component = _find_component(vb_project, module_name)
        if component is None:
            raise VBAImportError(
                str(module_file),
                f"Module document '{module_name}' introuvable dans le projet. "
                f"Les modules document doivent déjà exister dans le classeur.",
            )

        try:
            # Remplacer le code existant
            code_module = component.CodeModule
            if code_module.CountOfLines > 0:
                code_module.DeleteLines(1, code_module.CountOfLines)

            if code_content.strip():
                code_module.AddFromString(code_content)

            lines_count = code_module.CountOfLines
            return VBAModuleInfo(
                name=module_name,
                module_type="document",
                lines_count=lines_count,
                has_predeclared_id=True,
            )

        except pywintypes.com_error as e:
            raise VBAImportError(str(module_file), f"Erreur COM: {e}") from e

    def export_module(
        self, module_name: str, output_file: Path, workbook: Path | None = None
    ) -> Path:
        """Exporte un module VBA vers un fichier.

        Les modules standard, classe et UserForms sont exportés via component.Export().
        Les modules de document (ThisWorkbook, Sheet1, etc.) nécessitent un export
        manuel car Excel ne supporte pas Export() pour eux.

        Args:
            module_name: Nom du module dans le projet VBA
            output_file: Chemin de destination (doit inclure l'extension)
            workbook: Classeur source. Si None, utilise le classeur actif

        Returns:
            Path: Chemin effectif du fichier exporté

        Raises:
            VBAModuleNotFoundError: Module introuvable dans le projet
            VBAExportError: Échec d'écriture ou permissions insuffisantes
            VBAProjectAccessError: Trust Center refuse l'accès

        Example:
            >>> vba_mgr.export_module("Module1", Path("backup/Module1.bas"))
            Path('backup/Module1.bas')

            >>> # Exporter un module de document
            >>> vba_mgr.export_module("ThisWorkbook", Path("ThisWorkbook.cls"))
        """
        # Résoudre le classeur
        from .worksheet_manager import _resolve_workbook

        wb = _resolve_workbook(self.app, workbook)

        # Accéder au VBProject
        vb_project = _get_vba_project(wb)

        # Trouver le composant
        component = _find_component(vb_project, module_name)
        if component is None:
            from .exceptions import VBAModuleNotFoundError

            raise VBAModuleNotFoundError(module_name, wb.Name)

        # Vérifier le type de module
        module_type_code = component.Type

        try:
            # Les modules de document (Type 100) nécessitent un export manuel
            if module_type_code == VBEXT_CT_DOCUMENT:
                return self._export_document_module(component, output_file)
            else:
                # Export standard pour les autres types
                return self._export_standard_component(component, output_file)

        except PermissionError as e:
            from .exceptions import VBAExportError

            raise VBAExportError(
                module_name, str(output_file), f"Permission refusée: {e}"
            ) from e
        except pywintypes.com_error as e:
            from .exceptions import VBAExportError

            raise VBAExportError(
                module_name, str(output_file), f"Erreur COM: {e}"
            ) from e

    def _export_standard_component(
        self, component: CDispatch, output_file: Path
    ) -> Path:
        """Exporte un composant VBA standard via component.Export().

        Args:
            component: Composant VBA à exporter
            output_file: Chemin de destination

        Returns:
            Path: Chemin du fichier exporté
        """
        # Créer le dossier parent si nécessaire
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # Export via COM
        component.Export(str(output_file.resolve()))

        return output_file

    def _export_document_module(self, component: CDispatch, output_file: Path) -> Path:
        """Exporte manuellement un module de document.

        Les modules de document (ThisWorkbook, Sheet1, etc.) ne supportent pas
        component.Export(). On doit extraire le code via CodeModule.Lines()
        et reconstruire les en-têtes Attribute VB_* pour permettre la
        ré-importation via import_module().

        Args:
            component: Module de document à exporter
            output_file: Chemin de destination

        Returns:
            Path: Chemin du fichier exporté
        """
        # Créer le dossier parent si nécessaire
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # Construire l'en-tête standard d'un module document.
        # Pour les modules document (Type 100), les attributs VB_ sont
        # intrinsèques et invariants — l'API COM ne les expose pas via
        # component.Properties (qui retourne les propriétés de l'objet
        # hôte Workbook/Worksheet). Seul component.Name est dynamique.
        module_name = component.Name
        header_lines = [
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            f'Attribute VB_Name = "{module_name}"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
        ]
        header = "\r\n".join(header_lines) + "\r\n"

        # Extraire le code source
        code_module = component.CodeModule
        line_count = code_module.CountOfLines

        if line_count > 0:
            code_content = code_module.Lines(1, line_count)
        else:
            code_content = ""

        # Écrire en-tête + code en Windows-1252 avec CRLF
        output_file.write_bytes(
            header.encode(VBA_ENCODING) + code_content.encode(VBA_ENCODING)
        )

        return output_file

    def list_modules(self, workbook: Path | None = None) -> list[VBAModuleInfo]:
        """Liste tous les modules VBA du classeur.

        Inclut tous les types de modules : standard, classe, UserForms,
        et modules de document (ThisWorkbook, Sheet1, etc.).

        Args:
            workbook: Classeur à analyser. Si None, utilise le classeur actif

        Returns:
            list[VBAModuleInfo]: Liste des modules avec leurs informations

        Raises:
            VBAProjectAccessError: Trust Center refuse l'accès
            VBAWorkbookFormatError: Classeur au format .xlsx

        Example:
            >>> modules = vba_mgr.list_modules()
            >>> for module in modules:
            ...     print(f"{module.name} ({module.module_type}): {module.lines_count}")
            Module1 (standard): 42
            MyClass (class): 15
            ThisWorkbook (document): 8
        """
        # Résoudre le classeur
        from .worksheet_manager import _resolve_workbook

        wb = _resolve_workbook(self.app, workbook)

        # Accéder au VBProject
        vb_project = _get_vba_project(wb)

        modules: list[VBAModuleInfo] = []

        # Itérer sur tous les composants VBA
        for component in vb_project.VBComponents:
            module_name = component.Name
            module_type_code = component.Type
            lines_count = component.CodeModule.CountOfLines

            # Mapper le code type vers le nom lisible
            module_type = VBA_TYPE_NAMES.get(module_type_code, "unknown")

            # Extraire PredeclaredId pour les classes
            has_predeclared_id = False
            if module_type_code == VBEXT_CT_CLASS_MODULE:
                try:
                    has_predeclared_id = component.Properties("PredeclaredId").Value
                except pywintypes.com_error:
                    # Si la propriété n'existe pas, False par défaut
                    has_predeclared_id = False

            # Créer VBAModuleInfo
            info = VBAModuleInfo(
                name=module_name,
                module_type=module_type,
                lines_count=lines_count,
                has_predeclared_id=has_predeclared_id,
            )
            modules.append(info)

        return modules

    def delete_module(
        self,
        module_name: str,
        workbook: Path | None = None,
        force: bool = False,
    ) -> None:
        """Supprime un module VBA du projet.

        Seuls les modules standard, classe et UserForms peuvent être supprimés.
        Les modules de document (ThisWorkbook, Sheet1, etc.) sont intégrés
        au classeur et ne peuvent pas être retirés.

        Args:
            module_name: Nom du module à supprimer
            workbook: Classeur cible. Si None, utilise le classeur actif
            force: Paramètre réservé (aucun dialogue dans Excel)

        Raises:
            VBAModuleNotFoundError: Module introuvable ou non supprimable
            VBAProjectAccessError: Trust Center refuse l'accès

        Example:
            >>> vba_mgr.delete_module("Module1")

            >>> # Erreur avec module de document
            >>> vba_mgr.delete_module("ThisWorkbook")
            VBAModuleNotFoundError: Cannot delete document module 'ThisWorkbook'
        """
        # Résoudre le classeur
        from .worksheet_manager import _resolve_workbook

        wb = _resolve_workbook(self.app, workbook)

        # Accéder au VBProject
        vb_project = _get_vba_project(wb)

        # Trouver le composant
        component = _find_component(vb_project, module_name)
        if component is None:
            from .exceptions import VBAModuleNotFoundError

            raise VBAModuleNotFoundError(module_name, wb.Name)

        # Vérifier le type de module
        module_type_code = component.Type

        # Les modules de document (Type 100) ne peuvent PAS être supprimés
        if module_type_code == VBEXT_CT_DOCUMENT:
            from .exceptions import VBAModuleNotFoundError

            module_type_name = VBA_TYPE_NAMES.get(module_type_code, "unknown")
            raise VBAModuleNotFoundError(
                module_name,
                wb.Name,
                reason=f"Cannot delete document module. Type: {module_type_name}",
            )

        try:
            # Supprimer le composant du projet
            vb_project.VBComponents.Remove(component)

            # Libérer la référence COM (IMPORTANT)
            del component

        except pywintypes.com_error as e:
            from .exceptions import VBAModuleNotFoundError

            raise VBAModuleNotFoundError(module_name, wb.Name) from e
