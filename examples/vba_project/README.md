# Projet VBA de Test pour xlManage

Ce répertoire contient un projet VBA complet pour tester et démontrer les fonctionnalités de xlManage.

## Structure

```
vba_project/
├── modules/              # Modules VBA source
│   ├── MainModule.bas   # Module standard avec macros
│   ├── ProductClass.cls # Module de classe
│   └── DemoForm.frm     # UserForm avec code VBA
├── output/              # Classeur généré et exports
└── demo_xlmanage.py    # Script Python de démonstration
```

## Modules VBA

### MainModule.bas
Module principal contenant :
- `HelloWorld()` - Message simple
- `ProcessData()` - Création de données test
- `FormatTable()` - Formatage en tableau Excel
- `ClearData()` - Effacement des données
- `TestWithParams(param1, param2)` - Test avec paramètres
- `CalculateTotal(quantity, price)` - Fonction avec retour
- `GetProductName(index)` - Fonction retournant une valeur
- `ConcatenateStrings(str1, str2, separator)` - Fonction avec paramètre optionnel
- `InitializeWorkbook()` - Initialisation du classeur

### ProductClass.cls
Classe VBA démontrant :
- Propriétés (Get/Let)
- Validation des données
- Méthodes de calcul
- Gestion d'erreurs

### DemoForm.frm
UserForm avec :
- Contrôles (labels, textboxes, combobox, boutons)
- Événements (Initialize, Click)
- Appel de méthodes du MainModule

## Utilisation

### 1. Importer le projet VBA dans Excel

```python
from xlmanage import ExcelManager

with ExcelManager() as excel:
    workbook = excel.create_workbook()

    # Importe les modules
    excel.vba.import_module("modules/MainModule.bas")
    excel.vba.import_module("modules/ProductClass.cls")
    excel.vba.import_module("modules/DemoForm.frm")

    workbook.save("output/demo.xlsm")
```

### 2. Exécuter des macros

```python
# Exécution simple
excel.macro.run("MainModule.HelloWorld")

# Avec paramètres
excel.macro.run("MainModule.TestWithParams", args=["test", 42])

# Fonction avec retour
result = excel.macro.run_function("MainModule.CalculateTotal", args=[10, 25.5])
print(f"Total: {result}")  # 255.0
```

### 3. Exporter les modules

```python
# Exporte tous les modules
excel.vba.export_all("output/exported_modules/")

# Exporte un module spécifique
excel.vba.export_module("MainModule", "output/MainModule.bas")
```

## Exécution de la démo

```bash
# Active l'environnement Poetry
poetry shell

# Lance la démonstration
cd examples/vba_project
python demo_xlmanage.py
```

## Tests couverts

Ce projet permet de tester :

1. **Import de modules**
   - Modules standards (.bas)
   - Modules de classe (.cls)
   - UserForms (.frm + .frx)

2. **Exécution de macros**
   - Sub sans paramètres
   - Sub avec paramètres
   - Functions avec retour
   - Paramètres optionnels

3. **Gestion des objets**
   - Workbooks
   - Worksheets
   - ListObjects (tableaux)
   - Classes VBA

4. **Export de modules**
   - Conservation des attributs
   - Fichiers binaires (.frx)

## Configuration requise

- Microsoft Excel installé
- xlManage configuré
- Accès VBA activé dans Excel (Trust Center)
