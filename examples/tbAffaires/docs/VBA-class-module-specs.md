# Spécifications des Modules de Classe VBA

## Format Physique des Fichiers .cls

Les modules de classe VBA sont stockés dans des fichiers avec l'extension `.cls` et doivent respecter un format spécifique au niveau physique.

### En-tête d'Attributs

Chaque fichier `.cls` commence par un ensemble d'attributs qui définissent les caractéristiques linguistiques du module :

```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NomClasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
```

#### Description des Attributs

- `VERSION 1.0 CLASS` : Indique qu'il s'agit d'une définition de classe VBA
- `MultiUse` : Permet d'utiliser la classe plusieurs fois dans un projet
  - `-1` signifie `True` (peut être instancié plusieurs fois)
  - `0` signifie `False`
- `VB_Name` : Nom de la classe tel qu'il apparaît dans l'EDI VBA
- `VB_GlobalNameSpace` : Indique si la classe est accessible globalement
  - Doit être `False` dans les projets VBA standards
- `VB_Creatable` : Indique si la classe peut être créée depuis d'autres projets
  - Doit être `False` dans les projets VBA standards
- `VB_PredeclaredId` : Indique s'il existe une instance par défaut de la classe
- `VB_Exposed` : Indique si la classe est exposée à d'autres applications

### Structure du Corps de la Classe

Après l'en-tête d'attributs, le corps de la classe suit la structure suivante :

```
Option Explicit

' Variables membres
Private m_Variable As Type

' Propriétés
Public Property Get NomPropriete() As Type
    ' Implémentation
End Property

' Méthodes
Private Sub Class_Initialize()
    ' Code d'initialisation
End Sub

Private Sub Class_Terminate()
    ' Code de nettoyage
End Sub

Public Sub NomMethode()
    ' Implémentation
End Sub
```

### Règles de Codage

1. **Option Explicit** : Doit toujours être présent en haut du module
2. **Codage** : Fichiers encodés en Windows-1252 avec fins de ligne CRLF (`\r\n`)
3. **Nommage** : Respecter le PascalCase en français
4. **Structure** : Respecter l'ordre (en-tête, Option Explicit, variables, propriétés, événements, méthodes)

## Différences avec les Modules Standard

Les modules de classe (`.cls`) diffèrent des modules standard (`.bas`) par :

- La présence obligatoire d'un en-tête d'attributs
- La possibilité de définir des constructeurs (`Class_Initialize`) et destructeurs (`Class_Terminate`)
- La possibilité de créer des objets multiples à partir de la classe
- L'utilisation de propriétés en plus des procédures
- Le support des événements personnalisés

## Mise en Œuvre des Meilleures Pratiques

Lors de la création de modules de classe, il est important de respecter les modèles de conception tels que le pattern RAII (Resource Acquisition Is Initialization) pour la gestion des ressources système, ce qui permet d'assurer la stabilité et les performances de l'application Excel.
