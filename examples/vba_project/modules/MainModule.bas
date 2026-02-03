Attribute VB_Name = "MainModule"
' ============================================================================
' MainModule - Module principal de démonstration xlManage
'
' Ce module contient des procédures de test pour xlManage
' ============================================================================

' Constantes
Public Const APP_NAME As String = "xlManage Demo"
Public Const VERSION As String = "1.0"

' ============================================================================
' Procédures Sub (sans retour)
' ============================================================================

Public Sub HelloWorld()
    ' Affiche un message de bienvenue
    MsgBox "Hello from xlManage!", vbInformation, APP_NAME
End Sub

Public Sub ProcessData()
    ' Simule un traitement de données
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)

    ' Ajoute des données de test
    ws.Range("A1").Value = "Produit"
    ws.Range("B1").Value = "Quantité"
    ws.Range("C1").Value = "Prix"
    ws.Range("D1").Value = "Total"

    ' Remplit quelques lignes
    ws.Range("A2").Value = "Produit A"
    ws.Range("B2").Value = 10
    ws.Range("C2").Value = 25.5
    ws.Range("D2").Formula = "=B2*C2"

    ws.Range("A3").Value = "Produit B"
    ws.Range("B3").Value = 5
    ws.Range("C3").Value = 40
    ws.Range("D3").Formula = "=B3*C3"

    MsgBox "Données créées avec succès!", vbInformation, APP_NAME
End Sub

Public Sub FormatTable()
    ' Formate la plage de données comme tableau
    Dim tbl As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)

    ' Supprime l'ancien tableau s'il existe
    On Error Resume Next
    ws.ListObjects("TableauDemo").Delete
    On Error GoTo 0

    ' Crée un nouveau tableau
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D3"), , xlYes)
    tbl.Name = "TableauDemo"
    tbl.TableStyle = "TableStyleMedium2"

    MsgBox "Tableau formaté!", vbInformation, APP_NAME
End Sub

Public Sub ClearData()
    ' Efface les données de la feuille
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear
    MsgBox "Données effacées!", vbInformation, APP_NAME
End Sub

Public Sub TestWithParams(param1 As String, param2 As Integer)
    ' Test avec paramètres
    MsgBox "Paramètre 1: " & param1 & vbCrLf & _
           "Paramètre 2: " & param2, vbInformation, APP_NAME
End Sub

' ============================================================================
' Procédures Function (avec retour)
' ============================================================================

Public Function CalculateTotal(quantity As Integer, price As Double) As Double
    ' Calcule un total
    CalculateTotal = quantity * price
End Function

Public Function GetProductName(index As Integer) As String
    ' Retourne un nom de produit basé sur l'index
    Dim products As Variant
    products = Array("Produit A", "Produit B", "Produit C", "Produit D")

    If index >= 0 And index < UBound(products) + 1 Then
        GetProductName = products(index)
    Else
        GetProductName = "Inconnu"
    End If
End Function

Public Function ConcatenateStrings(str1 As String, str2 As String, Optional separator As String = " ") As String
    ' Concatène deux chaînes avec séparateur
    ConcatenateStrings = str1 & separator & str2
End Function

' ============================================================================
' Procédure d'initialisation
' ============================================================================

Public Sub InitializeWorkbook()
    ' Initialise le classeur pour les tests
    Dim ws As Worksheet

    ' Renomme la première feuille
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Name = "Données"

    ' Crée une deuxième feuille
    On Error Resume Next
    ThisWorkbook.Worksheets("Résultats").Delete
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Résultats"

    MsgBox "Classeur initialisé!", vbInformation, APP_NAME
End Sub
