VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' ProductClass - Classe de démonstration xlManage
'
' Exemple de module de classe pour tester l'import/export
' ============================================================================

' Propriétés privées
Private m_Name As String
Private m_Price As Double
Private m_Quantity As Integer

' ============================================================================
' Propriétés publiques
' ============================================================================

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(value As String)
    m_Name = value
End Property

Public Property Get Price() As Double
    Price = m_Price
End Property

Public Property Let Price(value As Double)
    If value >= 0 Then
        m_Price = value
    Else
        Err.Raise 1001, "ProductClass", "Le prix doit être positif"
    End If
End Property

Public Property Get Quantity() As Integer
    Quantity = m_Quantity
End Property

Public Property Let Quantity(value As Integer)
    If value >= 0 Then
        m_Quantity = value
    Else
        Err.Raise 1002, "ProductClass", "La quantité doit être positive"
    End If
End Property

' ============================================================================
' Méthodes publiques
' ============================================================================

Public Function GetTotal() As Double
    ' Calcule le total pour ce produit
    GetTotal = m_Price * m_Quantity
End Function

Public Function ToString() As String
    ' Retourne une représentation texte du produit
    ToString = m_Name & " (" & m_Quantity & " x " & m_Price & ") = " & GetTotal()
End Function

Public Sub Initialize(name As String, price As Double, quantity As Integer)
    ' Initialise le produit
    Me.Name = name
    Me.Price = price
    Me.Quantity = quantity
End Sub
