VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DemoForm
   Caption         =   "xlManage Demo Form"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "DemoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DemoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' DemoForm - UserForm de démonstration xlManage
'
' Exemple de UserForm pour tester l'import/export avec xlManage
' ============================================================================

Private Sub UserForm_Initialize()
    ' Initialisation du formulaire
    Me.Caption = MainModule.APP_NAME & " - Demo"
    lblVersion.Caption = "Version: " & MainModule.VERSION

    ' Remplit la liste déroulante
    cmbOptions.Clear
    cmbOptions.AddItem "Option 1 - Créer données"
    cmbOptions.AddItem "Option 2 - Formater tableau"
    cmbOptions.AddItem "Option 3 - Effacer données"
    cmbOptions.ListIndex = 0
End Sub

Private Sub btnExecute_Click()
    ' Exécute l'option sélectionnée
    Select Case cmbOptions.ListIndex
        Case 0
            MainModule.ProcessData
        Case 1
            MainModule.FormatTable
        Case 2
            MainModule.ClearData
    End Select
End Sub

Private Sub btnClose_Click()
    ' Ferme le formulaire
    Unload Me
End Sub

Private Sub btnTest_Click()
    ' Bouton de test
    Dim result As String
    result = MainModule.ConcatenateStrings(txtInput1.Value, txtInput2.Value, " - ")
    MsgBox "Résultat: " & result, vbInformation, "Test"
End Sub
