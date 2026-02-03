' Classe pour gérer l'état Excel (RAII)
' Préfixe cls pour les classes

Private Sub Class_Initialize()
    ' Désactiver les mises à jour d'écran, calculs, et événements
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    EnregistrerLog "Initialisation", "SUCCES"
End Sub

Private Sub Class_Terminate()
    ' Restaurer l'état initial
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    EnregistrerLog "Fermeture", "SUCCES"
End Sub
