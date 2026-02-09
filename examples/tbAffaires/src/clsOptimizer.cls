VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsOptimizer.cls - Gestionnaire d'optimisation Excel tbAffaires
'
'-------------------------------------------------------------------------------
' GPL v3 - LICENSE
'-------------------------------------------------------------------------------
'
' Auteur : Equipe tbAffaires
' Date : 2026-02-08
' Objet : Classe pour gérer l'optimisation des performances Excel via le pattern RAII
' Liste des éléments public : clsOptimizer (classe complète)

Option Explicit

' Variables membres pour sauvegarder l'état initial d'Excel
Private m_ScreenUpdating As Boolean
Private m_Calculation As XlCalculation
Private m_EnableEvents As Boolean
Private m_DisplayAlerts As Boolean

'-------------------------------------------------------------------------------
' Class_Initialize - Initialise l'optimisation Excel en désactivant les fonctions non essentielles
' Parameters : None
' Return     : None
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ' Sauvegarder l'état actuel d'Excel
    m_ScreenUpdating = Application.ScreenUpdating
    m_Calculation = Application.Calculation
    m_EnableEvents = Application.EnableEvents
    m_DisplayAlerts = Application.DisplayAlerts

    ' Appliquer les optimisations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Journaliser l'initialisation
    Call modLogging.EnregistrerSucces("Initialisation Optimizer", "RAII activé")
End Sub

'-------------------------------------------------------------------------------
' Class_Terminate - Restaure l'état original d'Excel lors de la destruction de l'objet
' Parameters : None
' Return     : None
'-------------------------------------------------------------------------------
Private Sub Class_Terminate()
    On Error Resume Next

    ' Restaurer l'état initial d'Excel
    Application.ScreenUpdating = m_ScreenUpdating
    Application.Calculation = m_Calculation
    Application.EnableEvents = m_EnableEvents
    Application.DisplayAlerts = m_DisplayAlerts

    ' Journaliser la fermeture
    Call modLogging.EnregistrerSucces("Fermeture Optimizer", "RAII terminé")

    Err.Clear
End Sub

'-------------------------------------------------------------------------------
' GetOriginalScreenUpdating - Retourne l'état original de ScreenUpdating
' Parameters : None
' Return     : Boolean - Etat original de ScreenUpdating
'-------------------------------------------------------------------------------
Public Property Get GetOriginalScreenUpdating() As Boolean
    GetOriginalScreenUpdating = m_ScreenUpdating
End Property

'-------------------------------------------------------------------------------
' GetOriginalCalculation - Retourne l'état original de Calculation
' Parameters : None
' Return     : XlCalculation - Etat original de Calculation
'-------------------------------------------------------------------------------
Public Property Get GetOriginalCalculation() As XlCalculation
    GetOriginalCalculation = m_Calculation
End Property

'-------------------------------------------------------------------------------
' GetOriginalEnableEvents - Retourne l'état original de EnableEvents
' Parameters : None
' Return     : Boolean - Etat original de EnableEvents
'-------------------------------------------------------------------------------
Public Property Get GetOriginalEnableEvents() As Boolean
    GetOriginalEnableEvents = m_EnableEvents
End Property

'-------------------------------------------------------------------------------
' GetOriginalDisplayAlerts - Retourne l'état original de DisplayAlerts
' Parameters : None
' Return     : Boolean - Etat original de DisplayAlerts
'-------------------------------------------------------------------------------
Public Property Get GetOriginalDisplayAlerts() As Boolean
    GetOriginalDisplayAlerts = m_DisplayAlerts
End Property

'-------------------------------------------------------------------------------
' ForceOptimization - Applique manuellement les optimisations (utilisé si besoin)
' Parameters : None
' Return     : None
'-------------------------------------------------------------------------------
Public Sub ForceOptimization()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

'-------------------------------------------------------------------------------
' RestoreState - Restaure manuellement l'état original d'Excel
' Parameters : None
' Return     : None
'-------------------------------------------------------------------------------
Public Sub RestoreState()
    Application.ScreenUpdating = m_ScreenUpdating
    Application.Calculation = m_Calculation
    Application.EnableEvents = m_EnableEvents
    Application.DisplayAlerts = m_DisplayAlerts
End Sub
