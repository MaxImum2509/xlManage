VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' ExcelOptimizer
' Description : RAII class for automatic Excel performance optimization
'               Saves settings on initialization, restores on termination
' Author      : tbAffaires Team
' Date        : 2026-01-31
'-------------------------------------------------------------------------------

Option Explicit

'===============================================================================
' SECTION: Private variables
'===============================================================================

Private mScreenUpdating As Boolean
Private mCalculation As XlCalculation
Private mEnableEvents As Boolean
Private mDisplayAlerts As Boolean
Private mInitialized As Boolean

'===============================================================================
' SECTION: Class events
'===============================================================================

Private Sub Class_Initialize()
    'RAII: Save current settings
    mScreenUpdating = Application.ScreenUpdating
    mCalculation = Application.Calculation
    mEnableEvents = Application.EnableEvents
    mDisplayAlerts = Application.DisplayAlerts
    mInitialized = False
End Sub

Private Sub Class_Terminate()
    'RAII: Restore settings automatically
    On Error Resume Next
    If mInitialized Then
        Application.ScreenUpdating = mScreenUpdating
        Application.Calculation = mCalculation
        Application.EnableEvents = mEnableEvents
        Application.DisplayAlerts = mDisplayAlerts
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' SECTION: Public methods
'===============================================================================

Public Sub Initialize()
    'Activate optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    mInitialized = True
End Sub

Public Sub Restore()
    'Manual restore (optional)
    If mInitialized Then
        Application.ScreenUpdating = mScreenUpdating
        Application.Calculation = mCalculation
        Application.EnableEvents = mEnableEvents
        Application.DisplayAlerts = mDisplayAlerts
        mInitialized = False
    End If
End Sub
