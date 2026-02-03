VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Patterns
   Caption         =   "Advanced UserForm Example"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "UserForm_Patterns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Patterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' ADVANCED USERFORM PATTERN - MVVM ARCHITECTURE
' =============================================================================
' This UserForm demonstrates enterprise-grade patterns including:
' - Model-View-ViewModel (MVVM) architecture
' - Dynamic control creation
' - Validation framework
' - Event delegation
' - Data binding helpers
' =============================================================================

' =============================================================================
' VIEW MODEL CLASS
' =============================================================================
' Create a class module named "CUserFormViewModel" with this code:

' CUserFormViewModel.cls
' Option Explicit
'
' Private mView As UserForm_Patterns
' Private mData As Object 'Scripting.Dictionary
' Private mValidators As Collection
'
' Public Sub Initialize(view As UserForm_Patterns)
'     Set mView = view
'     Set mData = CreateObject("Scripting.Dictionary")
'     Set mValidators = New Collection
'     SetupValidators
' End Sub
'
' Private Sub SetupValidators()
'     ' Add validation rules
'     AddValidator "txtName", "Required", True
'     AddValidator "txtName", "MinLength", 2
'     AddValidator "txtEmail", "Required", True
'     AddValidator "txtEmail", "Pattern", "^[^\s@]+@[^\s@]+\.[^\s@]+$"
'     AddValidator "txtAge", "Numeric", True
'     AddValidator "txtAge", "Range", Array(18, 120)
' End Sub
'
' Public Sub AddValidator(controlName As String, ruleType As String, ruleValue As Variant)
'     Dim validator As Object
'     Set validator = CreateObject("Scripting.Dictionary")
'     validator("Control") = controlName
'     validator("Type") = ruleType
'     validator("Value") = ruleValue
'     mValidators.Add validator
' End Sub
'
' Public Function Validate() As Boolean
'     Dim validator As Variant
'     Dim isValid As Boolean
'     isValid = True
'
'     For Each validator In mValidators
'         If Not ValidateControl(validator) Then
'             isValid = False
'         End If
'     Next validator
'
'     Validate = isValid
' End Function
'
' Private Function ValidateControl(validator As Object) As Boolean
'     Dim controlName As String
'     Dim ruleType As String
'     Dim ruleValue As Variant
'     Dim control As Control
'
'     controlName = validator("Control")
'     ruleType = validator("Type")
'     ruleValue = validator("Value")
'
'     On Error Resume Next
'     Set control = mView.Controls(controlName)
'     On Error GoTo 0
'
'     If control Is Nothing Then
'         ValidateControl = True
'         Exit Function
'     End If
'
'     Dim value As String
'     value = control.Value
'
'     Select Case ruleType
'         Case "Required"
'             If ruleValue And Trim(value) = "" Then
'                 HighlightError control, "This field is required"
'                 ValidateControl = False
'             Else
'                 ClearError control
'                 ValidateControl = True
'             End If
'
'         Case "MinLength"
'             If Len(value) < ruleValue Then
'                 HighlightError control, "Minimum " & ruleValue & " characters required"
'                 ValidateControl = False
'             Else
'                 ClearError control
'                 ValidateControl = True
'             End If
'
'         Case "Pattern"
'             If Not value Like ruleValue Then
'                 HighlightError control, "Invalid format"
'                 ValidateControl = False
'             Else
'                 ClearError control
'                 ValidateControl = True
'             End If
'
'         Case "Numeric"
'             If ruleValue And Not IsNumeric(value) Then
'                 HighlightError control, "Must be a number"
'                 ValidateControl = False
'             Else
'                 ClearError control
'                 ValidateControl = True
'             End If
'
'         Case "Range"
'             If Not IsNumeric(value) Then
'                 ValidateControl = False
'             Else
'                 Dim numValue As Double
'                 numValue = CDbl(value)
'                 If numValue < ruleValue(0) Or numValue > ruleValue(1) Then
'                     HighlightError control, "Must be between " & ruleValue(0) & " and " & ruleValue(1)
'                     ValidateControl = False
'                 Else
'                     ClearError control
'                     ValidateControl = True
'                 End If
'             End If
'     End Select
' End Function
'
' Private Sub HighlightError(control As Control, message As String)
'     control.BackColor = RGB(255, 200, 200)
'     ' Could also show tooltip or label
' End Sub
'
' Private Sub ClearError(control As Control)
'     control.BackColor = RGB(255, 255, 255)
' End Sub
'
' Public Sub SaveData()
'     ' Save to data dictionary
'     mData("Name") = mView.txtName.Value
'     mData("Email") = mView.txtEmail.Value
'     mData("Age") = mView.txtAge.Value
'     mData("Department") = mView.cboDepartment.Value
' End Sub
'
' Public Property Get Data() As Object
'     Set Data = mData
' End Property

' =============================================================================
' USERFORM CODE
' =============================================================================

Private mViewModel As CUserFormViewModel
Private mDynamicControls As Collection

' =============================================================================
' INITIALIZATION
' =============================================================================

Private Sub UserForm_Initialize()
    ' Initialize ViewModel
    Set mViewModel = New CUserFormViewModel
    mViewModel.Initialize Me

    ' Initialize dynamic controls collection
    Set mDynamicControls = New Collection

    ' Setup UI
    SetupUI

    ' Load initial data
    LoadDepartments
End Sub

Private Sub SetupUI()
    ' Configure form appearance
    Me.Caption = "Employee Data Entry"
    Me.Height = 400
    Me.Width = 500

    ' Position controls
    PositionControls
End Sub

Private Sub PositionControls()
    ' Example of programmatic control positioning
    With lblName
        .Left = 20
        .Top = 20
        .Width = 100
    End With

    With txtName
        .Left = 130
        .Top = 20
        .Width = 300
    End With

    With lblEmail
        .Left = 20
        .Top = 50
        .Width = 100
    End With

    With txtEmail
        .Left = 130
        .Top = 50
        .Width = 300
    End With

    ' Continue for other controls...
End Sub

Private Sub LoadDepartments()
    ' Populate combo box with data
    With cboDepartment
        .Clear
        .AddItem "Sales"
        .AddItem "Marketing"
        .AddItem "IT"
        .AddItem "HR"
        .AddItem "Finance"
        .ListIndex = 0
    End With
End Sub

' =============================================================================
' DYNAMIC CONTROL CREATION
' =============================================================================

Public Sub AddDynamicButton(ByVal caption As String, _
                           ByVal top As Long, _
                           ByVal left As Long)
    'Dynamically creates a button at runtime

    Dim btn As MSForms.CommandButton

    Set btn = Me.Controls.Add("Forms.CommandButton.1", _
                             "btnDynamic" & mDynamicControls.Count + 1, _
                             True)

    With btn
        .caption = caption
        .top = top
        .left = left
        .Width = 100
        .Height = 25
        .Visible = True
    End With

    ' Store reference
    mDynamicControls.Add btn

    ' Wire up event (requires class-based event handling for dynamic controls)
    ' See CButtonEvents class below
End Sub

' =============================================================================
' EVENT HANDLERS
' =============================================================================

Private Sub btnSave_Click()
    'Validate and save

    If mViewModel.Validate Then
        mViewModel.SaveData

        ' Raise event to notify caller
        RaiseEvent DataSaved(mViewModel.Data)

        Me.Hide
    Else
        MsgBox "Please correct the errors before saving.", vbExclamation
    End If
End Sub

Private Sub btnCancel_Click()
    'Cancel and close
    Me.Hide
End Sub

Private Sub btnAddDynamic_Click()
    'Example: Add dynamic control
    AddDynamicButton "New Button", 300, 20
End Sub

Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Real-time validation on exit
    mViewModel.ValidateControlByName "txtName"
End Sub

Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Real-time validation on exit
    mViewModel.ValidateControlByName "txtEmail"
End Sub

Private Sub txtAge_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Real-time validation on exit
    mViewModel.ValidateControlByName "txtAge"
End Sub

' =============================================================================
' PUBLIC INTERFACE
' =============================================================================

Public Event DataSaved(data As Object)

Public Sub ShowModal()
    'Show form modally with proper initialization
    Me.Show vbModal
End Sub

Public Function ShowDialog() As Object
    'Show form and return data
    Me.Show vbModal
    Set ShowDialog = mViewModel.Data
End Function

Public Property Get FormData() As Object
    Set FormData = mViewModel.Data
End Property

' =============================================================================
' DYNAMIC EVENT HANDLER CLASS
' =============================================================================
' Create a class module named "CButtonEvents" for handling dynamic buttons:

' CButtonEvents.cls
' Option Explicit
'
' Public WithEvents btn As MSForms.CommandButton
'
' Private Sub btn_Click()
'     MsgBox "Dynamic button clicked: " & btn.caption
' End Sub

' =============================================================================
' CLEANUP
' =============================================================================

Private Sub UserForm_Terminate()
    ' Cleanup
    Set mViewModel = Nothing
    Set mDynamicControls = Nothing
End Sub
