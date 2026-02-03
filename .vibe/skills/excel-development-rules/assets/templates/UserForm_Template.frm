VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplate
   Caption         =   "Form Title"
   ClientHeight    =   4000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5000
   OleObjectBlob   =   "frmTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' frmTemplate
' Description : [Form description]
' Author      : [Name]
' Date        : [YYYY-MM-DD]
'-------------------------------------------------------------------------------

Option Explicit

'===============================================================================
' SECTION: Private variables
'===============================================================================

Private mData As Object
Private mValidators As Collection
Private mIsDirty As Boolean

'===============================================================================
' SECTION: Class events
'===============================================================================

Private Sub UserForm_Initialize()
    'Initialize form
    Set mData = CreateObject("Scripting.Dictionary")
    Set mValidators = New Collection
    mIsDirty = False

    'Setup UI
    SetupControls
    LoadData
End Sub

Private Sub UserForm_Terminate()
    Set mData = Nothing
    Set mValidators = Nothing
End Sub

' =============================================================================
' INITIALIZATION
' =============================================================================

Private Sub SetupControls()
    'Configure control positions and properties
End Sub

Private Sub LoadData()
    'Load initial data into controls
End Sub

' =============================================================================
' EVENT HANDLERS
' =============================================================================

Private Sub btnOK_Click()
    'Handle OK button
    If ValidateForm Then
        SaveData
        Me.Hide
    End If
End Sub

Private Sub btnCancel_Click()
    'Handle Cancel button
    Me.Hide
End Sub

Private Sub txtInput_Change()
    'Mark form as modified
    mIsDirty = True
End Sub

' =============================================================================
' VALIDATION
' =============================================================================

Private Function ValidateForm() As Boolean
    'Validate all form controls

    Dim isValid As Boolean
    isValid = True

    ' Example:
    ' If Not ValidateRequired(txtName, "Name") Then isValid = False

    ValidateForm = isValid
End Function

Private Function ValidateRequired(ByRef ctrl As Control, ByVal fieldName As String) As Boolean
    'Validate that a field is not empty

    If Trim(ctrl.Value) = vbNullString Then
        HighlightError ctrl
        MsgBox fieldName & " is required.", vbExclamation
        ValidateRequired = False
    Else
        ClearError ctrl
        ValidateRequired = True
    End If
End Function

Private Sub HighlightError(ByRef ctrl As Control)
    ctrl.BackColor = RGB(255, 200, 200)
End Sub

Private Sub ClearError(ByRef ctrl As Control)
    ctrl.BackColor = RGB(255, 255, 255)
End Sub

' =============================================================================
' DATA MANAGEMENT
' =============================================================================

Private Sub SaveData()
    'Save form data to dictionary

    mData("Field1") = txtField1.Value
    mData("Field2") = txtField2.Value
    ' Add more fields...

    mIsDirty = False
End Sub

Public Sub LoadFromData(ByRef data As Object)
    'Load form from existing data

    Set mData = data

    txtField1.Value = mData("Field1")
    txtField2.Value = mData("Field2")
    ' Add more fields...

    mIsDirty = False
End Sub

' =============================================================================
' PUBLIC INTERFACE
' =============================================================================

Public Function ShowDialog() As Object
    'Show form modally and return data

    Me.Show vbModal

    If mData.Count > 0 Then
        Set ShowDialog = mData
    Else
        Set ShowDialog = Nothing
    End If
End Function

Public Property Get FormData() As Object
    Set FormData = mData
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property
