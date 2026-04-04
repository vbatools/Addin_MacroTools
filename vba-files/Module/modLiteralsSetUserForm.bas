Attribute VB_Name = "modLiteralsSetUserForm"
Option Explicit
Option Private Module

' Enumeration for indexing array columns arrUserForm
Private Enum UserFormCols
    moduleName = 1
    ControlType = 2
    ParentName = 3
    ControlName = 4
    PropertyName = 5
    NewValue = 7
    Status = 8
End Enum

' Constants for operation statuses
Private Const STATUS_CHANGED As String = "modified"
Private Const STATUS_PROP_NOT_FOUND As String = "property not found"
Private Const STATUS_CTRL_NOT_FOUND As String = "control not found"
Private Const STATUS_PARENT_NOT_FOUND As String = "parent control not found"
Private Const STATUS_MODULE_NOT_FOUND As String = "module not found"
Private Const STATUS_NOT_A_FORM As String = "module is not a form"

Public Function renameLiteralsToUserForm(ByRef vbProj As VBIDE.vbProject, ByRef arrUserForm As Variant) As Boolean
    On Error GoTo ErrorHandler

    If IsEmpty(arrUserForm) Then Exit Function

    Dim oVBModule    As VBIDE.vbComponent
    Dim sCurrentModuleName As String
    Dim iRowCount   As Long
    Dim i           As Long
    Dim ctrl        As control

    iRowCount = UBound(arrUserForm, 1)
    sCurrentModuleName = vbNullString

    For i = 1 To iRowCount
        ' Skip empty rows
        If VBA.Len(arrUserForm(i, UserFormCols.moduleName)) = 0 Then GoTo NextIteration

        ' Optimization: get the module only when the name changes
        If sCurrentModuleName <> arrUserForm(i, UserFormCols.moduleName) Then
            sCurrentModuleName = arrUserForm(i, UserFormCols.moduleName)
            Set oVBModule = getVBModuleByName(vbProj, sCurrentModuleName)
            oVBModule.Activate
        End If

        ' Check the module
        If oVBModule Is Nothing Then
            arrUserForm(i, UserFormCols.Status) = STATUS_MODULE_NOT_FOUND
            GoTo NextIteration
        End If

        ' Check the type (must be a form)
        If oVBModule.Type <> vbext_ct_MSForm Then
            arrUserForm(i, UserFormCols.Status) = STATUS_NOT_A_FORM
            GoTo NextIteration
        End If

        ' Processing logic with status return via function
        If IsFormSelfReference(arrUserForm, i) Then
            ' Modify the form's own properties
            arrUserForm(i, UserFormCols.Status) = UpdateFormProperty(oVBModule, arrUserForm(i, UserFormCols.NewValue))
        ElseIf IsDirectControl(arrUserForm, i) Then
            ' Modify a control on the form
            Set ctrl = getControl(oVBModule, arrUserForm(i, UserFormCols.ControlName))
            arrUserForm(i, UserFormCols.Status) = UpdateObjectProperty(ctrl, arrUserForm(i, UserFormCols.PropertyName), arrUserForm(i, UserFormCols.NewValue), STATUS_CTRL_NOT_FOUND)
        Else
            ' Modify a nested element (Tab/Page)
            arrUserForm(i, UserFormCols.Status) = UpdateNestedControl(oVBModule, arrUserForm, i)
        End If

NextIteration:
    Next i

    renameLiteralsToUserForm = True
    Exit Function

ErrorHandler:
    Debug.Print ">> Error in renameLiteralsToUserForm: " & Err.Description
    renameLiteralsToUserForm = False
End Function

Private Function IsFormSelfReference(ByRef arr As Variant, ByVal i As Long) As Boolean
    IsFormSelfReference = (arr(i, UserFormCols.moduleName) = arr(i, UserFormCols.ParentName)) And _
            (arr(i, UserFormCols.ParentName) = arr(i, UserFormCols.ControlName))
End Function

Private Function IsDirectControl(ByRef arr As Variant, ByVal i As Long) As Boolean
    IsDirectControl = (arr(i, UserFormCols.moduleName) = arr(i, UserFormCols.ParentName)) And _
            (arr(i, UserFormCols.ParentName) <> arr(i, UserFormCols.ControlName))
End Function

' Returns the operation status for the form itself
Private Function UpdateFormProperty(ByRef VBModule As VBIDE.vbComponent, ByVal sValue As String) As String
    On Error Resume Next
    VBModule.Properties("Caption").value = sValue
    If Err.Number = 0 Then
        UpdateFormProperty = STATUS_CHANGED
    Else
        UpdateFormProperty = STATUS_PROP_NOT_FOUND
    End If
    On Error GoTo 0
End Function

' Universal function for updating a property of any object
' Returns a status string
Private Function UpdateObjectProperty(ByRef obj As Object, ByVal sProp As String, ByVal sValue As String, ByVal sNotFoundMsg As String) As String
    If obj Is Nothing Then
        UpdateObjectProperty = sNotFoundMsg
        Exit Function
    End If

    If setValueInControl(obj, sProp, sValue) Then
        UpdateObjectProperty = STATUS_CHANGED
    Else
        UpdateObjectProperty = STATUS_PROP_NOT_FOUND
    End If
End Function

' Processing nested controls
Private Function UpdateNestedControl(ByRef VBModule As VBIDE.vbComponent, ByRef arr As Variant, ByVal i As Long) As String
    Dim parentCtrl  As control
    Dim nestedObj   As Object
    Dim sParentName As String
    Dim sChildName  As String
    Dim sCtrlType   As String

    sParentName = arr(i, UserFormCols.ParentName)
    sChildName = arr(i, UserFormCols.ControlName)
    sCtrlType = arr(i, UserFormCols.ControlType)

    Set parentCtrl = getControl(VBModule, sParentName)

    If parentCtrl Is Nothing Then
        UpdateNestedControl = STATUS_PARENT_NOT_FOUND
        Exit Function
    End If

    ' Search for nested object
    On Error Resume Next
    Select Case sCtrlType
             Case "Tab"
            Set nestedObj = parentCtrl.Tabs(VBA.CStr(sChildName))
        Case "Page"
            Set nestedObj = parentCtrl.Pages(VBA.CStr(sChildName))
    End Select
    On Error GoTo 0
    ' Delegate property update to the universal function
    UpdateNestedControl = UpdateObjectProperty(nestedObj, arr(i, UserFormCols.PropertyName), arr(i, UserFormCols.NewValue), STATUS_CTRL_NOT_FOUND)
End Function

Private Function getControl(ByRef VBModule As VBIDE.vbComponent, ByVal sNameControl As String) As control
    On Error Resume Next
    Set getControl = VBModule.Designer.Controls(sNameControl)
    On Error GoTo 0
End Function

Private Function setValueInControl(ByRef cnt As Object, ByVal sNameProperty As String, ByVal sValue As String) As Boolean
    On Error Resume Next
    Call CallByName(cnt, sNameProperty, VbLet, sValue)
    setValueInControl = (Err.Number = 0)
    On Error GoTo 0
End Function