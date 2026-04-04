Attribute VB_Name = "modLiteralsGetUserForm"
Option Explicit
Option Private Module

Private Enum DataFields
    dfModuleName = 1
    dfType
    dfParentName
    dfItemName
    dfPropertyName
    dfValue
End Enum

Public Function parserLiteralsFormControls(ByRef wb As Workbook) As Object
    Dim vbProj      As VBIDE.vbProject
    Dim VBCom       As VBIDE.vbComponent
    Dim objCont     As MSForms.control
    Dim oDic        As Object
    Dim sNameModule As String
    Dim sPropCaption As String

    Set oDic = CreateObject("Scripting.Dictionary")
    Set vbProj = wb.vbProject
    For Each VBCom In vbProj.VBComponents
        With VBCom
            If .Type = vbext_ct_MSForm Then
                sNameModule = .Name
                .Activate
                sPropCaption = .Properties("Caption").value
                Call AddItemToDictionary(oDic, sNameModule, sNameModule, sNameModule, "Form", "Caption", sPropCaption)
                For Each objCont In .Designer.Controls
                    Call ProcessControl(objCont, sNameModule, oDic)
                Next objCont
            End If
        End With
    Next VBCom
    Set parserLiteralsFormControls = oDic
End Function

Private Sub ProcessControl(ByRef objCont As MSForms.control, ByVal sModuleName As String, ByRef oDic As Object)
    Dim sContName   As String
    Dim sContType   As String
    Dim sValue      As String
    Dim sPropertyName As String
    Dim objSubItem  As Object
    Dim sItemType   As String

    sContName = objCont.Name
    sContType = TypeName(objCont)
    Select Case sContType
             Case "CommandButton", "Label", "OptionButton", "CheckBox", "ToggleButton", "Frame"
            sValue = GetPropertySafe(objCont, "Caption")
            sPropertyName = "Caption"
        Case "TextBox", "ComboBox"
            sValue = GetPropertySafe(objCont, "Value")
            sPropertyName = "Value"
        Case "TabStrip"
            For Each objSubItem In objCont.Tabs
                sItemType = "Tab"
                Call AddItemToDictionary(oDic, sModuleName, sContName, objSubItem.Name, sItemType, "Caption", _
                        GetPropertySafe(objSubItem, "Caption"))
                Call AddItemToDictionary(oDic, sModuleName, sContName, objSubItem.Name, sItemType, "ControlTipText", _
                        GetPropertySafe(objSubItem, "ControlTipText"))
            Next objSubItem
        Case "MultiPage"
            For Each objSubItem In objCont.Pages
                sItemType = "Page"
                Call AddItemToDictionary(oDic, sModuleName, sContName, objSubItem.Name, sItemType, "Caption", _
                        GetPropertySafe(objSubItem, "Caption"))
                Call AddItemToDictionary(oDic, sModuleName, sContName, objSubItem.Name, sItemType, "ControlTipText", _
                        GetPropertySafe(objSubItem, "ControlTipText"))
            Next objSubItem
    End Select
    If VBA.Len(sPropertyName) > 0 Then Call AddItemToDictionary(oDic, sModuleName, sModuleName, sContName, sContType, sPropertyName, sValue)
    Call AddItemToDictionary(oDic, sModuleName, sModuleName, sContName, sContType, "ControlTipText", GetPropertySafe(objCont, "ControlTipText"))
End Sub

Private Function GetPropertySafe(ByVal obj As Object, ByVal sPropName As String) As String
    On Error Resume Next
    GetPropertySafe = CallByName(obj, sPropName, VbGet)
    If Err.Number <> 0 Then GetPropertySafe = vbNullString
    On Error GoTo 0
End Function

Private Sub AddItemToDictionary(ByRef oDic As Object, _
        ByVal sModule As String, _
        ByVal sParent As String, _
        ByVal sName As String, _
        ByVal sType As String, _
        ByVal sPropertyName As String, _
        ByVal sValue As String)

    Dim arrData(1 To 1, 1 To 6) As String
    Dim sKey        As String

    arrData(1, DataFields.dfModuleName) = sModule
    arrData(1, DataFields.dfParentName) = sParent
    arrData(1, DataFields.dfItemName) = sName
    arrData(1, DataFields.dfType) = sType
    arrData(1, DataFields.dfPropertyName) = sPropertyName
    arrData(1, DataFields.dfValue) = sValue

    ' Form a unique key
    sKey = sModule & "." & sParent & "." & sName & "." & sType & "." & sPropertyName
    If Not oDic.Exists(sKey) Then oDic.Add sKey, arrData
End Sub