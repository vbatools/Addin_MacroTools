VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfoFile 
   Caption         =   "File Properties:"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13410
   OleObjectBlob   =   "frmInfoFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInfoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmInfoFile - Manage file properties
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   20-07-2020 15:34
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private Sub cmbMain_Change()
    If cmbMain.value = vbNullString Then Exit Sub

    On Error Resume Next
    Dim arr         As Variant
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)
    arr = getFilePropertiesList(wb)
    If Not IsEmpty(arr) Then
        ListProp.List = arr
    Else
        ListProp.Clear
    End If

    arr = getFilePropertiesCustomList(wb)
    If Not IsEmpty(arr) Then
        ListCustomProp.List = arr
    Else
        ListCustomProp.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub LbDelAllProper_Click()
    If cmbMain.value = vbNullString Then
        Call MsgBox("No workbook selected!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)

    If MsgBox("Delete ALL properties?", vbYesNo + vbQuestion, "Deleting Properties:") = vbYes Then
        Dim iCount  As Byte
        iCount = delFilePropertiesAll(wb)
        Call cmbMain_Change
        Call MsgBox("Properties deleted:" & iCount, vbInformation, "Deleting Properties:")
    End If
End Sub
Private Sub LbEdit_Click()
    Call editProperty
End Sub

Private Sub lbLastAutor_Click()
    Me.Hide
    frmInfoFileLastAutor.Show
    Me.Show
End Sub

Private Sub lbTemplete_Click()
    If cmbMain.value = vbNullString Then
        Call MsgBox("No workbook selected!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)

    Dim tbData      As Variant
    Dim i           As Integer
    tbData = shSettings.ListObjects("TB_COMMENTS").DataBodyRange.Value2
    For i = 1 To UBound(tbData)
        If tbData(i, 2) <> vbNullString Then Call addFilePropertyCustom(wb, tbData(i, 1), tbData(i, 2))
    Next i
    Call cmbMain_Change
End Sub

Private Sub ListProp_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call editProperty
End Sub
Private Sub editProperty()
    If cmbMain.value = vbNullString Then
        Call MsgBox("No workbook selected!", vbCritical)
        Exit Sub
    End If
    On Error Resume Next
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)

    Dim i           As Long
    With ListProp
        i = .ListIndex
        If i < 0 Then Exit Sub
        Dim sNameProp As String
        Dim sValueProp As String
        sNameProp = .List(i, 1)
        sValueProp = VBA.Trim$(.List(i, 2))
    End With
    Dim sNewValueProp As String

    sNewValueProp = InputBox("Edit property [" & sNameProp & " ] ?", "Editing Property:", sValueProp)
    If sNewValueProp <> sValueProp Then
        If addFileProperty(wb, sNameProp, sNewValueProp) Then Call cmbMain_Change
    End If
End Sub

Private Sub lbAddCustProp_Click()
    Call AddCustProp(vbNullString, vbNullString)
End Sub

Private Sub lbEditCustProp_Click()

    Dim i           As Long
    With ListCustomProp
        i = .ListIndex
        If i < 0 Then Exit Sub
        Dim sNameProp As String
        Dim sValueProp As String
        sNameProp = .List(i, 1)
        sValueProp = VBA.Trim$(.List(i, 2))
    End With

    Call AddCustProp(sNameProp, sValueProp)
End Sub
Private Sub lbDelOneCustProp_Click()

    If cmbMain.value = vbNullString Then
        Call MsgBox("No workbook selected!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)

    Dim i           As Long
    With ListCustomProp
        i = .ListIndex
        If i < 0 Then Exit Sub
        Dim sNameProp As String
        sNameProp = .List(i, 1)
    End With
    If MsgBox("Delete property [" & sNameProp & " ] ?", vbYesNo + vbQuestion, "Deleting Property:") = vbYes Then
        Call delFilePropertyCustom(wb, sNameProp)
        Call cmbMain_Change
    End If
End Sub
Private Sub AddCustProp(ByVal txtPropName As String, ByVal txtPropValue As String)

    If cmbMain.value = vbNullString Then
        Call MsgBox("No workbook selected!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)

    txtPropName = InputBox("Enter property name", "Creating Property:", txtPropName)
    If txtPropName <> vbNullString Then
        txtPropValue = InputBox("Enter property value", "Creating Property:", txtPropValue)
        If txtPropValue <> vbNullString Then
            Call addFilePropertyCustom(wb, txtPropName, txtPropValue)
            Call cmbMain_Change
        End If
    End If
End Sub

Private Sub lbDelAllCustomProp_Click()
    If cmbMain.value = vbNullString Then
        Call MsgBox("No workbook selected!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)

    If MsgBox("Delete ALL properties?", vbYesNo + vbQuestion, "Deleting Properties:") = vbYes Then
        Dim iCount  As Byte
        iCount = delFilePropertiesCustomAll(wb)
        Call cmbMain_Change
        Call MsgBox("Properties deleted: " & iCount, vbInformation, "Deleting Properties:")
    End If
End Sub

Private Sub UserForm_Activate()
    Dim vbProj      As VBIDE.vbProject
    If Workbooks.Count = 0 Then
        Unload Me
        Call MsgBox("No open" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
        Exit Sub
    End If
    With Me.cmbMain
        .Clear
        On Error Resume Next
        For Each vbProj In Application.VBE.VBProjects
            Call .AddItem(sGetFileName(vbProj.FileName))
        Next
        .value = ActiveWorkbook.Name
        On Error GoTo 0
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Call btnCancel_Click
End Sub
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub