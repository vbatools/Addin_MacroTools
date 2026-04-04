VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHideModule 
   Caption         =   "Hide VBA Modules:"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415.001
   OleObjectBlob   =   "frmHideModule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHideModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : HiddenModule - Hide VBA modules
'* Created    : 12-02-2020 10:19
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub cmbCancel_Click()
    Unload Me
End Sub

Private Sub cmbMain_Change()
    Call AddListCode
End Sub

Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub
Private Sub CheckAll_Click()
    Dim i           As Integer
    With ListCode
        For i = 0 To .ListCount - 1
            .Selected(i) = CheckAll.value
        Next i
    End With
End Sub

Private Sub ListCode_Change()
    Dim i           As Integer
    With ListCode
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                lbMsg.Visible = False
                lbOk.Enabled = True
                Call MsgSaveFile(cmbMain.value)
                Exit Sub
            End If
        Next i
    End With

    lbMsg.Visible = True
    lbOk.Enabled = False
End Sub
Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

    Dim wb          As Workbook
    On Error GoTo ErrorHandler
    If Workbooks.Count = 0 Then
        Unload Me
        Call MsgBox("No open Excel files" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
        Exit Sub
    End If
    With Me.cmbMain
        .Clear
        For Each wb In Workbooks
            .AddItem wb.Name
        Next
        .value = ActiveWorkbook.Name
        Call MsgSaveFile(.value)
    End With
    Call AddListCode
    lbOk.Enabled = False

    Exit Sub
ErrorHandler:
    Unload Me
    Select Case Err.Number
        Case Else:
            Call WriteErrorLog(Me.Name & ".UserForm_Activate", True)
    End Select
    Err.Clear
End Sub

Private Sub MsgSaveFile(ByVal WBName As String)
    Dim wb          As Workbook
    Set wb = Workbooks(WBName)
    With wb
        If .Path = vbNullString Then
            lbSave.Visible = True
            lbOk.Enabled = False
        Else
            lbSave.Visible = False
            lbOk.Enabled = True
        End If
    End With
End Sub

Private Sub AddListCode()
    If cmbMain.value = vbNullString Then Exit Sub
    Dim wb          As Workbook
    Dim i           As Long
    Set wb = Workbooks(cmbMain.value)
    With wb.vbProject
        ListCode.Clear
        If .Protection <> vbext_pp_none Then
            Call MsgBox("VBA project in workbook - " & cmbMain.value & " is password protected!" & vbCrLf & "Remove the password!", vbCritical, "Error:")
            Exit Sub
        End If
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = vbext_ct_StdModule Then
                ListCode.AddItem .VBComponents(i).Name
            End If
        Next i

        If ListCode.ListCount < 1 Then Exit Sub

        Dim clsSort As clsSort2DArray
        Set clsSort = New clsSort2DArray
        With clsSort
            Call .SortQuick(ListCode.List, 0, True, sdtString)
            ListCode.List = .ListArray
        End With
        Set clsSort = Nothing
    End With
End Sub

Private Sub lbOK_Click()
    If cmbMain.value = vbNullString Then Exit Sub

    Dim i           As Long
    Dim k           As Long
    Dim iCount      As Long
    With ListCode
        iCount = .ListCount
        If iCount < 0 Then Exit Sub
        ReDim arr(1 To iCount, 1 To 1) As String
        For i = 0 To iCount - 1
            If .Selected(i) Then
                k = k + 1
                arr(k, 1) = .List(i, 0)
            End If
        Next i
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)
    Dim sNameFileName As String
    With wb
        .Save
        sNameFileName = .Path & Application.PathSeparator & sGetBaseName(.Name) & "_hidden" & "." & sGetExtensionName(.Name)
        Call .SaveAs(sNameFileName)
        If chbAddModule.value Then
            For i = 1 To 3
                Call AddModuleToProject(.vbProject, "Module" & i, vbext_ct_StdModule, vbNullString, False)
            Next i
            .Save
        End If
    End With

    Call wb.Close
    Call hideModules(sNameFileName, arr)
    Call Workbooks.Open(sNameFileName)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Unload Me
    Call MsgBox("VBA modules hidden!", vbInformation, "Hide VBA Modules:")
End Sub