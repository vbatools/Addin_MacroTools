VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTODO 
   Caption         =   "VBA Project Manager:"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13575
   OleObjectBlob   =   "frmTODO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTODO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ModuleTODO - Module for searching TODO tags
'* Created    : 01-20-2020 12:34
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private clsAnc      As clsAnchors

Private Sub lbLoad_Click()
    With ListCode
        If .ListCount = 0 Then Exit Sub
        Dim arr     As Variant
        arr = .List
        Debug.Print addTabelFromArray(.List, "|", False, .ColumnCount)
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    Set clsAnc = New clsAnchors
    Call clsAnc.Initialize(Me, 600, 1000)

    With clsAnc
        Call .SetAnchorStyleByName(cmbMain.Name, anchorRight Or anchorTop Or anchorLeft)
        Call .SetAnchorStyleByName(ListCode.Name, anchorRight Or anchorTop Or anchorLeft Or anchorBottom)
        Call .SetAnchorStyleByName(lbCancel.Name, anchorBottom Or anchorRight)
        Call .SetAnchorStyleByName(lbLoad.Name, anchorBottom Or anchorRight)
    End With
End Sub
Private Sub UserForm_Activate()
    Dim vbProj      As VBIDE.vbProject
    If Workbooks.Count = 0 Then
        Unload Me
        Call MsgBox("No open Excel files" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
        Exit Sub
    End If
    With Me.cmbMain
        .Clear
        On Error Resume Next
        For Each vbProj In Application.VBE.VBProjects
            .AddItem sGetFileName(vbProj.FileName)
        Next
        .value = ActiveWorkbook.Name
        On Error GoTo 0
        Call AddTODOList(.value)
    End With
End Sub
Private Sub UserForm_Terminate()
    Set clsAnc = Nothing
End Sub

Private Sub cmbCancel_Click()
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub
Private Sub cmbMain_Change()
    If cmbMain.value = vbNullString Then Exit Sub
    Call AddTODOList(cmbMain.value)
End Sub
Private Sub ListCode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i           As Long
    Dim wb          As Workbook
    Dim VBComp      As VBIDE.vbComponent
    Dim iLine       As Long

    On Error GoTo ErrorHandler

    If cmbMain.value = vbNullString Then Exit Sub
    Set wb = Workbooks(cmbMain.value)
    For i = 0 To ListCode.ListCount
        If ListCode.Selected(i) = True Then
            Set VBComp = wb.vbProject.VBComponents(ListCode.List(i, 2))
            With VBComp
                If .Type = vbext_ct_MSForm Then
                    .codeModule.CodePane.Show
                Else
                    .Activate
                End If
                iLine = VBA.CLng(ListCode.List(i, 3))
                Call .codeModule.CodePane.SetSelection(iLine + 1, 1, iLine + 1, 1)
            End With
            Exit Sub
        End If
    Next i
    Exit Sub
ErrorHandler:
    Unload Me
    Select Case Err.Number
        Case Else:
            Call WriteErrorLog(Me.Name & ".ListCode_DblClick", True)
    End Select
    Err.Clear
End Sub

Private Sub AddTODOList(sWb As String)
    If VBA.Len(sWb) = 0 Then Exit Sub
    Dim iFile       As Integer
    Dim wb          As Workbook
    On Error GoTo ErrorHandler
    Set wb = Workbooks(sWb)
    If wb.vbProject.Protection = vbext_pp_none Then
        ListCode.Clear
        For iFile = 1 To wb.vbProject.VBComponents.Count
            Call listLinesinModuleWhereFound(wb.vbProject.VBComponents(iFile), sCREATED)
        Next iFile
    Else
        ListCode.Clear
        Call MsgBox("VBA project in workbook - " & wb.Name & "is password protected!" & vbCrLf & "Remove the password!", vbCritical, "Error:")
    End If
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 4160:
            ListCode.Clear
            Call MsgBox("Error! No access to VBA project!", vbOKOnly + vbExclamation, "Error:")
        Case Else:
            Call WriteErrorLog(Me.Name & ".AddTODOList", True)
    End Select
    Err.Clear
End Sub

Sub listLinesinModuleWhereFound(ByVal oComponent As Object, ByVal sSearchTerm As String)
    Dim lTotalNoLines As Long
    Dim lLineNo     As Long
    Dim lListRow    As Long
    Dim arr         As Variant

    On Error GoTo ErrorHandler

    lLineNo = 1
    lListRow = ListCode.ListCount
    With oComponent
        lTotalNoLines = .codeModule.CountOfLines
        Do While .codeModule.Find(sSearchTerm, lLineNo, 1, -1, -1, False, False, False) = True
            ListCode.AddItem lListRow + 1
            ListCode.List(lListRow, 1) = moduleTypeName(.Type)
            ListCode.List(lListRow, 2) = .Name
            ListCode.List(lListRow, 3) = lLineNo
            arr = VBA.Split(VBA.Trim$(.codeModule.Lines(lLineNo, 1)), ": ")
            If UBound(arr) >= 2 Then
                ListCode.List(lListRow, 4) = VBA.Replace(arr(1), " Author", vbNullString)
                ListCode.List(lListRow, 5) = arr(2)
            Else
                ListCode.List(lListRow, 4) = vbNullString
                ListCode.List(lListRow, 5) = vbNullString
            End If
            ListCode.List(lListRow, 6) = VBA.Replace(VBA.Trim$(.codeModule.Lines(lLineNo + 1, 1)), "'*", vbNullString)
            lLineNo = lLineNo + 1
            lListRow = lListRow + 1
        Loop
    End With
    Exit Sub
ErrorHandler:
    Unload Me
    Select Case Err.Number
        Case Else:
            Call WriteErrorLog(Me.Name & ".listLinesinModuleWhereFound", True)
    End Select
    Err.Clear
End Sub