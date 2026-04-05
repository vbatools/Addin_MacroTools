VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVariableUnUsed 
   Caption         =   "Search for Unused Variables:"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14415
   OleObjectBlob   =   "frmVariableUnUsed.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVariableUnUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim clsAnc          As clsAnchors

Private Sub cmbMain_Change()
      With cmbMain
          If VBA.Len(.value) = 0 Then Exit Sub
          Dim wb As Workbook
        Set wb = Workbooks(.value)
        If wb.vbProject.Protection = vbext_pp_locked Then
            lbLockedFile.Visible = True
            lbAnaliz.Enabled = False
        Else
             lbLockedFile.Visible = False
             lbAnaliz.Enabled = True
        End If
    End With
End Sub

Private Sub lbAnaliz_Click()
    lbOK.Caption = 0
    Me.Hide
End Sub

Private Sub lbLoad_Click()
    With ListCode
        If .ListCount > 0 Then
            Debug.Print addTabelFromArray(.List, "|", False, 9)
        Else
            Debug.Print ">> Analysis completed: no data to analyze."
        End If
    End With
End Sub

Private Sub lbLoadWB_Click()
    Dim sNameWB     As String
    sNameWB = cmbMain.value
    If VBA.Len(sNameWB) = 0 Then
        Call MsgBox("No workbook selected!", vbCritical)
        Exit Sub
    End If
    If ListCode.ListCount = 0 Then
        Call MsgBox("Nothing found or analysis not performed", vbCritical)
        Exit Sub
    End If

    Dim wb          As Workbook
    Set wb = Workbooks(sNameWB)
    Dim arr         As Variant
    arr = ListCode.List
    Call OutputResults(wb, "UnUsedVar", wb.FullName, arr)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With

    Set clsAnc = New clsAnchors
    With clsAnc
        Call .Initialize(Me, 600, 1000)
        Call .SetAnchorStyleByName(cmbMain.Name, anchorRight Or anchorTop Or anchorLeft)
        Call .SetAnchorStyleByName(ListCode.Name, anchorRight Or anchorTop Or anchorLeft Or anchorBottom)
        Call .SetAnchorStyleByName(lbAnaliz.Name, anchorTop Or anchorRight)
        Call .SetAnchorStyleByName(lbCancel.Name, anchorBottom Or anchorRight)
        Call .SetAnchorStyleByName(lbLoad.Name, anchorBottom Or anchorRight)
    End With
    lbLockedFile.Caption = VBA.ChrW$(60848)
End Sub

Private Sub UserForm_Activate()
    Dim vbProj      As VBIDE.vbProject

    If Workbooks.Count = 0 Then
        Me.Hide
        Call MsgBox("No open" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
        Unload Me
        Exit Sub
    End If
    With Me.cmbMain
        .Clear
        On Error Resume Next
        For Each vbProj In Application.VBE.VBProjects
            Call .AddItem(sGetFileName(vbProj.FileName))
        Next
        On Error GoTo 0
        .value = ActiveWorkbook.Name
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Set clsAnc = Nothing
End Sub
