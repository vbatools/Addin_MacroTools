VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettingsKomments 
   Caption         =   "Comment Settings:"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10710
   OleObjectBlob   =   "frmSettingsKomments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettingsKomments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bChange         As Boolean

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub lbCancel_Click()
    Call btnCancel_Click
End Sub

Private Sub lbSave_Click()
    If MsgBox("Save changes?", vbYesNo + vbQuestion, "Saving Data:") = vbNo Then Exit Sub
    With shSettings.ListObjects("TB_COMMENTS")
        If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
        If listComments.ListCount > 0 Then
            Dim arr As Variant
            arr = listComments.List
            .Range(1, 1).Offset(1, 0).Resize(UBound(arr, 1) + 1, 2).Value2 = arr
        End If
    End With
End Sub

Private Sub listComments_Change()
    If bChange Then Exit Sub
    bChange = True
    Dim iRow        As Integer
    With listComments
        iRow = .ListIndex
        If iRow < 0 Then Exit Sub
        txtRow.value = iRow
        lbKey.Caption = .List(iRow, 0)
        txtValue.value = .List(iRow, 1)
    End With
    bChange = False
End Sub

Private Sub lbKey_Change()
    Call changeRow
End Sub

Private Sub txtValue_Change()
    Call changeRow
End Sub

Private Sub changeRow()
    If bChange Then Exit Sub
    bChange = True
    Dim iRow        As Integer
    iRow = VBA.CInt(txtRow.value)
    If iRow < 0 Then Exit Sub
    With listComments
        .List(iRow, 0) = lbKey.Caption
        .List(iRow, 1) = txtValue.value
    End With
    bChange = False
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    With shSettings.ListObjects("TB_COMMENTS")
        If Not .DataBodyRange Is Nothing Then listComments.List = .DataBodyRange.Value2
    End With
End Sub