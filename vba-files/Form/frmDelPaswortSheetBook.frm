VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDelPaswortSheetBook 
   Caption         =   "Password Removal:"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13560
   OleObjectBlob   =   "frmDelPaswortSheetBook.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDelPaswortSheetBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub lbCancel_Click()
    Call btnCancel_Click
End Sub

Private Sub lbOK_Click()
    lbValue.Caption = -1
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With

    Dim wb          As Workbook
    For Each wb In Workbooks
        cmbMain.AddItem wb.Name
    Next
    cmbMain.value = ActiveWorkbook.Name
    Call loadDataToList
End Sub

Private Sub cmbMain_Change()
    Call loadDataToList
End Sub

Private Sub loadDataToList()
    Dim wb          As Workbook
    Dim sh          As Worksheet
    Dim i           As Long
    Dim bHaveShProtect As Boolean
    Set wb = Workbooks(cmbMain.value)
    lbMsg.Visible = wb.ProtectStructure

    With ListMain
        .Clear
        For Each sh In wb.Worksheets
            .AddItem sh.Name
            .List(i, 1) = "no"
            If sh.ProtectContents Then
                .List(i, 1) = "yes"
                bHaveShProtect = True
            End If

            i = i + 1
        Next sh
    End With
    lbOk.Enabled = (bHaveShProtect Or lbMsg.Visible)
End Sub