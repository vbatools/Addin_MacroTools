VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListWBOpen 
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11625
   OleObjectBlob   =   "frmListWBOpen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmListWBOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : AddStatistic - Form for file selection
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub cmbCancel_Click()
    cmbMain.Clear
    cmbMain.value = vbNullString
    Unload Me
End Sub

Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub
Private Sub lbOK_Click()
    lbRes.Caption = -1
    Unload Me
End Sub
Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

    Dim vbProj      As VBIDE.vbProject
    On Error Resume Next
    With cmbMain
        .Clear
        For Each vbProj In Application.VBE.VBProjects
            .AddItem sGetFileName(vbProj.FileName)
        Next
        If lbWord.Caption = "1" Then Call getWord(cmbMain)
        .value = ActiveWorkbook.Name
    End With
    Exit Sub
End Sub

Private Sub getWord(ByRef oList As MSForms.ComboBox)
    On Error Resume Next
    Dim objW        As Object
    Dim vbProj      As VBIDE.vbProject
    Dim sVal        As String
    Set objW = GetObject(, "Word.Application")
    For Each vbProj In objW.VBE.VBProjects
        sVal = sGetFileName(vbProj.FileName)
        If sVal Like "*.docm" Or sVal Like "*.DOCM" Then oList.AddItem sVal
    Next
End Sub
