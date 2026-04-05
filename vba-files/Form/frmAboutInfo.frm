VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAboutInfo 
   Caption         =   "About Add-in:"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8520.001
   OleObjectBlob   =   "frmAboutInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAboutInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub lbCancel_Click()
    Unload Me
End Sub

Private Sub lbGoGitHub_Click()
    Call URLLinks("https://github.com/vbatools/Addin_MacroTools")
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    lbAbout.Caption = Version(enAll)
End Sub
