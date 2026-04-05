VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBECommandHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : VBECommandHandler - callback routines for VBE menu
'* Created    : 01-20-2020 15:55
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public WithEvents cmdButton  As CommandBarButton
Attribute cmdButton.VB_VarHelpID = -1

Private Sub cmdButton_Click(ByVal ctrl As CommandBarButton, CancelDefault As Boolean)
    On Error GoTo ErrorHandler
    Call Application.Run(ctrl.OnAction)
    Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then
        Call WriteErrorLog("clsVBECommandHandler.cmdButton_Click", False)
        Err.Clear
    End If
End Sub
