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
'* Module     : VBECommandHandler - Handles events for buttons in VBE
'* Created    : 01-20-2020 15:55
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public WithEvents EvtHandler As CommandBarButton
Attribute EvtHandler.VB_VarHelpID = -1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : EvtHandler_Click - Event handler for button clicks in the VBE environment
'* Created    : 23-03-2023 10:28
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                     Description
'*
'* ByVal Ctrl As CommandBarButton :
'* CancelDefault As Boolean       :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub EvtHandler_Click(ByVal Ctrl As CommandBarButton, CancelDefault As Boolean)
    On Error GoTo ErrorHandler
    Call Application.Run(Ctrl.OnAction)
    Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error in EvtHandler_Click" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
        'Call WriteErrorLog("EvtHandler_Click")
        Err.Clear
    End If
End Sub
