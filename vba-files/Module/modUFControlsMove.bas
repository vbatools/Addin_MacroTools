Attribute VB_Name = "modUFControlsMove"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MoveControl - fine-tuning of form elements
'* Created    : 08-10-2020 14:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub MoveControlRight()
    Call MoveControls(1)
End Sub

Public Sub MoveControlLeft()
    Call MoveControls(2)
End Sub

Public Sub MoveControlDown()
    Call MoveControls(3)
End Sub

Public Sub MoveControlUp()
    Call MoveControls(4)
End Sub

Private Sub MoveControls(ByVal id As Byte)
    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub
    On Error GoTo ErrorHandler
    Dim sComBoxText As String
    sComBoxText = Application.VBE.CommandBars(modAddinConst.MENU_MOVE_CONTROLS).Controls(1).text

    Dim selectedControls As Object

    Set selectedControls = GetSelectControl()
    If selectedControls Is Nothing Then Exit Sub
    Select Case TypeName(selectedControls)
             Case "Controls"
            Dim cnt As control
            For Each cnt In selectedControls
                Call MoveControl(cnt, id, sComBoxText)
            Next cnt
        Case Else
            Call MoveControl(selectedControls, id, sComBoxText)
    End Select
    Set selectedControls = Nothing
    Exit Sub
ErrorHandler:
    WriteErrorLog "MoveControls", False
    On Error GoTo 0
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MoveControl - change control position on UserForms
'* Created    : 22-03-2023 15:58
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByRef cnt As control        : control
'* ByVal iVal As Integer       : move value
'* ByVal sComBoxText As String : move type
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MoveControl(ByRef cnt As control, ByVal id As Byte, ByVal sComBoxText As String)
    Const Shag = 0.4
    With cnt
        Select Case sComBoxText
            Case "Control":
                Select Case id
                    Case 1:
                        .Left = .Left - Shag
                    Case 2:
                        .Left = .Left + Shag
                    Case 3:
                        .Top = .Top + Shag
                    Case 4:
                        .Top = .Top - Shag
                End Select
            Case "Top Left":
                Select Case id
                    Case 1:
                        .Left = .Left - Shag
                        .Width = .Width + Shag
                    Case 2:
                        .Left = .Left + Shag
                        .Width = .Width - Shag
                    Case 3:
                        .Top = .Top + Shag
                        .Height = .Height - Shag
                    Case 4:
                        .Top = .Top - Shag
                        .Height = .Height + Shag
                End Select
            Case "Bottom Right":
                Select Case id
                    Case 1:
                        .Width = .Width - Shag
                    Case 2:
                        .Width = .Width + Shag
                    Case 3:
                        .Height = .Height + Shag
                    Case 4:
                        .Height = .Height - Shag
                End Select
        End Select
    End With
End Sub