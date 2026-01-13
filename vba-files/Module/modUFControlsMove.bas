Attribute VB_Name = "modUFControlsMove"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MoveControl - Moves controls in the designer
'* Created    : 08-10-2020 14:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MoveControlRight()
    Call MoveControls(1)
End Sub

Private Sub MoveControlLeft()
    Call MoveControls(2)
End Sub

Private Sub MoveControlDown()
    Call MoveControls(3)
End Sub

Private Sub MoveControlUp()
    Call MoveControls(4)
End Sub

Private Sub MoveControls(ByVal id As Byte)
    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub

    Dim myCommandBar As CommandBar
    Dim combox      As CommandBarComboBox
    Dim sComBoxText As String
    Dim cnt         As Control

    Set myCommandBar = Application.VBE.CommandBars(modAddinConst.MENU_MOVE_CONTROLS)
    Set combox = myCommandBar.Controls(1)
    sComBoxText = combox.Text

    Dim objActiveModule As VBComponent
    Set objActiveModule = getActiveModule()
    For Each cnt In objActiveModule.Designer.Selected
        If Not cnt Is Nothing Then
            Call MoveControl(cnt, id, sComBoxText)
        End If
    Next cnt
End Sub

'* * * * *
'* Sub        : MoveControl - Moves controls on UserForms
'* Created    : 22-03-2023 15:58
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByRef cnt As control        : Control
'* ByVal iVal As Integer       : Direction value
'* ByVal sComBoxText As String : Direction text
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MoveControl(ByRef cnt As Control, ByVal id As Byte, ByVal sComBoxText As String)
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
                        .top = .top + Shag
                    Case 4:
                        .top = .top - Shag
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
                        .top = .top + Shag
                        .Height = .Height - Shag
                    Case 4:
                        .top = .top - Shag
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
