Attribute VB_Name = "modToolsOther"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CloseAllWindowsVBE - closes all VBE windows except the active one
'* Created    : 01-20-2020 14:32
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CloseAllWindowsVBE()
    Dim vbWin       As VBIDE.Window
    For Each vbWin In Application.VBE.Windows
        If (vbWin.Type = vbext_wt_CodeWindow Or vbWin.Type = vbext_wt_Designer) And Not vbWin Is Application.VBE.ActiveWindow Then
            vbWin.Close
        End If
    Next vbWin
    Application.VBE.ActiveWindow.WindowState = vbext_ws_Maximize
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddLegendHotKeys - create message in Immediate window for Hot Keys
'* Created    : 22-03-2023 15:26
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddLegendHotKeys()
    Dim sPatpApp    As String
    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & "MacroToolsHotKeys.exe"
    If Not FileHave(sPatpApp, vbNormal) Then
        Debug.Print ">> This function is not available, file not found: " & sPatpApp
        Debug.Print ">> Download available at                        : " & "https://github.com/vbatools/MacroToolsVBAHotKeys" & vbNewLine
    End If
    Debug.Print addTabelFromArray(shSettings.ListObjects("TB_HOT_KEYS").Range.Value2, "|", True, 3)
End Sub

Public Sub showMsgBoxGenerator()
    Call frmBilderMsgBoxGenerator.Show
End Sub

Public Sub showBilderFormat()
    Call frmBilderFormat.Show
End Sub

Public Sub showBilderProcedure()
    Call frmBilderProcedure.Show
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ShowTODOList - call form with TODO list
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub showTODOList()
    Call frmTODO.Show
End Sub