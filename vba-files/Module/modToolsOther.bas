Attribute VB_Name = "modToolsOther"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1



'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CloseAllWindowsVBE - закрывает все окна VBE, кроме активного
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
