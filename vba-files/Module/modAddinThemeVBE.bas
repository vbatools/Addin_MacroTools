Attribute VB_Name = "modAddinThemeVBE"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : V_BlackAndWiteTheme - switch VBE editor theme between light and dark
'* Created    : 19-02-2020 12:57
'* Created    : 22-03-2023 14:33
'* Author     : VBATools
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Const REG   As String = "HKEY_CURRENT_USER\Software\Microsoft\VBA\"
Private Const REG_BACK_COLOR As String = "\Common\CodeBackColors"
Private Const REG_FORE_COLOR As String = "\Common\CodeForeColors"
Private Const BACK_COLOR_BLACK_THEME As String = "4 0 4 7 6 4 4 4 11 4 0 0 0 0 0 0"
Private Const FORE_COLOR_BLACK_THEME As String = "1 0 5 14 1 9 11 15 4 1 0 0 0 0 0 0"
Private Const BACK_COLOR_WHITE_THEME As String = "0 0 0 7 6 0 0 0 0 0 0 0 0 0 0 0"
Private Const FORE_COLOR_WHITE_THEME As String = "0 0 5 0 1 10 14 0 0 0 0 0 0 0 0 0"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ChangeColorWhiteTheme - switch VBE editor theme to light
'* Created    : 23-03-2023 10:01
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub changeColorWhiteTheme()
    Call changeColorTheme(BACK_COLOR_WHITE_THEME, FORE_COLOR_WHITE_THEME, "Light theme enabled, please restart MS Excel")
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ChangeColorDarkTheme - switch VBE editor theme to dark
'* Created    : 23-03-2023 10:02
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub changeColorDarkTheme()
    Call changeColorTheme(BACK_COLOR_BLACK_THEME, FORE_COLOR_BLACK_THEME, "Dark theme enabled, please restart MS Excel")
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ChangeColorTheme - Main procedure for switching themes
'* Created    : 19-02-2020 19:12
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                     Description
'*
'* ByVal sBackColorTheme As String : theme background color
'* ByVal sForeColorTheme As String : theme foreground color (style)
'* sMsg As String                  : message string
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub changeColorTheme(ByVal sBackColorTheme As String, ByVal sForeColorTheme As String, sMsg As String)
    Dim BackColor   As String
    Dim ForeColor   As String

    On Error GoTo ErrorHandler

    BackColor = REG & GetVersionVBE & REG_BACK_COLOR
    ForeColor = REG & GetVersionVBE & REG_FORE_COLOR

    With CreateObject("WScript.Shell")
        .RegWrite BackColor, sBackColorTheme, "REG_SZ"
        .RegWrite ForeColor, sForeColorTheme, "REG_SZ"
    End With
    Call MsgBox(sMsg, vbInformation, "Theme Change:")

    Exit Sub
ErrorHandler:
    Call WriteErrorLog("changeColorTheme", True)
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetVersionVBE - returns the VBA version used in the system
'* Created    : 19-02-2020 19:12
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetVersionVBE() As String
    Dim sVersion    As String

    sVersion = VBA.Replace(Application.VBE.Version, 0, vbNullString)
    If VBA.Right$(sVersion, 1) = "." Then
        sVersion = sVersion & "0"
    End If
    GetVersionVBE = sVersion
End Function