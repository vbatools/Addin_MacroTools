Attribute VB_Name = "modUFControlsStyleCopyPaste"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1
Public tpStyle      As ProperControlStyle
Type ProperControlStyle
    sError          As String
    snHeight        As Single
    snWidth         As Single
    bVisible        As Boolean
    bEnabled        As Boolean
    bLocked         As Boolean
    lBackColor      As Long
    lForeColor      As Long
    lBackStyle      As Long
    lBorderColor    As Long
    lBorderStyle    As Long
    bFontBold       As Boolean
    bFontItalic     As Boolean
    bFontStrikethru As Boolean
    bFontUnderline  As Boolean
    sFontName       As String
    cuFontSize      As Currency
End Type

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CopyStyleControl - Copies the style properties of a control
'* Created    : 22-03-2023 16:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CopyStyleControl()
    Dim cnt         As Object
    Set cnt = GetSelectControl(True)
    If cnt Is Nothing Then Exit Sub

    'Initialize default control properties
    tpStyle.lBackStyle = 1
    tpStyle.lBorderColor = -2147483642
    tpStyle.lBorderStyle = 0
    tpStyle.bVisible = True
    tpStyle.bLocked = False
    tpStyle.bEnabled = True
    tpStyle.lBackStyle = 1

    On Error Resume Next
    With cnt
        tpStyle.bEnabled = .Enabled
        tpStyle.bFontBold = .Font.Bold
        tpStyle.bFontItalic = .Font.Italic
        tpStyle.bFontStrikethru = .Font.Strikethrough
        tpStyle.bFontUnderline = .Font.Underline
        tpStyle.bLocked = .Locked
        tpStyle.bVisible = .Visible
        tpStyle.cuFontSize = .Font.Size
        tpStyle.lBackColor = .BackColor
        tpStyle.lForeColor = .ForeColor
        tpStyle.sFontName = .Font.Name
        tpStyle.snHeight = .Height
        tpStyle.snWidth = .Width

        tpStyle.lBackStyle = .BackStyle
        tpStyle.lBorderColor = .BorderColor
        tpStyle.lBorderStyle = .BorderStyle
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : PasteStyleControl - Applies copied style properties to selected controls
'* Created    : 22-03-2023 16:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub PasteStyleControl()
    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub
    Dim objActiveModule As VBComponent
    Dim cnt         As Control
    Set objActiveModule = getActiveModule()
    For Each cnt In objActiveModule.Designer.Selected
        Call setPropertisControl(cnt)
    Next cnt
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : PasteStyleForms - Applies copied style properties to forms
'* Created    : 22-03-2023 16:11
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub PasteStyleForms()
    Dim cnt         As Object
    Set cnt = GetSelectControl(True)
    If cnt Is Nothing Then Exit Sub
    Call setPropertisControl(cnt)
End Sub

'* * * *
'* Sub        : setPropertisControl - Sets the style properties of a control
'* Created    : 22-03-2023 16:11
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):         Description
'*
'* ByVal cnt As Object :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub setPropertisControl(ByVal cnt As Object)
    On Error Resume Next
    With cnt
        .Enabled = tpStyle.bEnabled
        .Font.Bold = tpStyle.bFontBold
        .Font.Italic = tpStyle.bFontItalic
        .Font.Strikethrough = tpStyle.bFontStrikethru
        .Font.Underline = tpStyle.bFontUnderline
        .Locked = tpStyle.bLocked
        .Visible = tpStyle.bVisible
        .Font.Size = tpStyle.cuFontSize
        .BackColor = tpStyle.lBackColor
        .ForeColor = tpStyle.lForeColor
        .Font.Name = tpStyle.sFontName
        .Height = tpStyle.snHeight
        .Width = tpStyle.snWidth

        .BackStyle = tpStyle.lBackStyle
        .BorderColor = tpStyle.lBorderColor
        .BorderStyle = tpStyle.lBorderStyle
    End With
    On Error GoTo 0
End Sub
