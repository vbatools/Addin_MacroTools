Attribute VB_Name = "modUFControlsStyleCopyPaste"
Option Explicit
Option Private Module

Public tpStyle      As ProperControlStyle
Type ProperControlStyle
    bCopy           As Boolean
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

Public Sub CopyStyleControl()
    Call CopyStyle(False)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CopyStyleControl - copy control style formatting
'* Created    : 22-03-2023 16:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CopyStyle(bUserForm As Boolean)
    Dim cnt         As Object
    Set cnt = GetSelectControl(bUserForm)
    If cnt Is Nothing Then Exit Sub

    If TypeName(cnt) = "Controls" Then Set cnt = cnt.Item(0)

    'set default values
    tpStyle.bCopy = True
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
'* Sub        : PasteStyleControl - paste control style to selected controls
'* Created    : 22-03-2023 16:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub PasteStyleControl()
    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub

    Dim selectedControls As Object
    Dim selectedControl As control

    Set selectedControls = GetSelectControl()
    If selectedControls Is Nothing Then Exit Sub
    Select Case TypeName(selectedControls)
             Case "Controls"
            Dim cnt As control
            For Each cnt In selectedControls
                Call setPropertisControl(cnt)
            Next cnt
        Case Else
            Call setPropertisControl(selectedControls)
    End Select
    Set selectedControls = Nothing
    Set selectedControl = Nothing
End Sub

Public Sub CopyStyleForms()
    Call CopyStyle(True)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : PasteStyleForms - paste control style to form
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
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : setPropertisControl - copy control style formatting
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
    If Not tpStyle.bCopy Then
        Debug.Print ">> Style was not copied"
        Exit Sub
    End If
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
