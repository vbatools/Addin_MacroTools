Attribute VB_Name = "modToolsRegExp"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modRegExp - regular expression testing
'* Created    : 22-04-2020 23:02
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RegExpStart - start checking regular expression
'* Created    : 23-04-2020 00:03
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RegExpStart()
    Dim sSTR        As String
    Dim sPattern    As String
    Dim sReplace    As String
    Dim sMsgErr     As String
    Dim bGloba1     As Boolean
    Dim bIgnoreCase As Boolean
    Dim bMultiLine  As Boolean

    Application.ScreenUpdating = False
    With ActiveSheet
        sSTR = VBA.Trim$(.Cells(11, 3).value)
        sPattern = VBA.Trim$(.Cells(2, 3).value)
        sReplace = VBA.Trim$(.Cells(24, 3).value)
        bGloba1 = VBA.CBool(.Cells(7, 3).value)
        bIgnoreCase = VBA.CBool(.Cells(8, 3).value)
        bMultiLine = VBA.CBool(.Cells(9, 3).value)
    End With

    If sPattern = vbNullString Then sMsgErr = "Regular expression not specified!" & vbNewLine
    If sSTR = vbNullString Then sMsgErr = sMsgErr & "Source text not specified!"

    Call RegExpClearCells
    If sMsgErr <> vbNullString Then
        Call MsgBox(sMsgErr, vbCritical, "Search for Matches:")
        Exit Sub
    End If
    'reset formatting
    With ActiveSheet.Cells(11, 3).Font
        .ColorIndex = xlAutomatic
        .Underline = xlUnderlineStyleNone
    End With
    With ActiveSheet.Cells(26, 3).Font
        .ColorIndex = xlAutomatic
        .Underline = xlUnderlineStyleNone
    End With

    Call RegExpEnjoyReplace(sSTR, sPattern, sReplace, bGloba1, bIgnoreCase, bMultiLine)
    Call RegExpGetMatches(sSTR, sPattern, sReplace, bGloba1, bIgnoreCase, bMultiLine)
    Application.ScreenUpdating = True
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RegExpGetMatches - start highlighting regex matches with color
'* Created    : 23-03-2023 10:04
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByVal sSTR As String                    :
'* ByVal sPattern As String                :
'* ByVal sReplace As String                :
'* Optional bGloba1 As Boolean = True      :
'* Optional bIgnoreCase As Boolean = False :
'* Optional bMultiLine As Boolean = False  :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RegExpGetMatches(ByVal sSTR As String, ByVal sPattern As String, ByVal sReplace As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False)

    Dim objMatches  As Object
    Dim itemMatch   As Object
    Dim lRow        As Long
    Dim iFerstChr   As Integer
    Dim i           As Integer

    lRow = 2
    i = 1
    iFerstChr = 0

    With ActiveSheet
        Set objMatches = RegExpExecuteCollection(sSTR, sPattern, bGloba1, bIgnoreCase, bMultiLine)
        If objMatches Is Nothing Then
            Call MsgBox("No matches found!", vbInformation + vbOKOnly, "Search for Matches:")
            .Range("M:P").EntireColumn.AutoFit
        ElseIf objMatches.Count = 0 Then
            Call MsgBox("No matches found!", vbInformation + vbOKOnly, "Search for Matches:")
            .Range("M:P").EntireColumn.AutoFit
        Else
            For Each itemMatch In objMatches
                With itemMatch
                    ActiveSheet.Cells(lRow, 13).value = lRow - 1
                    ActiveSheet.Cells(lRow, 14).value = .FirstIndex
                    ActiveSheet.Cells(lRow, 15).value = .Length
                    ActiveSheet.Cells(lRow, 16).value = .value
                End With

                With ActiveSheet.Cells(11, 3).Characters(Start:=itemMatch.FirstIndex + 1, Length:=itemMatch.Length).Font
                    .Color = -16776961
                    .Underline = xlUnderlineStyleSingle
                End With

                sReplace = RegExpFindReplace(sReplace, "\$[1-9]{1}", vbNullString, True, False, True)
                iFerstChr = VBA.InStr(iFerstChr + 1, ActiveSheet.Cells(26, 3).value, sReplace)
                If iFerstChr > 0 And sReplace <> vbNullString Then
                    With ActiveSheet.Cells(26, 3).Characters(Start:=iFerstChr, Length:=VBA.Len(sReplace)).Font
                        .Color = -16776961
                        .Underline = xlUnderlineStyleSingle
                    End With
                End If
                lRow = lRow + 1
            Next itemMatch
            .Range("M:P").EntireColumn.AutoFit
        End If
    End With
    'Clean up object references
    Set itemMatch = Nothing
    Set objMatches = Nothing
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RegExpEnjoyReplace - perform regex replacement
'* Created    : 22-04-2020 23:24
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RegExpEnjoyReplace(ByVal sSTR As String, ByVal sPattern As String, ByVal sReplace As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False)
    With ActiveSheet
        .Cells(26, 3).value = RegExpFindReplace(sSTR, sPattern, sReplace, bGloba1, bIgnoreCase, bMultiLine)
    End With
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RegExpFindReplace - find and replace using regex
'* Created    : 22-04-2020 23:07
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* sStr As String                          : source string
'* sPattern As String                      : search pattern
'* sReplace As String                      : replacement string
'* Optional bGloba1 As Boolean = True      : FALSE - check until first match, TRUE - check entire text
'* Optional bIgnoreCase As Boolean = False : FALSE - case sensitive, TRUE - case insensitive
'* Optional bMultiline As Boolean = False  : FALSE - single-line mode, TRUE - multi-line mode
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function RegExpFindReplace(ByVal sSTR As String, ByVal sPattern As String, ByVal sReplace As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False) As String
    RegExpFindReplace = sSTR
    If Not sPattern Like vbNullString Then
        Dim regExp  As Object
        Set regExp = CreateObject("VBScript.RegExp")
        With regExp
            .Global = bGloba1
            .IgnoreCase = bIgnoreCase
            .MultiLine = bMultiLine
            .pattern = sPattern
        End With

        On Error Resume Next
        RegExpFindReplace = regExp.Replace(sSTR, sReplace)
        Set regExp = Nothing
    End If
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RegExpExecuteCollection - collection of regex matches
'* Created    : 22-04-2020 22:59
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* sStr As String                          : source string
'* Pattern As String                       : search pattern
'* Optional bGloba1 As Boolean = True      : FALSE - check until first match, TRUE - check entire text
'* Optional bIgnoreCase As Boolean = False : FALSE - case sensitive, TRUE - case insensitive
'* Optional bMultiline As Boolean = False  : FALSE - single-line mode, TRUE - multi-line mode
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function RegExpExecuteCollection(ByVal sSTR As String, ByVal sPattern As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False) As Object
    Set RegExpExecuteCollection = Nothing
    If Not sPattern Like vbNullString Then
        Dim regExp  As Object
        Set regExp = CreateObject("VBScript.RegExp")
        With regExp
            .Global = bGloba1
            .IgnoreCase = bIgnoreCase
            .MultiLine = bMultiLine
            .pattern = sPattern
        End With

        On Error Resume Next
        'Get the collection of matches
        Set RegExpExecuteCollection = regExp.Execute(sSTR)
        Set regExp = Nothing
    End If
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RegExpClearCellsAll - clear the form before running regex
'* Created    : 22-04-2020 23:11
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RegExpClearCellsAll()
    With ActiveSheet
        .Range("C24:K24").ClearContents
    End With
    Call RegExpClearCells
    Call RegExpClearCellsPattern
    Call RegExpClearCellsStr
End Sub
Private Sub RegExpClearCells()
    With ActiveSheet
        .Range(.Cells(26, 3), .Cells(37, 11)).ClearContents
        .Range(.Cells(2, 13), .Cells(.Cells(.Rows.Count, 13).End(xlUp).Row + 1, 16)).ClearContents
    End With
End Sub
Public Sub RegExpClearCellsPattern()
    With ActiveSheet
        .Range(.Cells(2, 3), .Cells(3, 11)).ClearContents
    End With
End Sub
Public Sub RegExpClearCellsStr()
    With ActiveSheet
        .Range(.Cells(11, 3), .Cells(22, 11)).ClearContents
    End With
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddSheetTestRegExp - create sheet for pattern testing
'* Created    : 25-04-2020 21:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddSheetTestRegExp()
    Dim shName      As String
    shName = shRegExp.Name
    'create sheet in active workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(shName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets(shName).Copy After:=ActiveWorkbook.ActiveSheet
    With ActiveWorkbook.Sheets(shName)
        .Visible = True
        .Activate
    End With
End Sub
