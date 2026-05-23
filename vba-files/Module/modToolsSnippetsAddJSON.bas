Attribute VB_Name = "modToolsSnippetsAddJSON"
Option Explicit
Option Private Module

Private Const QUOTE         As String = """"
Private Const QUOTE_COLON   As String = """: """
Private Const COMMA_NEW_LINE As String = ", " & vbNewLine

Private Const indent        As String = "  "
Private Const INDENT_TWO    As String = indent & indent
Private Const INDENT_THREE  As String = INDENT_TWO & indent

Public Sub addJSONFromDirectory()

    Dim sPathRoot   As String
    
    sPathRoot = Environ("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator & m_ROOT_FOLDER
    If Application.Workbooks.Count > 0 Then
        If ActiveWorkbook.Path <> vbNullString Then
            sPathRoot = ActiveWorkbook.Path & Application.PathSeparator & m_ROOT_FOLDER
        End If
    End If
    
    If Not FileHave(sPathRoot, vbDirectory) Then
        Call MsgBox("Not finde path: " & sPathRoot, vbCritical)
        Exit Sub
    End If

    Dim fileDict    As Dictionary
    Set fileDict = getDictFile(sPathRoot)

    Dim sJSON       As String
    Dim sPathJSON   As String
    Dim fileNum     As Integer
    fileNum = FreeFile()

    sPathJSON = sPathRoot & Application.PathSeparator & "SNIPPETS.json"
    If FileHave(sPathJSON, vbNormal) Then Call Kill(sPathJSON)

    sJSON = "{" & vbNewLine & indent & QUOTE & "tableName" & QUOTE_COLON & "TB_SNIPETS" & QUOTE & "," & vbNewLine
    sJSON = sJSON & indent & QUOTE & "version_add_in" & QUOTE_COLON & Version(enVersion, True) & QUOTE & "," & vbNewLine
    sJSON = sJSON & indent & QUOTE & "data" & QUOTE & ": ["

    Open sPathJSON For Output As #fileNum
    Print #fileNum, sJSON
    sJSON = vbNullString

    Dim i           As Long
    Dim iCount      As Long
    Dim arr         As Variant
    arr = fileDict.Items
    iCount = fileDict.Count - 1
    Set fileDict = Nothing
    For i = 0 To iCount
        If i <> iCount Then
            sJSON = INDENT_TWO & "{" & vbNewLine & arr(i) & vbNewLine & INDENT_TWO & "}, "
        Else
            sJSON = INDENT_TWO & "{" & vbNewLine & arr(i) & vbNewLine & INDENT_TWO & "}" & vbNewLine & indent & "]" & vbNewLine & "}"
        End If
        Print #fileNum, sJSON
    Next i
    Close #fileNum

End Sub

Private Function getDictFile(ByVal sPathRoot As String) As Dictionary
    Dim oFSO        As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim rootFolder  As Object
    Set rootFolder = oFSO.GetFolder(sPathRoot)
    Set oFSO = Nothing

    Dim fileDict    As Dictionary
    Set fileDict = New Dictionary
    Call GetAllFiles(rootFolder, fileDict)
    Set rootFolder = Nothing

    Set getDictFile = fileDict
End Function

Private Sub GetAllFiles(currentFolder As Object, ByRef fileDict As Dictionary)

    Dim subFolder   As Object
    Dim file        As Object
    Dim sKey        As String
    Dim sJSON       As String

    For Each file In currentFolder.Files
        With file
            sKey = .ParentFolder.Path
            If Not fileDict.Exists(.ParentFolder.Path) Then
                sJSON = INDENT_THREE & QUOTE & "CODE_GRUP" & QUOTE_COLON & .ParentFolder.ParentFolder.Name & QUOTE & COMMA_NEW_LINE
                sJSON = sJSON & INDENT_THREE & QUOTE & "CODE_SNIPPET" & QUOTE_COLON & .ParentFolder.Name & QUOTE
                fileDict.Add sKey, sJSON
            Else
                sJSON = fileDict(sKey)
            End If

            Select Case .Name
                Case m_FILE_CODE
                    sJSON = sJSON & COMMA_NEW_LINE & INDENT_THREE & QUOTE & "CODE" & QUOTE_COLON & EscapeJSON(loadTextFromTextFile(.Path, m_ENC_WIN1251)) & QUOTE
                Case m_FILE_DESC
                    sJSON = sJSON & COMMA_NEW_LINE & INDENT_THREE & QUOTE & "DISCRIPTION" & QUOTE_COLON & EscapeJSON(loadTextFromTextFile(.Path, m_ENC_UTF8)) & QUOTE
                Case Else
                    Dim sEXP As String
                    sEXP = VBA.LCase$(getFileExeption(.Name))

                    Select Case sEXP
                        Case "cls", "bas"
                            sJSON = sJSON & COMMA_NEW_LINE & INDENT_THREE & QUOTE & .Name & QUOTE_COLON & EscapeJSON(loadTextFromTextFile(.Path, m_ENC_WIN1251)) & QUOTE
                        Case "frm"
                            sJSON = sJSON & COMMA_NEW_LINE & INDENT_THREE & QUOTE & .Name & QUOTE_COLON & EscapeJSON(loadTextFromTextFile(.Path, m_ENC_UTF8)) & QUOTE
                        Case "frx"
                            sJSON = sJSON & COMMA_NEW_LINE & INDENT_THREE & QUOTE & .Name & QUOTE_COLON & EscapeJSON(fileToBase64(.Path)) & QUOTE
                    End Select
                    sEXP = vbNullString
            End Select
            fileDict(sKey) = sJSON
            sJSON = vbNullString
        End With
    Next file

    Set file = Nothing

    For Each subFolder In currentFolder.SubFolders
        If VBA.Left(subFolder.Name, 1) <> "." Then Call GetAllFiles(subFolder, fileDict)
    Next subFolder
End Sub

Private Function getFileExeption(ByVal sNameFile As String) As String
    Dim dotPos      As Long
    dotPos = InStrRev(sNameFile, ".")
    If dotPos > 0 Then getFileExeption = Right$(sNameFile, Len(sNameFile) - dotPos)
End Function

Private Function fileToBase64(ByVal sFilePath As String) As String
    Dim l           As Long
    l = FileLen(sFilePath)
    ReDim byteArr(0 To l) As Byte
    Open sFilePath For Binary As #1
    Get #1, 1, byteArr
    Close #1
    Dim oBase       As Object
    Set oBase = CreateObject("MSXML2.DOMDocument").createElement("b64")
    With oBase
        .DataType = "bin.base64"
        .nodeTypedValue = byteArr
        fileToBase64 = .text
    End With
End Function

Private Function EscapeJSON(ByVal sText As String) As String
    If Len(Trim$(sText)) = 0 Then
        EscapeJSON = vbNullString
        Exit Function
    End If

    Dim sResult     As String
    sResult = sText

    sResult = Replace(sResult, "\", "\\")
    sResult = Replace(sResult, """", "\" & """")
    sResult = Replace(sResult, vbCr, "\r")
    sResult = Replace(sResult, vbLf, "\n")
    sResult = Replace(sResult, vbTab, "\t")
    sResult = Replace(sResult, vbBack, "\b")
    sResult = Replace(sResult, vbFormFeed, "\f")
    sResult = Replace(sResult, vbNullChar, "\u0000")

    sResult = Replace(sResult, ChrW(&H2028), "\u2028")
    sResult = Replace(sResult, ChrW(&H2029), "\u2029")

    EscapeJSON = sResult
End Function

Private Function loadTextFromTextFile(ByVal fileName As String, Optional ByVal encoding As String) As String

    On Error Resume Next
    If Len(Trim(encoding)) = 0 Then encoding = "windows-1251"
    With CreateObject("ADODB.Stream")
        .Type = 2:
        If Len(encoding) Then .Charset = encoding
        .Open
        .LoadFromFile fileName
        loadTextFromTextFile = .ReadText
        .Close
    End With
End Function
