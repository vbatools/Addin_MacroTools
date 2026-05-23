Attribute VB_Name = "modToolsSnippetsAddDirectory"
Option Explicit
Option Private Module

Public Const m_ROOT_FOLDER As String = "ADDIN_MACRO_TOOLS_SNIPPETS"
Public Const m_ENC_WIN1251 As String = "Windows - 1251"
Public Const m_ENC_UTF8 As String = "utf-8"
Public Const m_FILE_CODE As String = "CODE.bas"
Public Const m_FILE_DESC As String = "DISCRIPTION.txt"

Private Const m_PREFIX_CLS As String = "TB_CLS_"
Private Const m_PREFIX_MOD As String = "TB_MOD_"

Public Sub addSnipetsDirectory()
    Dim arrBody()   As Variant
    Dim sPathRoot   As String
    Dim sItemPath   As String
    Dim i           As Long
    Dim lCount      As Long

    On Error GoTo ErrorHandler

    arrBody = getArrayFromTable()
    If Not isArray(arrBody) Then
        MsgBox "Failed to retrieve data from the table.", vbExclamation
        Exit Sub
    End If
 
    sPathRoot = Environ("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator & m_ROOT_FOLDER
    If Application.Workbooks.Count > 0 Then
        If ActiveWorkbook.Path <> vbNullString Then
            sPathRoot = ActiveWorkbook.Path & Application.PathSeparator & m_ROOT_FOLDER
        End If
    End If

    Call EnsureDirectoryExists(sPathRoot)
    sPathRoot = EnsurePathSeparator(sPathRoot)

    lCount = UBound(arrBody, 1)

    For i = 1 To lCount

        sItemPath = sPathRoot & arrBody(i, 1) & Application.PathSeparator
        Call EnsureDirectoryExists(sItemPath)

        sItemPath = sItemPath & arrBody(i, 2) & Application.PathSeparator
        If FileHave(sItemPath, vbDirectory) Then
            Call ClearDirectoryContents(sItemPath)
        Else
            Call MkDir(sItemPath)
        End If

        Call saveTextToFile(CStr(arrBody(i, 3)), sItemPath & m_FILE_CODE, m_ENC_WIN1251)
        Call saveTextToFile(CStr(arrBody(i, 4)), sItemPath & m_FILE_DESC, m_ENC_UTF8)

        If CStr(arrBody(i, 6)) <> vbNullString Then
            Call addSnipetModules(sItemPath, CStr(arrBody(i, 5)), CStr(arrBody(i, 6)))
        End If

        If CStr(arrBody(i, 7)) <> vbNullString Then
            Call addSnipetForms(sItemPath, CStr(arrBody(i, 7)), CStr(arrBody(i, 8)), CStr(arrBody(i, 9)))
        End If
    Next i

    Call addJSONFromDirectory
    Call MsgBox("The snippet database has been exported!", vbInformation)
    Exit Sub

ErrorHandler:
    MsgBox "Error in procedure addSnipetsDir: " & Err.Description, vbCritical
End Sub

Private Function addSnipetModules(ByVal sPath As String, ByVal sClassName As String, ByVal sClassCode As String) As Boolean
    Dim arrNames()  As String
    Dim arrCodes()  As String
    Dim i           As Long
    Dim iMax        As Long

    arrNames = VBA.Split(sClassName, ";")
    arrCodes = VBA.Split(sClassCode, ";")

    If UBound(arrNames) <> UBound(arrCodes) Then
        Debug.Print ">> Error creating module: number of names and code blocks mismatch." & vbCrLf & _
                "Names: " & UBound(arrNames) + 1 & " [" & sClassName & "]" & vbCrLf & _
                "Codes: " & UBound(arrCodes) + 1 & " [" & sClassCode & "]"
        Exit Function
    End If

    iMax = UBound(arrNames)
    For i = 0 To iMax
        Call addSnipetModule(sPath, CStr(arrNames(i)), CStr(arrCodes(i)))
    Next i

    addSnipetModules = True
End Function

Private Sub addSnipetModule(ByVal sPath As String, ByVal sNameModule As String, ByVal sNameShapeOrCode As String)
    Dim sCode       As String
    Dim sFullPath   As String

    If sNameModule = vbNullString Or sNameShapeOrCode = vbNullString Then Exit Sub

    sPath = EnsurePathSeparator(sPath)
    sCode = GetCodeFromShape(sNameShapeOrCode)

    If sCode = vbNullString Then sCode = sNameShapeOrCode

    sFullPath = sPath & sNameModule

    If sNameShapeOrCode Like m_PREFIX_CLS & "*" Then
        Call saveTextToFile(sCode, sFullPath & ".cls", m_ENC_WIN1251)
    ElseIf sNameShapeOrCode Like m_PREFIX_MOD & "*" Then
        Call saveTextToFile(sCode, sFullPath & ".bas", m_ENC_WIN1251)
    Else
        Debug.Print ">> Unidentified module type [" & sNameModule & "].[" & sNameShapeOrCode & "]"
    End If
End Sub

Private Function addSnipetForms(ByVal sPath As String, ByVal sFormaName As String, ByVal sFormaFRM As String, ByVal sFormaFRX As String) As Boolean
    Dim arrName()   As String
    Dim arrFRM()    As String
    Dim arrFRX()    As String
    Dim i           As Long
    Dim iMax        As Long

    arrName = VBA.Split(sFormaName, ";")
    arrFRM = VBA.Split(sFormaFRM, ";")
    arrFRX = VBA.Split(sFormaFRX, ";")

    If UBound(arrName) <> UBound(arrFRM) Or UBound(arrName) <> UBound(arrFRX) Then
        Debug.Print ">> Error creating form: array dimension mismatch." & vbCrLf & _
                "Name: " & UBound(arrName) + 1 & ", FRM: " & UBound(arrFRM) + 1 & ", FRX: " & UBound(arrFRX) + 1
        Exit Function
    End If

    iMax = UBound(arrName)
    For i = 0 To iMax
        Call addSnipetForm(sPath, CStr(arrName(i)), CStr(arrFRM(i)), CStr(arrFRX(i)))
    Next i

    addSnipetForms = True
End Function

Private Sub addSnipetForm(ByVal sPath As String, ByVal sFormaName As String, ByVal sFormaFRM As String, ByVal sFormaFRX As String)
    Dim sContentFRM As String
    Dim sContentFRX As String
    Dim sTargetPath As String

    sContentFRM = GetCodeFromShape(sFormaFRM)
    sContentFRX = GetCodeFromShape(sFormaFRX)

    If sContentFRM <> vbNullString And sContentFRX <> vbNullString Then
        sTargetPath = EnsurePathSeparator(sPath)
        Call base64ToFile(sContentFRM, sTargetPath & sFormaName & ".frm")
        Call base64ToFile(sContentFRX, sTargetPath & sFormaName & ".frx")
    End If
End Sub

Private Function GetCodeFromShape(ByRef sNameShape As String) As String
    On Error GoTo ErrorHandler

    If Not shSettings Is Nothing Then
        GetCodeFromShape = shSettings.Shapes(sNameShape).TextFrame2.TextRange.text
    End If

    Exit Function

ErrorHandler:
    Debug.Print ">> Shape with code not found [" & sNameShape & "]"
    Err.Clear
End Function

Private Sub base64ToFile(ByVal sHashBase64 As String, ByVal sFilePath As String)
    Dim byteArr()   As Byte
    Dim oBase       As Object
    Dim lFileNum    As Integer

    On Error GoTo ErrorHandler

    Set oBase = CreateObject("MSXML2.DOMDocument").createElement("b64")
    With oBase
        .DataType = "bin.base64"
        .text = sHashBase64
        byteArr = .nodeTypedValue
    End With

    If Len(Dir(sFilePath)) > 0 Then Kill sFilePath

    lFileNum = FreeFile
    Open sFilePath For Binary Access Write As #lFileNum
    Put #lFileNum, 1, byteArr
    Close #lFileNum

    Exit Sub

ErrorHandler:
    If lFileNum > 0 Then Close #lFileNum
    MsgBox "Error writing Base64 file: " & sFilePath & vbCrLf & Err.Description, vbCritical
End Sub

Private Function EnsurePathSeparator(ByVal sPath As String) As String
    If Right(sPath, 1) <> Application.PathSeparator Then
        EnsurePathSeparator = sPath & Application.PathSeparator
    Else
        EnsurePathSeparator = sPath
    End If
End Function

Private Sub EnsureDirectoryExists(ByVal sPath As String)
    If FileHave(sPath, vbDirectory) Then Exit Sub
    MkDir sPath
End Sub

Private Sub ClearDirectoryContents(ByVal sPath As String)
    Dim sFile       As String

    On Error Resume Next

    sPath = EnsurePathSeparator(sPath)
    sFile = Dir(sPath & "*.*")

    Do While sFile <> ""
        Kill sPath & sFile
        sFile = Dir()
    Loop

    On Error GoTo 0
End Sub

Public Function getArrayFromTable() As Variant
    getArrayFromTable = getListSnipets().DataBodyRange.Value2
End Function

Private Function getListSnipets() As ListObject
    Set getListSnipets = shSettings.ListObjects("TB_SNIPETS")
End Function

Private Function saveTextToFile(ByVal txt As String, ByVal fileName As String, Optional ByVal encoding As String = "windows-1251") As Boolean

    On Error Resume Next: Err.Clear
    Dim FSO         As Object
    Dim ts          As Object
    Select Case encoding
        Case "windows-1251", "", "ansi"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(fileName, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing

        Case "utf-16", "utf-16LE"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(fileName, True, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing

        Case "utf-8noBOM"
            Dim binaryStream As Object
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = "utf-8": .Open
                .WriteText txt$

                Set binaryStream = CreateObject("ADODB.Stream")
                binaryStream.Type = 1: binaryStream.mode = 3: binaryStream.Open
                .Position = 3: .CopyTo binaryStream
                .flush: .Close
                binaryStream.SaveToFile fileName$, 2
                binaryStream.Close
            End With

        Case Else
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = encoding$: .Open
                .WriteText txt$
                .SaveToFile fileName$, 2
                .Close
            End With
    End Select
    saveTextToFile = Err = 0: DoEvents
End Function



